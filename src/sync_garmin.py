"""Sync Garmin - 主程式入口

從 Garmin Connect 抓取睡眠和活動資料，同步到 Google Sheets。
"""

import sys
import time
from datetime import datetime, timedelta
from typing import Optional

from .config import (
    Config,
    seconds_to_chinese_duration,
    seconds_to_duration,
    meters_to_km,
    pace_to_string,
    timestamp_to_local_time,
    safe_get,
)
from .garmin_client import GarminClient, AuthenticationError
from .api_sheets_client import SheetsError

try:
    from .sheets_client import SheetsClient
except ImportError:
    SheetsClient = None  # template repo 不需要直連模式


def log(message: str) -> None:
    """輸出日誌訊息"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


def get_date_range(days: int) -> list[str]:
    """取得最近 N 天的日期列表

    Args:
        days: 天數

    Returns:
        日期字串列表 (YYYY-MM-DD)，從最近的日期開始
    """
    today = datetime.now().date()
    dates = []
    for i in range(days):
        date = today - timedelta(days=i)
        dates.append(date.strftime("%Y-%m-%d"))
    return dates


def extract_sleep_data(
    garmin: GarminClient, date: str, timezone: str
) -> Optional[list]:
    """從 Garmin 抓取並轉換睡眠資料

    Args:
        garmin: Garmin 客戶端
        date: 日期字串 (YYYY-MM-DD)
        timezone: 時區字串

    Returns:
        睡眠資料列表（符合 SLEEP_FIELDS 順序，13欄），無資料時回傳 None
    """
    sleep_data = garmin.get_sleep_data(date)
    if not sleep_data:
        return None

    daily_sleep = safe_get(sleep_data, "dailySleepDTO")
    if not daily_sleep:
        return None

    # 取得輔助資料
    hr_data = garmin.get_heart_rates(date)
    hrv_data = garmin.get_hrv_data(date)
    spo2_data = garmin.get_spo2_data(date)

    # 取得 sleepNeed（可能是 dict 或數字）
    sleep_need = safe_get(daily_sleep, "sleepNeed")
    if isinstance(sleep_need, dict):
        # sleepNeed.actual 是分鐘數，轉換為秒數
        actual_minutes = safe_get(sleep_need, "actual")
        sleep_need = actual_minutes * 60 if actual_minutes else None

    # 取得 sleepScores.overall.value
    sleep_score = safe_get(daily_sleep, "sleepScores", "overall", "value")
    if sleep_score is None:
        # 嘗試其他路徑
        sleep_score = safe_get(daily_sleep, "sleepScores", "totalScore")

    # 組裝資料（順序需與 SLEEP_FIELDS 一致，共 13 欄）
    row = [
        # 1. 睡眠分數 4 週（日期）
        safe_get(daily_sleep, "calendarDate", default=date),
        # 2. 分數
        sleep_score if sleep_score is not None else "--",
        # 3. 靜止心率
        safe_get(hr_data, "restingHeartRate", default="--"),
        # 4. 身體能量指數 - bodyBatteryChange 在 sleep_data 頂層
        safe_get(sleep_data, "bodyBatteryChange", default="--"),
        # 5. 脈搏血氧
        safe_get(spo2_data, "averageSpO2", default=safe_get(daily_sleep, "averageSpO2Value", default="--")),
        # 6. 呼吸
        safe_get(daily_sleep, "averageRespirationValue", default="--"),
        # 7. 皮膚溫度變化
        safe_get(sleep_data, "avgSkinTempDeviationC", default="--"),
        # 8. HRV狀態
        safe_get(sleep_data, "avgOvernightHrv", default=safe_get(hrv_data, "hrvSummary", "lastNightAvg", default="--")),
        # 9. 品質
        "--",
        # 10. 持續時間
        seconds_to_chinese_duration(safe_get(daily_sleep, "sleepTimeSeconds")),
        # 11. 睡眠需求
        seconds_to_chinese_duration(sleep_need),
        # 12. 就寢時間
        timestamp_to_local_time(
            safe_get(daily_sleep, "sleepStartTimestampGMT"), timezone
        ),
        # 13. 起床時間
        timestamp_to_local_time(
            safe_get(daily_sleep, "sleepEndTimestampGMT"), timezone
        ),
    ]

    return row


def format_int_value(value) -> str:
    """將數值轉換為整數字串，None 時回傳 '--'"""
    if value is None:
        return "--"
    try:
        return str(int(value))
    except (ValueError, TypeError):
        return "--"


def format_float_value(value, decimals: int = 2) -> str:
    """將數值轉換為浮點數字串，None 時回傳 '--'"""
    if value is None:
        return "--"
    try:
        return f"{float(value):.{decimals}f}"
    except (ValueError, TypeError):
        return "--"


def format_elevation(value) -> str:
    """將海拔數值格式化，None 時回傳 '--'"""
    if value is None:
        return "--"
    try:
        return str(int(value))
    except (ValueError, TypeError):
        return "--"


def extract_activity_data(activity: dict) -> list:
    """從 Garmin 活動資料轉換為 Sheet 格式

    Args:
        activity: Garmin 活動資料字典

    Returns:
        活動資料列表（符合 ACTIVITY_FIELDS 順序，共 33 欄）
    """
    # 計算配速（秒/公里）
    distance = safe_get(activity, "distance")
    duration = safe_get(activity, "duration")
    avg_pace = None
    if distance and duration and distance > 0:
        avg_pace = duration / (distance / 1000)

    # 計算最佳配速（從 maxSpeed 公尺/秒 轉換）
    max_speed = safe_get(activity, "maxSpeed")
    best_pace = None
    if max_speed and max_speed > 0:
        best_pace = 1000 / max_speed  # 秒/公里

    # 計算坡度校正配速（從 avgGradeAdjustedSpeed 公尺/秒 轉換）
    grade_adj_speed = safe_get(activity, "avgGradeAdjustedSpeed")
    grade_adj_pace = None
    if grade_adj_speed and grade_adj_speed > 0:
        grade_adj_pace = 1000 / grade_adj_speed  # 秒/公里

    # 組裝資料（順序需與 ACTIVITY_FIELDS 一致，共 33 欄）
    row = [
        # 1. 活動類型
        safe_get(activity, "activityType", "typeKey", default="--"),
        # 2. 日期
        safe_get(activity, "startTimeLocal", default="--"),
        # 3. 標題
        safe_get(activity, "activityName", default="--"),
        # 4. 距離
        meters_to_km(distance),
        # 5. 卡路里
        format_int_value(safe_get(activity, "calories")),
        # 6. 時間
        seconds_to_duration(duration),
        # 7. 平均心率
        format_int_value(safe_get(activity, "averageHR")),
        # 8. 最大心率
        format_int_value(safe_get(activity, "maxHR")),
        # 9. 有氧訓練效果
        format_float_value(safe_get(activity, "aerobicTrainingEffect"), 1),
        # 10. 平均步頻
        format_int_value(safe_get(activity, "averageRunningCadenceInStepsPerMinute")),
        # 11. 最高步頻
        format_int_value(safe_get(activity, "maxRunningCadenceInStepsPerMinute")),
        # 12. 平均配速
        pace_to_string(avg_pace),
        # 13. 最佳配速（從 maxSpeed 轉換）
        pace_to_string(best_pace),
        # 14. 總爬升
        format_elevation(safe_get(activity, "elevationGain")),
        # 15. 總下降
        format_elevation(safe_get(activity, "elevationLoss")),
        # 16. 平均步幅（從公分轉為公尺）
        format_float_value(safe_get(activity, "avgStrideLength") / 100 if safe_get(activity, "avgStrideLength") else None, 2),
        # 17. 平均移動效率
        format_float_value(safe_get(activity, "avgVerticalRatio"), 1),
        # 18. 平均垂直振幅
        format_float_value(safe_get(activity, "avgVerticalOscillation"), 1),
        # 19. 平均觸地時間
        format_int_value(safe_get(activity, "avgGroundContactTime")),
        # 20. 平均坡度校正配速（從 avgGradeAdjustedSpeed 轉換）
        pace_to_string(grade_adj_pace),
        # 21. Normalized Power® (NP®)
        format_int_value(safe_get(activity, "normPower")),
        # 22. Training Stress Score®
        format_int_value(safe_get(activity, "trainingStressScore")),
        # 23. 平均功率
        format_int_value(safe_get(activity, "avgPower")),
        # 24. 最大功率
        format_int_value(safe_get(activity, "maxPower")),
        # 25. 步數
        format_int_value(safe_get(activity, "steps")),
        # 26. 身體能量指數消耗量（differenceBodyBattery）
        format_int_value(safe_get(activity, "differenceBodyBattery")),
        # 27. 減壓
        safe_get(activity, "decompression", default="否"),
        # 28. 最佳圈耗時（minActivityLapDuration）
        seconds_to_duration(safe_get(activity, "minActivityLapDuration")),
        # 29. 圈數
        format_int_value(safe_get(activity, "lapCount")),
        # 30. 移動時間
        seconds_to_duration(safe_get(activity, "movingDuration")),
        # 31. 經過時間
        seconds_to_duration(safe_get(activity, "elapsedDuration")),
        # 32. 最低海拔
        format_elevation(safe_get(activity, "minElevation")),
        # 33. 最高海拔
        format_elevation(safe_get(activity, "maxElevation")),
    ]

    return row


def extract_lap_data(activity: dict, lap: dict) -> list:
    """從 Garmin 分圈資料轉換為 Sheet 格式

    Args:
        activity: Garmin 活動資料字典
        lap: 單圈資料字典

    Returns:
        分圈資料列表（符合 LAPS_FIELDS 順序，共 9 欄）
    """
    # 取得活動日期（只取日期部分）
    start_time = safe_get(activity, "startTimeLocal", default="")
    date = start_time.split(" ")[0] if start_time else "--"

    # 計算配速（秒/公里）
    distance = safe_get(lap, "distance")  # 公尺
    duration = safe_get(lap, "duration")  # 秒
    pace = None
    if distance and duration and distance > 0:
        pace = duration / (distance / 1000)  # 秒/公里

    # 組裝資料（順序需與 LAPS_FIELDS 一致，共 9 欄）
    row = [
        # 1. 日期
        date,
        # 2. 活動名稱
        safe_get(activity, "activityName", default="--"),
        # 3. 圈數
        safe_get(lap, "lapIndex", default="--"),
        # 4. 距離 (km)
        f"{distance / 1000:.2f}" if distance else "--",
        # 5. 時間
        seconds_to_duration(duration),
        # 6. 配速
        pace_to_string(pace),
        # 7. 平均心率
        format_int_value(safe_get(lap, "averageHR")),
        # 8. 最大心率
        format_int_value(safe_get(lap, "maxHR")),
        # 9. 步頻
        format_int_value(safe_get(lap, "averageRunCadence")),
    ]

    return row


def sync_laps_data(
    garmin: GarminClient, sheets: SheetsClient, days: int
) -> int:
    """同步分圈資料

    Args:
        garmin: Garmin 客戶端
        sheets: Sheets 客戶端
        days: 同步天數

    Returns:
        新增的資料筆數
    """
    log(f"開始同步分圈資料（最近 {days} 天）...")

    # 檢查是否有 laps 分頁
    if not sheets.has_laps_sheet():
        log("  警告：找不到 'laps' 分頁，跳過分圈同步")
        log("  請在 Google Sheet 建立 'laps' 分頁，並在第一列加入標題：")
        log("  日期 | 活動名稱 | 圈數 | 距離 | 時間 | 配速 | 平均心率 | 最大心率 | 步頻")
        return 0

    # 取得現有分圈 ID
    existing_ids = sheets.get_existing_lap_ids()
    log(f"Sheet 中已有 {len(existing_ids)} 筆分圈資料")

    # 取得日期範圍
    dates = get_date_range(days)
    start_date = dates[-1]  # 最早的日期
    end_date = dates[0]  # 最近的日期

    # 取得活動（反轉順序，從舊到新，這樣最新的會在最上面）
    activities = garmin.get_activities(start_date, end_date)
    activities.reverse()  # 從舊到新
    log(f"Garmin 回傳 {len(activities)} 筆活動")

    count = 0
    for activity in activities:
        activity_id = safe_get(activity, "activityId")
        activity_name = safe_get(activity, "activityName", default="Unknown")
        start_time = safe_get(activity, "startTimeLocal", default="")
        date = start_time.split(" ")[0] if start_time else ""

        # 取得分圈資料
        splits_data = garmin.get_activity_splits(activity_id)
        if not splits_data:
            continue

        laps = safe_get(splits_data, "lapDTOs", default=[])
        if not laps:
            continue

        # 反轉圈數順序，讓最新的圈在最上面
        laps_reversed = list(reversed(laps))

        activity_lap_count = 0
        for lap in laps_reversed:
            lap_index = safe_get(lap, "lapIndex", default="")
            lap_id = f"{date}|{activity_name}|{lap_index}"

            if lap_id in existing_ids:
                continue

            lap_row = extract_lap_data(activity, lap)
            sheets.insert_lap_row(lap_row)
            count += 1
            activity_lap_count += 1
            if not hasattr(sheets, 'flush'):
                time.sleep(1.5)  # 直連模式才需要等，API 模式是 buffer

        if activity_lap_count > 0:
            log(f"  新增 {activity_name} 的 {activity_lap_count} 圈資料")
        elif laps:
            log(f"  跳過 {activity_name}（{len(laps)} 圈已存在）")

    log(f"分圈資料同步完成，新增 {count} 筆")
    return count


def sync_sleep_data(
    garmin: GarminClient, sheets: SheetsClient, days: int, timezone: str
) -> int:
    """同步睡眠資料

    Args:
        garmin: Garmin 客戶端
        sheets: Sheets 客戶端
        days: 同步天數
        timezone: 時區

    Returns:
        新增的資料筆數
    """
    log(f"開始同步睡眠資料（最近 {days} 天）...")

    # 取得現有日期
    existing_dates = sheets.get_existing_sleep_dates()
    log(f"Sheet 中已有 {len(existing_dates)} 筆睡眠資料")

    # 取得日期範圍（從最舊的開始處理，這樣最新的會在最上面）
    dates = get_date_range(days)
    dates.reverse()  # 從舊到新
    count = 0

    for date in dates:
        if date in existing_dates:
            log(f"  跳過 {date}（已存在）")
            continue

        sleep_row = extract_sleep_data(garmin, date, timezone)
        if sleep_row:
            sheets.insert_sleep_row(sleep_row)
            log(f"  新增 {date} 睡眠資料")
            count += 1
            if not hasattr(sheets, 'flush'):
                time.sleep(1.5)  # 直連模式才需要等，API 模式是 buffer
        else:
            log(f"  跳過 {date}（無資料）")

    log(f"睡眠資料同步完成，新增 {count} 筆")
    return count


def sync_activity_data(
    garmin: GarminClient, sheets: SheetsClient, days: int
) -> int:
    """同步活動資料

    Args:
        garmin: Garmin 客戶端
        sheets: Sheets 客戶端
        days: 同步天數

    Returns:
        新增的資料筆數
    """
    log(f"開始同步活動資料（最近 {days} 天）...")

    # 取得現有活動（用日期時間作為 ID）
    existing_ids = sheets.get_existing_activity_ids()
    log(f"Sheet 中已有 {len(existing_ids)} 筆活動資料")

    # 取得日期範圍
    dates = get_date_range(days)
    start_date = dates[-1]  # 最早的日期
    end_date = dates[0]  # 最近的日期

    # 取得活動（反轉順序，從舊到新，這樣最新的會在最上面）
    activities = garmin.get_activities(start_date, end_date)
    activities.reverse()  # 從舊到新
    log(f"Garmin 回傳 {len(activities)} 筆活動")

    count = 0
    for activity in activities:
        # 用 startTimeLocal 作為唯一識別
        activity_id = safe_get(activity, "startTimeLocal", default="")
        if activity_id in existing_ids:
            log(f"  跳過活動（已存在）: {activity_id}")
            continue

        activity_row = extract_activity_data(activity)
        sheets.insert_activity_row(activity_row)
        log(f"  新增活動: {safe_get(activity, 'activityName', default='Unknown')}")
        count += 1
        if not hasattr(sheets, 'flush'):
            time.sleep(1.5)  # 直連模式才需要等，API 模式是 buffer

    log(f"活動資料同步完成，新增 {count} 筆")
    return count


def main() -> int:
    """主程式入口

    Returns:
        Exit code: 0 成功，1 失敗
    """
    log("=" * 50)
    log("Garmin to Google Sheets 同步開始")
    log("=" * 50)

    # 讀取設定
    config = Config()

    # 驗證設定
    missing = config.validate()
    if missing:
        log(f"錯誤：缺少必要環境變數: {', '.join(missing)}")
        return 1

    try:
        # 登入 Garmin
        log("正在登入 Garmin Connect...")
        garmin = GarminClient(config.garmin_email, config.garmin_password)
        garmin.login()
        log("Garmin 登入成功")

        # 連接 Google Sheets
        if config.use_api_mode:
            from .api_sheets_client import ApiSheetsClient
            log(f"使用 API 模式: {config.api_url}")
            sheets = ApiSheetsClient(config.api_url, config.api_key, config.google_sheet_id)
        else:
            log("正在連接 Google Sheets（直連模式）...")
            sheets = SheetsClient(config.google_credentials, config.google_sheet_id)
        log("Google Sheets 連接成功")

        # API 模式：先確保分頁存在，並判斷是否首次執行
        is_first_run = False
        if hasattr(sheets, '_post'):
            log("正在初始化 Sheet 分頁...")
            try:
                init_result = sheets._post("/api/init", {
                    "sheet_id": config.google_sheet_id,
                    "sheets": ["sleep", "activity", "laps"],
                })
                created = init_result.get("created", [])
                if created:
                    log(f"  已建立分頁: {', '.join(created)}")
                    if "sleep" in created:
                        is_first_run = True
                else:
                    log("  所有分頁已存在")
            except Exception as e:
                log(f"  初始化分頁失敗: {e}（繼續嘗試同步）")

        sync_days = config.sync_days
        if is_first_run:
            sync_days = 90
            log(f"首次執行，自動擴大同步範圍至 {sync_days} 天")

        # 同步睡眠資料
        sleep_count = sync_sleep_data(
            garmin, sheets, sync_days, config.timezone
        )

        # 同步活動資料
        activity_count = sync_activity_data(garmin, sheets, sync_days)

        # 同步分圈資料
        laps_count = sync_laps_data(garmin, sheets, sync_days)

        # API 模式：flush buffer 送出資料
        if hasattr(sheets, 'flush'):
            flush_result = sheets.flush()
            log(f"API 模式批次寫入完成: {flush_result}")

        # 總結
        log("=" * 50)
        log(f"同步完成！睡眠: +{sleep_count}, 活動: +{activity_count}, 分圈: +{laps_count}")
        log("=" * 50)
        return 0

    except AuthenticationError as e:
        log(f"Garmin 認證錯誤: {e}")
        return 1
    except SheetsError as e:
        log(f"Google Sheets 錯誤: {e}")
        return 1
    except Exception as e:
        log(f"未預期的錯誤: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
