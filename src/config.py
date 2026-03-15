"""Config Module - 設定與欄位對應"""

import os
import json
from datetime import datetime
from typing import Optional
import pytz
from dotenv import load_dotenv

# 自動載入 .env 檔案
load_dotenv()


class Config:
    """應用程式設定，從環境變數讀取"""

    def __init__(self):
        self.garmin_email = os.environ.get("GARMIN_EMAIL", "")
        self.garmin_password = os.environ.get("GARMIN_PASSWORD", "")
        self.google_sheet_id = os.environ.get("GOOGLE_SHEET_ID", "")
        self.sync_days = int(os.environ.get("SYNC_DAYS", "7"))
        self.timezone = os.environ.get("TIMEZONE", "Asia/Taipei")

        # API 模式設定（選填）
        self.api_url = os.environ.get("API_URL", "")
        self.api_key = os.environ.get("API_KEY", "")

        # Google credentials 是 JSON 字串（直連模式需要）
        creds_str = os.environ.get("GOOGLE_CREDENTIALS", "{}")
        try:
            self.google_credentials = json.loads(creds_str)
        except json.JSONDecodeError:
            self.google_credentials = {}

    @property
    def use_api_mode(self) -> bool:
        """是否使用 API 模式（API_URL 和 API_KEY 都有值時啟用）"""
        return bool(self.api_url and self.api_key)

    def validate(self) -> list[str]:
        """驗證必要設定是否存在，回傳缺少的設定名稱"""
        missing = []
        if not self.garmin_email:
            missing.append("GARMIN_EMAIL")
        if not self.garmin_password:
            missing.append("GARMIN_PASSWORD")
        if not self.google_sheet_id:
            missing.append("GOOGLE_SHEET_ID")
        if not self.use_api_mode and not self.google_credentials:
            missing.append("GOOGLE_CREDENTIALS (或設定 API_URL + API_KEY 使用 API 模式)")
        return missing


# Sleep 欄位對應（13欄，與原 Sheet 一致）
SLEEP_FIELDS = [
    "睡眠分數 4 週",  # 日期
    "分數",
    "靜止心率",
    "身體能量指數",
    "脈搏血氧",
    "呼吸",
    "皮膚溫度變化",
    "HRV狀態",
    "品質",
    "持續時間",
    "睡眠需求",
    "就寢時間",
    "起床時間",
]

# Laps 欄位對應（9欄）
LAPS_FIELDS = [
    "日期",
    "活動名稱",
    "圈數",
    "距離",
    "時間",
    "配速",
    "平均心率",
    "最大心率",
    "步頻",
]

# Activity 欄位對應（33欄，與原 Sheet 一致）
ACTIVITY_FIELDS = [
    "活動類型",
    "日期",
    "標題",
    "距離",
    "卡路里",
    "時間",
    "平均心率",
    "最大心率",
    "有氧訓練效果",
    "平均步頻",
    "最高步頻",
    "平均配速",
    "最佳配速",
    "總爬升",
    "總下降",
    "平均步幅",
    "平均移動效率",
    "平均垂直振幅",
    "平均觸地時間",
    "平均坡度校正配速",
    "Normalized Power® (NP®)",
    "Training Stress Score®",
    "平均功率",
    "最大功率",
    "步數",
    "身體能量指數消耗量",
    "減壓",
    "最佳圈耗時",
    "圈數",
    "移動時間",
    "經過時間",
    "最低海拔",
    "最高海拔",
]


def seconds_to_chinese_duration(seconds) -> str:
    """將秒數轉換為 'X時 Y分鐘' 格式

    Args:
        seconds: 秒數，可為 None 或 int/float

    Returns:
        格式化的時間字串，無資料時回傳 '--'
    """
    if seconds is None:
        return "--"
    seconds = int(seconds)
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    return f"{hours}時 {minutes}分鐘"


def seconds_to_duration(seconds) -> str:
    """將秒數轉換為 'H:MM:SS' 格式

    Args:
        seconds: 秒數，可為 None 或 int/float

    Returns:
        格式化的時間字串，無資料時回傳 '--'
    """
    if seconds is None:
        return "--"
    seconds = int(seconds)
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    return f"{hours}:{minutes:02d}:{secs:02d}"


def meters_to_km(meters: Optional[float]) -> str:
    """將公尺轉換為公里

    Args:
        meters: 公尺數，可為 None

    Returns:
        格式化的公里數，無資料時回傳 '--'
    """
    if meters is None:
        return "--"
    return f"{meters / 1000:.2f}"


def pace_to_string(pace_seconds_per_km) -> str:
    """將配速（秒/公里）轉換為 'M:SS' 格式

    Args:
        pace_seconds_per_km: 每公里秒數，可為 None 或 int/float

    Returns:
        格式化的配速字串，無資料時回傳 '--'
    """
    if pace_seconds_per_km is None or pace_seconds_per_km <= 0:
        return "--"
    pace_seconds_per_km = int(pace_seconds_per_km)
    minutes = pace_seconds_per_km // 60
    seconds = pace_seconds_per_km % 60
    return f"{minutes}:{seconds:02d}"


def timestamp_to_local_time(
    timestamp_ms: Optional[int], timezone_str: str = "Asia/Taipei"
) -> str:
    """將 GMT 時間戳轉換為本地時間 '上午/下午 HH:MM' 格式

    Args:
        timestamp_ms: 毫秒時間戳，可為 None
        timezone_str: 時區字串，預設為台北

    Returns:
        格式化的本地時間字串，無資料時回傳 '--'
    """
    if timestamp_ms is None:
        return "--"

    try:
        dt = datetime.fromtimestamp(timestamp_ms / 1000, tz=pytz.UTC)
        local_tz = pytz.timezone(timezone_str)
        local_dt = dt.astimezone(local_tz)

        period = "上午" if local_dt.hour < 12 else "下午"
        hour = local_dt.hour % 12 or 12
        return f"{period} {hour}:{local_dt.minute:02d}"
    except Exception:
        return "--"


def safe_get(data: Optional[dict], *keys, default=None):
    """安全地從巢狀字典取值

    Args:
        data: 字典資料
        *keys: 要取的 key 路徑
        default: 預設值

    Returns:
        取得的值或預設值
    """
    if data is None:
        return default
    result = data
    for key in keys:
        if isinstance(result, dict):
            result = result.get(key)
        else:
            return default
        if result is None:
            return default
    return result
