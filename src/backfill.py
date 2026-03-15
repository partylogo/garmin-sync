"""Backfill - 手動補歷史資料

用法:
    python -m src.backfill --type sleep --days 90
    python -m src.backfill --type activity --days 90
    python -m src.backfill --type laps --days 90
    python -m src.backfill --type both --days 90
    python -m src.backfill --type all --days 90
"""

import argparse
import sys
from datetime import datetime

from .config import Config
from .garmin_client import GarminClient, AuthenticationError
from .sheets_client import SheetsClient, SheetsError
from .sync_garmin import sync_sleep_data, sync_activity_data, sync_laps_data, log


def main() -> int:
    """主程式入口"""
    parser = argparse.ArgumentParser(
        description="手動補 Garmin 歷史資料到 Google Sheets"
    )
    parser.add_argument(
        "--type",
        choices=["sleep", "activity", "laps", "both", "all"],
        required=True,
        help="要補的資料類型：sleep（睡眠）、activity（活動）、laps（分圈）、both（睡眠+活動）、all（全部）",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=90,
        help="要補的天數（預設 90 天）",
    )

    args = parser.parse_args()

    log("=" * 50)
    log(f"Garmin 歷史資料補齊 - {args.type} - 最近 {args.days} 天")
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

        sleep_count = 0
        activity_count = 0
        laps_count = 0

        # 同步睡眠資料
        if args.type in ["sleep", "both", "all"]:
            sleep_count = sync_sleep_data(
                garmin, sheets, args.days, config.timezone
            )

        # 同步活動資料
        if args.type in ["activity", "both", "all"]:
            activity_count = sync_activity_data(garmin, sheets, args.days)

        # 同步分圈資料
        if args.type in ["laps", "all"]:
            laps_count = sync_laps_data(garmin, sheets, args.days)

        # API 模式：flush buffer 送出資料
        if hasattr(sheets, 'flush'):
            flush_result = sheets.flush()
            log(f"API 模式批次寫入完成: {flush_result}")

        # 總結
        log("=" * 50)
        if args.type == "sleep":
            log(f"補齊完成！睡眠: +{sleep_count}")
        elif args.type == "activity":
            log(f"補齊完成！活動: +{activity_count}")
        elif args.type == "laps":
            log(f"補齊完成！分圈: +{laps_count}")
        elif args.type == "both":
            log(f"補齊完成！睡眠: +{sleep_count}, 活動: +{activity_count}")
        else:  # all
            log(f"補齊完成！睡眠: +{sleep_count}, 活動: +{activity_count}, 分圈: +{laps_count}")
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
