"""Sheets Client - Google Sheets API 封裝"""

from typing import Optional
import gspread
from google.oauth2.service_account import Credentials


class SheetsError(Exception):
    """Google Sheets 操作錯誤"""

    pass


class SheetsClient:
    """Google Sheets API 客戶端"""

    # Google Sheets API 需要的 scopes
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    def __init__(self, credentials: dict, sheet_id: str):
        """初始化 Sheets 客戶端

        Args:
            credentials: Service Account JSON 認證資訊（dict 格式）
            sheet_id: Google Sheet 的 ID
        """
        self.sheet_id = sheet_id
        self._client: Optional[gspread.Client] = None
        self._spreadsheet: Optional[gspread.Spreadsheet] = None
        self._sleep_sheet: Optional[gspread.Worksheet] = None
        self._activity_sheet: Optional[gspread.Worksheet] = None
        self._laps_sheet: Optional[gspread.Worksheet] = None

        try:
            creds = Credentials.from_service_account_info(
                credentials, scopes=self.SCOPES
            )
            self._client = gspread.authorize(creds)
            self._spreadsheet = self._client.open_by_key(sheet_id)
            self._sleep_sheet = self._spreadsheet.worksheet("sleep")
            self._activity_sheet = self._spreadsheet.worksheet("activity")
            # laps sheet 是可選的
            try:
                self._laps_sheet = self._spreadsheet.worksheet("laps")
            except gspread.WorksheetNotFound:
                self._laps_sheet = None
        except Exception as e:
            raise SheetsError(f"無法連接 Google Sheets: {e}")

    def get_existing_sleep_dates(self) -> set[str]:
        """取得所有已存在的睡眠日期

        Returns:
            日期字串的 set（跳過標題列）
        """
        if not self._sleep_sheet:
            raise SheetsError("Sleep worksheet 未初始化")

        try:
            # 取得第一欄（日期欄）的所有值，跳過標題列
            dates = self._sleep_sheet.col_values(1)[1:]
            return set(d for d in dates if d)  # 過濾空值
        except Exception as e:
            raise SheetsError(f"無法讀取睡眠日期: {e}")

    def get_existing_activity_ids(self) -> set[str]:
        """取得所有已存在的活動 ID

        Activity Sheet 的最後一欄用來存放 activityId（隱藏欄位）

        Returns:
            activityId 字串的 set（跳過標題列）
        """
        if not self._activity_sheet:
            raise SheetsError("Activity worksheet 未初始化")

        try:
            # 取得第二欄（日期時間欄）作為 ID（因為日期時間是唯一的）
            # 或者如果有 activityId 欄位，取得那一欄
            # 這裡我們用日期時間作為唯一識別
            ids = self._activity_sheet.col_values(2)[1:]
            return set(i for i in ids if i)  # 過濾空值
        except Exception as e:
            raise SheetsError(f"無法讀取活動 ID: {e}")

    def insert_sleep_row(self, data: list) -> bool:
        """插入睡眠資料到第 2 列（新資料在最上面）

        Args:
            data: 要插入的資料列表，順序需與 SLEEP_FIELDS 一致

        Returns:
            True 如果插入成功
        """
        if not self._sleep_sheet:
            raise SheetsError("Sleep worksheet 未初始化")

        try:
            self._sleep_sheet.insert_row(data, index=2)
            return True
        except Exception as e:
            raise SheetsError(f"無法插入睡眠資料: {e}")

    def insert_activity_row(self, data: list) -> bool:
        """插入活動資料到第 2 列（新資料在最上面）

        Args:
            data: 要插入的資料列表，順序需與 ACTIVITY_FIELDS 一致

        Returns:
            True 如果插入成功
        """
        if not self._activity_sheet:
            raise SheetsError("Activity worksheet 未初始化")

        try:
            self._activity_sheet.insert_row(data, index=2)
            return True
        except Exception as e:
            raise SheetsError(f"無法插入活動資料: {e}")

    def has_laps_sheet(self) -> bool:
        """檢查是否有 laps 分頁"""
        return self._laps_sheet is not None

    def get_existing_lap_ids(self) -> set[str]:
        """取得所有已存在的分圈 ID（用 日期+活動名稱+圈數 作為唯一識別）

        Returns:
            lap ID 字串的 set（跳過標題列）
        """
        if not self._laps_sheet:
            return set()

        try:
            # 取得前三欄（日期、活動名稱、圈數）組合成 ID
            all_values = self._laps_sheet.get_all_values()[1:]  # 跳過標題
            ids = set()
            for row in all_values:
                if len(row) >= 3 and row[0] and row[1] and row[2]:
                    ids.add(f"{row[0]}|{row[1]}|{row[2]}")
            return ids
        except Exception as e:
            raise SheetsError(f"無法讀取分圈 ID: {e}")

    def insert_lap_row(self, data: list) -> bool:
        """插入分圈資料到第 2 列（新資料在最上面）

        Args:
            data: 要插入的資料列表，順序需與 LAPS_FIELDS 一致

        Returns:
            True 如果插入成功
        """
        if not self._laps_sheet:
            raise SheetsError("Laps worksheet 未初始化，請先建立 'laps' 分頁")

        try:
            self._laps_sheet.insert_row(data, index=2)
            return True
        except Exception as e:
            raise SheetsError(f"無法插入分圈資料: {e}")
