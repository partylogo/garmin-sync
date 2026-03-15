"""API Sheets Client - 透過中間 API 操作 Google Sheets

與 SheetsClient 相同的介面（duck typing），底層改為 HTTP 呼叫。
insert_*_row() 會先 buffer 在本地，呼叫 flush() 時一次送出。
"""

import requests
from .sheets_client import SheetsError


class ApiSheetsClient:
    """透過 API 操作 Google Sheets 的客戶端"""

    def __init__(self, api_url: str, api_key: str, sheet_id: str):
        """初始化 API 客戶端

        Args:
            api_url: API 的 base URL（例如 https://your-project.vercel.app）
            api_key: API Key
            sheet_id: Google Sheet 的 ID
        """
        self.api_url = api_url.rstrip("/")
        self.sheet_id = sheet_id
        self._session = requests.Session()
        self._session.headers.update({
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        })

        # 本地 buffer
        self._sleep_buffer: list[list] = []
        self._activity_buffer: list[list] = []
        self._laps_buffer: list[list] = []

        # Cache
        self._laps_sheet_exists: bool | None = None

    def _post(self, endpoint: str, data: dict) -> dict:
        """發送 POST 請求到 API

        Args:
            endpoint: API endpoint（例如 /api/existing）
            data: 請求 body

        Returns:
            回應的 JSON dict

        Raises:
            SheetsError: API 回傳錯誤
        """
        url = f"{self.api_url}{endpoint}"
        try:
            resp = self._session.post(url, json=data, timeout=300)
        except requests.RequestException as e:
            raise SheetsError(f"API 連線錯誤: {e}")

        if resp.status_code != 200:
            try:
                error_data = resp.json()
                msg = error_data.get("error", resp.text)
            except Exception:
                msg = resp.text
            raise SheetsError(f"API 錯誤 ({resp.status_code}): {msg}")

        return resp.json()

    def _get_existing_ids(self, data_type: str) -> set[str]:
        """取得指定分頁已有的去重 ID"""
        result = self._post("/api/existing", {
            "sheet_id": self.sheet_id,
            "data_type": data_type,
        })
        return set(result.get("ids", []))

    def get_existing_sleep_dates(self) -> set[str]:
        """取得所有已存在的睡眠日期"""
        return self._get_existing_ids("sleep")

    def get_existing_activity_ids(self) -> set[str]:
        """取得所有已存在的活動 ID"""
        return self._get_existing_ids("activity")

    def get_existing_lap_ids(self) -> set[str]:
        """取得所有已存在的分圈 ID"""
        return self._get_existing_ids("laps")

    def insert_sleep_row(self, data: list) -> bool:
        """Buffer 睡眠資料（flush 時才送出）"""
        self._sleep_buffer.append(data)
        return True

    def insert_activity_row(self, data: list) -> bool:
        """Buffer 活動資料（flush 時才送出）"""
        self._activity_buffer.append(data)
        return True

    def insert_lap_row(self, data: list) -> bool:
        """Buffer 分圈資料（flush 時才送出）"""
        self._laps_buffer.append(data)
        return True

    def has_laps_sheet(self) -> bool:
        """檢查是否有 laps 分頁"""
        if self._laps_sheet_exists is None:
            result = self._post("/api/existing", {
                "sheet_id": self.sheet_id,
                "data_type": "laps",
            })
            self._laps_sheet_exists = result.get("sheet_exists", False)
        return self._laps_sheet_exists

    def flush(self) -> dict:
        """把 buffer 的 rows 一次送出到 API

        Returns:
            dict: {"sleep": N, "activity": N, "laps": N} 各插入幾筆
        """
        result = {"sleep": 0, "activity": 0, "laps": 0}

        for data_type, buffer in [
            ("sleep", self._sleep_buffer),
            ("activity", self._activity_buffer),
            ("laps", self._laps_buffer),
        ]:
            if not buffer:
                continue

            resp = self._post("/api/sync", {
                "sheet_id": self.sheet_id,
                "data_type": data_type,
                "rows": buffer,
            })
            result[data_type] = resp.get("inserted", 0)

        # 清空 buffer
        self._sleep_buffer.clear()
        self._activity_buffer.clear()
        self._laps_buffer.clear()

        return result
