"""Garmin Client - Garmin Connect API 封裝"""

from typing import Optional
from garminconnect import Garmin


class AuthenticationError(Exception):
    """Garmin 認證失敗"""

    pass


class GarminClient:
    """Garmin Connect API 客戶端"""

    def __init__(self, email: str, password: str):
        """初始化 Garmin 客戶端

        Args:
            email: Garmin 帳號 Email
            password: Garmin 密碼
        """
        self.email = email
        self.password = password
        self._client: Optional[Garmin] = None

    def login(self) -> bool:
        """登入 Garmin Connect

        Returns:
            True 如果登入成功

        Raises:
            AuthenticationError: 認證失敗時
        """
        try:
            self._client = Garmin(self.email, self.password)
            self._client.login()
            return True
        except Exception as e:
            raise AuthenticationError(f"Garmin 認證失敗: {e}")

    def get_sleep_data(self, date: str) -> Optional[dict]:
        """取得指定日期的睡眠資料

        Args:
            date: 日期字串 (YYYY-MM-DD)

        Returns:
            睡眠資料字典，無資料時回傳 None
        """
        if not self._client:
            raise AuthenticationError("尚未登入")

        try:
            data = self._client.get_sleep_data(date)
            # 檢查是否有有效資料
            if data and data.get("dailySleepDTO"):
                return data
            return None
        except Exception:
            return None

    def get_activities(self, start_date: str, end_date: str) -> list[dict]:
        """取得指定日期範圍內的活動

        Args:
            start_date: 開始日期 (YYYY-MM-DD)
            end_date: 結束日期 (YYYY-MM-DD)

        Returns:
            活動列表
        """
        if not self._client:
            raise AuthenticationError("尚未登入")

        try:
            activities = self._client.get_activities_by_date(start_date, end_date)
            return activities if activities else []
        except Exception:
            return []

    def get_heart_rates(self, date: str) -> Optional[dict]:
        """取得指定日期的心率資料

        Args:
            date: 日期字串 (YYYY-MM-DD)

        Returns:
            心率資料字典，無資料時回傳 None
        """
        if not self._client:
            raise AuthenticationError("尚未登入")

        try:
            data = self._client.get_heart_rates(date)
            return data if data else None
        except Exception:
            return None

    def get_hrv_data(self, date: str) -> Optional[dict]:
        """取得指定日期的 HRV 資料

        Args:
            date: 日期字串 (YYYY-MM-DD)

        Returns:
            HRV 資料字典，無資料時回傳 None
        """
        if not self._client:
            raise AuthenticationError("尚未登入")

        try:
            data = self._client.get_hrv_data(date)
            return data if data else None
        except Exception:
            return None

    def get_spo2_data(self, date: str) -> Optional[dict]:
        """取得指定日期的 SpO2 資料

        Args:
            date: 日期字串 (YYYY-MM-DD)

        Returns:
            SpO2 資料字典，無資料時回傳 None
        """
        if not self._client:
            raise AuthenticationError("尚未登入")

        try:
            data = self._client.get_spo2_data(date)
            return data if data else None
        except Exception:
            return None

    def get_activity_splits(self, activity_id: str) -> Optional[dict]:
        """取得活動的分圈資料

        Args:
            activity_id: 活動 ID

        Returns:
            分圈資料字典，無資料時回傳 None
        """
        if not self._client:
            raise AuthenticationError("尚未登入")

        try:
            data = self._client.get_activity_splits(activity_id)
            return data if data else None
        except Exception:
            return None
