# Garmin to Google Sheets Sync

自動從 Garmin Connect 抓取健康資料，同步到 Google Sheets。

## 功能

- 每天自動同步睡眠資料（分數、心率、HRV、血氧等）
- 每天自動同步活動資料（跑步、騎車、健行等）
- 智慧去重：不會產生重複資料
- 新資料插入在最上方

## 快速開始

### 1. 使用此範本建立你的 Repo

點擊右上角的 **Use this template** → **Create a new repository**。

### 2. 建立 Google Sheet

1. 前往 [Google Sheets](https://sheets.google.com/) 建立新的試算表
2. 將試算表**共用**給以下 Email（權限設為「編輯者」），取消勾選「通知使用者」：

   ```
   garmin-sync@fude-admin.iam.gserviceaccount.com
   ```

3. 從瀏覽器網址列複製你的 Sheet ID：
   ```
   https://docs.google.com/spreadsheets/d/[這一段就是SHEET_ID]/edit
   ```

> 分頁和標題列會在第一次同步時自動建立，不需要手動設定。

### 3. 一鍵設定 GitHub Secrets

打開 Setup 頁面，按照步驟完成設定：

1. 前往 **[Setup 頁面](https://garmin-sheets-api.vercel.app/setup)**
2. 點擊「連結 GitHub 帳號」並授權
3. 選擇你剛建立的 Repository
4. 填入 Garmin 帳密和 Google Sheet ID
5. 點擊「設定 Secrets」— 完成！

> 你的 Garmin 密碼會在瀏覽器端加密後才送出，不會經過我們的 server。

### 4. 啟用 GitHub Actions

1. 前往 **Actions** 分頁
2. 如果看到警告訊息，點擊 **I understand my workflows, go ahead and enable them**
3. 點擊左側的 **Sync Garmin Data**
4. 點擊 **Run workflow** → 綠色的 **Run workflow** 按鈕
5. 等待 1-2 分鐘，綠色勾勾 ✓ 表示成功
6. 去 Google Sheet 確認資料已寫入

## 同步規則

- **首次執行**：自動同步最近 **90 天**的資料，並自動建立分頁和標題列
- **之後每天**：台灣時間 12:00 自動同步最近 **7 天**
- **去重**：已存在的資料自動跳過，不會重複寫入

可以在 `.github/workflows/sync.yml` 中修改 cron 排程：

```yaml
schedule:
  - cron: '0 4 * * *'  # UTC 時間，台灣時間要 +8
```

## 常見問題

### Q: GitHub Actions 執行失敗？

1. 點擊失敗的 workflow run 查看詳細錯誤
2. 常見原因：
   - **Garmin 認證失敗**：確認帳號密碼正確
   - **Sheet 沒共用**：確認已將 Sheet 共用給 Service Account Email
   - **Sheet ID 錯誤**：確認 Secret 中的 ID 正確

### Q: 重複執行會產生重複資料嗎？

不會。程式會自動檢查已有的資料，跳過重複的。

### Q: Garmin 開啟了雙重驗證 (MFA)？

目前可能無法處理 MFA。建議暫時關閉 Garmin 帳號的 MFA。

## 技術細節

- [python-garminconnect](https://github.com/cyberjunky/python-garminconnect) — Garmin API
- [gspread](https://docs.gspread.org/) — Google Sheets API
- GitHub Actions — 定時執行

## License

MIT

