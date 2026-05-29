# 🌿 溫室帳目 Line Bot

## 部署步驟（Railway，約 20 分鐘）

### 1. 準備 Line Bot
1. 前往 https://developers.line.biz → 建立 Provider → 建立 Messaging API Channel
2. 記下 **Channel Secret** 和 **Channel Access Token**
3. Webhook URL 先空著，等部署完再填

### 2. 準備 Google Sheets API
1. 前往 https://console.cloud.google.com → 建立新專案
2. 啟用 **Google Sheets API** 和 **Google Drive API**
3. 建立憑證 → Service Account → 下載 JSON 金鑰
4. 開啟你的試算表 → 分享給 Service Account 的 Email（編輯權限）
5. 從試算表網址取得 SPREADSHEET_ID：
   https://docs.google.com/spreadsheets/d/**這段就是ID**/edit

### 3. 部署到 Railway
1. 前往 https://railway.app → 用 GitHub 登入
2. New Project → Deploy from GitHub repo（上傳此資料夾）
3. 設定環境變數（Variables）：
   - LINE_CHANNEL_SECRET
   - LINE_CHANNEL_ACCESS_TOKEN
   - SPREADSHEET_ID
   - GOOGLE_CREDENTIALS_JSON（把整份 JSON 貼進去）
4. 部署完成後，複製 Railway 給的網址

### 4. 設定 Webhook
1. 回到 Line Developers Console
2. Messaging API → Webhook settings
3. Webhook URL 填入：https://你的railway網址/callback
4. 打開 **Use webhook**
5. 按 Verify 測試

### 5. 測試
在 Line 輸入：
```
收入 50000 珍妮500顆 李淵男
```
看看試算表有沒有新增一行 ✅

## 輸入格式

| 格式 | 範例 |
|------|------|
| 基本收入 | `收入 50000 珍妮500顆 李淵男` |
| 基本支出 | `支出 18000 大陸運費` |
| 標記未付 | `收入 32000 爆米花1000顆 美琪 未付` |
| 簡短格式 | `+50000 侏儒黃月1000顆 吳政翰` |
| 簡短支出 | `-4200 淘寶悶箱` |

## 自動分類規則
| 關鍵字 | 分類 |
|--------|------|
| 顆/盒/植物/種苗 | 種苗銷售 |
| 運費/空軍/黑貓 | 運費 |
| 水苔/土/紙箱 | 材料耗材 |
| 電費/水電 | 水電費 |
| 換人民幣/換RMB | 換匯 |
| 燈/機器/噴霧 | 設備 |
| 薪水/薪資 | 薪資 |
