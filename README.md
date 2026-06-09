# 🌿 溫室帳目 Line Bot

透過 LINE 訊息自動記帳，即時同步至 Google Sheets。

---

## 部署步驟

### 1. 準備 LINE Bot
1. 前往 https://developers.line.biz → 建立 Provider → 建立 Messaging API Channel
2. 記下 **Channel Secret** 和 **Channel Access Token**
3. LINE Official Account Manager → 回應設定 → 關閉「聊天」、開啟「Webhook」

### 2. 準備 Google Sheets API
1. 前往 https://console.cloud.google.com → 建立新專案
2. 啟用 **Google Sheets API** 和 **Google Drive API**
3. 建立憑證 → Service Account → 下載 JSON 金鑰
4. 開啟試算表 → 共用給 Service Account Email（編輯者權限）
5. 從試算表網址取得 SPREADSHEET_ID：
   `https://docs.google.com/spreadsheets/d/`**這段就是 ID**`/edit`

### 3. 部署到 Railway
1. 前往 https://railway.app → 用 GitHub 登入
2. New Project → GitHub Repository → 選 linebot → Deploy Now
3. Variables 填入四個環境變數：

| 變數名稱 | 說明 |
|---------|------|
| `LINE_CHANNEL_SECRET` | LINE Developers → Basic settings |
| `LINE_CHANNEL_ACCESS_TOKEN` | LINE Developers → Messaging API |
| `SPREADSHEET_ID` | 試算表網址中間那段 |
| `GOOGLE_CREDENTIALS_JSON` | 服務帳戶 JSON 檔案完整內容 |

4. Settings → Domains → Generate Domain → 複製網址

### 4. 設定 Webhook
1. LINE Developers → Messaging API → Webhook settings
2. 填入 `https://你的Railway網址/callback`
3. 開啟 **Use webhook** → 按 **Verify** → 看到 Success ✅

---

## 輸入格式

### 新增交易
```
收入 金額 品項 客戶
支出 金額 品項
```

### 選填參數

| 參數 | 格式 | 範例 |
|------|------|------|
| 指定日期 | 日期放最前面 | `5/15 收入 50000 珍妮500顆 李淵男` |
| 付款狀態 | 放最後 | `收入 32000 爆米花1000顆 美琪 未付` |
| 進價（算毛利） | `進價XX` | `收入 50000 珍妮500顆 李淵男 進價30` |
| 換匯金額 | `人民幣XXXX` | `支出 18114 換人民幣4000 黃浩` |
| 付款期限 | `期限M/D` | `收入 50000 珍妮500顆 李淵男 未付 期限7/15` |
| 部分收款 | `部分收 金額` | `收入 50000 珍妮500顆 李淵男 部分收 30000` |
| 簡短格式 | `+` / `-` | `+50000 珍妮500顆 李淵男` |

### 更新付款狀態
```
更新 行號 已收
更新 行號 已付
更新 行號 部分收 金額
```

### 查看說明
```
說明
```

---

## 自動功能

- **自動分類**：根據品項關鍵字判斷（種苗銷售、運費、換匯、差旅費等）
- **自動管道**：根據客戶/品項判斷（直接客戶、蝦皮、大陸出口、韓國出口等）
- **毛利計算**：輸入進價後自動計算毛利與毛利率
- **匯率計算**：輸入人民幣金額後自動計算匯率
- **應收帳款同步**：收入標記「未付」或「部分收」時，自動寫入應收帳款分頁
- **逾期標色**：應收帳款逾期 30 / 60 / 90 天分別標示黃 / 橘 / 紅色
- **收款天數**：更新為已收時，自動計算從交易日到收款的天數

---

## 常見問題

**Verify 出現 400 Bad Request**  
→ Railway Variables 的 `LINE_CHANNEL_SECRET` 前後有空格，刪除後重新部署

**試算表沒有更新**  
→ 確認試算表已分享給 Service Account Email（編輯者）  
→ 確認工作表名稱為 `📋 交易記錄`（含 emoji）

**應收帳款沒有同步**  
→ 付款狀態需為「未付」或「部分收」才會觸發  
→ 確認有 `💰 應收帳款` 工作表且名稱完全一致
