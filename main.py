import os
import re
import json
from datetime import datetime, date
from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi,
    ReplyMessageRequest, TextMessage
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

# ── 設定 ────────────────────────────────────────────────────────
LINE_CHANNEL_SECRET      = os.environ["LINE_CHANNEL_SECRET"]
LINE_CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
SPREADSHEET_ID           = os.environ["SPREADSHEET_ID"]
GOOGLE_CREDS_JSON        = os.environ["GOOGLE_CREDENTIALS_JSON"]  # 整份 JSON 字串

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler       = WebhookHandler(LINE_CHANNEL_SECRET)

# ── Google Sheets 連線 ──────────────────────────────────────────
def get_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds  = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

# ── 解析訊息 ────────────────────────────────────────────────────
CATEGORIES = {
    # 種苗相關
    "顆": "種苗銷售", "盒": "種苗銷售", "種苗": "種苗銷售",
    "植物": "種苗銷售", "珍妮": "種苗銷售", "侏儒": "種苗銷售",
    "斑葉": "種苗銷售", "黃月": "種苗銷售", "神鉅": "種苗銷售",
    "妙蛙": "種苗銷售", "火鶴": "種苗銷售", "鹿角": "種苗銷售",
    "爆米花": "種苗銷售",
    # 運費
    "運費": "運費", "空軍": "運費", "黑貓": "運費", "郵局": "運費",
    # 耗材
    "水苔": "材料耗材", "土": "材料耗材", "紙箱": "材料耗材",
    "膠膜": "材料耗材", "悶箱": "材料耗材", "棉花": "材料耗材",
    "垃圾袋": "材料耗材", "盆": "材料耗材",
    # 水電
    "電費": "水電費", "水電": "水電費",
    # 固定支出
    "租金": "租金", "薪水": "薪資", "薪資": "薪資",
    # 換匯
    "換人民幣": "換匯", "換rmb": "換匯", "匯款": "換匯",
    # 設備
    "燈": "設備", "機器": "設備", "噴霧": "設備", "冰箱": "設備",
    # 農藥
    "農藥": "農藥肥料", "肥料": "農藥肥料",
}

def guess_category(item_text: str) -> str:
    text = item_text.lower()
    for keyword, cat in CATEGORIES.items():
        if keyword.lower() in text:
            return cat
    return "其他"

def parse_message(text: str) -> dict | None:
    """
    支援格式：
      收入 50000 珍妮500顆 李淵男
      支出 18000 大陸運費
      收入 32000 爆米花1000顆 美琪 未付
      +50000 珍妮500顆 李淵男
      -18000 大陸運費
    """
    text = text.strip()

    # 正規化類型標記
    if text.startswith("+"):
        text = "收入 " + text[1:]
    elif text.startswith("-"):
        text = "支出 " + text[1:]

    # 主要 regex：類型 金額 品項 [客戶] [未付]
    pattern = r"^(收入|支出)\s+([\d,]+)\s+(.+?)(?:\s+([\u4e00-\u9fff\w\s]+?))?(?:\s+(未付|已付|已收|部分收))?$"
    m = re.match(pattern, text, re.IGNORECASE)
    if not m:
        return None

    tx_type  = m.group(1)
    amount   = int(m.group(2).replace(",", ""))
    item     = m.group(3).strip()
    customer = (m.group(4) or "").strip()
    status_raw = m.group(5)

    # 付款狀態預設
    if status_raw:
        status = status_raw
    elif tx_type == "收入":
        status = "未付"    # 預設收入為待收，請確認後改成已收
    else:
        status = "已付"

    # 如果沒有客戶但品項裡有空格，嘗試把最後一段當客戶
    if not customer and " " in item:
        parts = item.rsplit(" ", 1)
        item     = parts[0].strip()
        customer = parts[1].strip()

    today    = date.today()
    month    = f"{today.month}月"
    category = guess_category(item)

    return {
        "date":     today.strftime("%Y/%m/%d"),
        "type":     tx_type,
        "amount":   amount,
        "item":     item,
        "customer": customer,
        "status":   status,
        "pay_date": today.strftime("%Y/%m/%d") if status in ("已收","已付") else "",
        "note":     "",
        "category": category,
        "month":    month,
        "rmb":      "",
        "raw":      text,
    }

# ── 寫入 Google Sheets ──────────────────────────────────────────
def append_to_sheet(data: dict) -> int:
    book  = get_sheet()
    sheet = book.worksheet("📋 交易記錄")

    row = [
        data["date"],
        data["type"],
        data["amount"],
        data["item"],
        data["customer"],
        data["status"],
        data["pay_date"],
        data["note"],
        data["category"],
        data["month"],
        data["rmb"],
        data["raw"],
    ]
    sheet.append_row(row, value_input_option="USER_ENTERED")

    # 取得剛寫入的行號（最後一行）
    all_rows = sheet.get_all_values()
    return len(all_rows)

# ── 格式化回覆訊息 ──────────────────────────────────────────────
def format_reply(data: dict, row_num: int) -> str:
    icon = "💰" if data["type"] == "收入" else "💸"
    status_icon = {
        "已收": "✅", "已付": "✅",
        "未付": "⏳", "部分收": "⚠️"
    }.get(data["status"], "")

    lines = [
        f"{icon} 已記錄到第 {row_num} 行",
        f"{'─' * 20}",
        f"日期｜{data['date']}",
        f"類型｜{data['type']}",
        f"金額｜NT$ {data['amount']:,}",
        f"品項｜{data['item']}",
    ]
    if data["customer"]:
        lines.append(f"客戶｜{data['customer']}")
    lines.append(f"狀態｜{status_icon} {data['status']}")
    lines.append(f"分類｜{data['category']}")
    lines.append(f"{'─' * 20}")
    lines.append("✏️ 如需修改請直接更新試算表")
    return "\n".join(lines)

HELP_TEXT = """🌿 溫室帳目機器人

📥 輸入格式：
收入 金額 品項 客戶
支出 金額 品項

📌 範例：
收入 50000 珍妮500顆 李淵男
支出 18000 大陸運費
收入 32000 爆米花1000顆 美琪 未付
+50000 侏儒黃月1000顆 吳政翰
-4200 淘寶悶箱

💡 付款狀態（選填）：
已收 / 已付 / 未付 / 部分收

輸入「說明」再次查看此訊息"""

# ── Flask 路由 ──────────────────────────────────────────────────
@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body      = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event: MessageEvent):
    user_text = event.message.text.strip()
    reply_token = event.reply_token

    with ApiClient(configuration) as api_client:
        line_api = MessagingApi(api_client)

        # 說明指令
        if user_text in ("說明", "help", "Help", "HELP", "?", "？"):
            reply = HELP_TEXT

        else:
            parsed = parse_message(user_text)
            if parsed is None:
                reply = (
                    "❌ 無法解析此訊息\n\n"
                    "請用以下格式輸入：\n"
                    "收入 50000 珍妮500顆 李淵男\n"
                    "支出 18000 大陸運費\n\n"
                    "輸入「說明」查看完整格式"
                )
            else:
                try:
                    row_num = append_to_sheet(parsed)
                    reply   = format_reply(parsed, row_num)
                except Exception as e:
                    reply = f"⚠️ 寫入試算表時發生錯誤\n{str(e)}\n\n請稍後再試"

        line_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=reply)]
            )
        )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
