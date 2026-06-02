import os
import re
import json
import logging
from collections import deque
from datetime import datetime, timedelta, timezone
from functools import lru_cache
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
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
logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"))
logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════
# 設定
# ══════════════════════════════════════════════════════════════
LINE_CHANNEL_SECRET       = os.environ["LINE_CHANNEL_SECRET"]
LINE_CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
SPREADSHEET_ID            = os.environ["SPREADSHEET_ID"]
GOOGLE_CREDS_JSON         = os.environ["GOOGLE_CREDENTIALS_JSON"]
APP_TIMEZONE              = os.environ.get("APP_TIMEZONE", "Asia/Taipei")

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler       = WebhookHandler(LINE_CHANNEL_SECRET)

# 工作表名稱（若未來改名只需改這裡）
SHEET_TRANSACTIONS = "📋 交易記錄"
SHEET_RECEIVABLES  = "💰 應收帳款"

TX_INCOME = "收入"
TX_EXPENSE = "支出"
STATUS_UNPAID = "未付"
STATUS_PAID = "已付"
STATUS_COLLECTED = "已收"
STATUS_PARTIAL = "部分收"
STATUS_VALUES = (STATUS_UNPAID, STATUS_PAID, STATUS_COLLECTED, STATUS_PARTIAL)
PAID_STATUSES = {STATUS_PAID, STATUS_COLLECTED}

RECENT_MESSAGE_ID_LIMIT = 500
_recent_message_ids = deque(maxlen=RECENT_MESSAGE_ID_LIMIT)
_recent_message_id_set: set[str] = set()

# ══════════════════════════════════════════════════════════════
# Google Sheets 連線
# ══════════════════════════════════════════════════════════════
@lru_cache(maxsize=1)
def get_workbook():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds  = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDS_JSON), scopes=scopes
    )
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def get_app_timezone():
    try:
        return ZoneInfo(APP_TIMEZONE)
    except ZoneInfoNotFoundError:
        logger.warning("Timezone %s not found, falling back to UTC+8", APP_TIMEZONE)
        return timezone(timedelta(hours=8))

def today_tw():
    return datetime.now(get_app_timezone()).date()

def to_int(number_text: str) -> int:
    return int(number_text.replace(",", ""))

def parse_append_row_number(response: dict, fallback_sheet=None) -> int:
    updated_range = response.get("updates", {}).get("updatedRange", "")
    m = re.search(r"![A-Z]+(\d+)(?::[A-Z]+\d+)?$", updated_range)
    if m:
        return int(m.group(1))
    return len(fallback_sheet.col_values(1)) if fallback_sheet is not None else 0

def has_seen_message(message_id: str | None) -> bool:
    return bool(message_id and message_id in _recent_message_id_set)

def remember_message(message_id: str | None) -> None:
    if not message_id or message_id in _recent_message_id_set:
        return
    if len(_recent_message_ids) == RECENT_MESSAGE_ID_LIMIT:
        _recent_message_id_set.discard(_recent_message_ids[0])
    _recent_message_ids.append(message_id)
    _recent_message_id_set.add(message_id)

# ══════════════════════════════════════════════════════════════
# 分類字典（新增品種直接加在這裡）
# ══════════════════════════════════════════════════════════════
CATEGORIES = {
    # 種苗銷售
    "顆":"種苗銷售","盒":"種苗銷售","株":"種苗銷售",
    "珍妮":"種苗銷售","侏儒":"種苗銷售","斑葉":"種苗銷售",
    "黃月":"種苗銷售","神鉅":"種苗銷售","妙蛙":"種苗銷售",
    "火鶴":"種苗銷售","鹿角":"種苗銷售","爆米花":"種苗銷售",
    "白怪":"種苗銷售","紅水晶":"種苗銷售","聖靈":"種苗銷售",
    "粉斑":"種苗銷售","豆豆龍":"種苗銷售","nano":"種苗銷售",
    "omg":"種苗銷售","delta":"種苗銷售","戰鬥機":"種苗銷售",
    "種苗":"種苗銷售","植物":"種苗銷售",
    # 運費
    "運費":"運費","空軍":"運費","黑貓":"運費","郵局":"運費",
    # 耗材
    "水苔":"材料耗材","紙箱":"材料耗材","膠膜":"材料耗材",
    "悶箱":"材料耗材","棉花":"材料耗材","垃圾袋":"材料耗材",
    "土":"材料耗材","盆":"材料耗材","木板":"材料耗材",
    # 水電
    "電費":"水電費","水電":"水電費",
    # 固定支出
    "租金":"租金","薪水":"薪資","薪資":"薪資",
    # 換匯
    "換人民幣":"換匯","換rmb":"換匯","匯款":"換匯","人民幣":"換匯",
    # 設備
    "燈":"設備","機器":"設備","噴霧":"設備","冰箱":"設備","ro機":"設備",
    # 農藥肥料
    "農藥":"農藥肥料","肥料":"農藥肥料",
    # 活動銷售
    "711":"活動銷售","福袋":"活動銷售",
}

# 支出 → 成本結構對應
COST_STRUCTURE = {
    "進貨":    "進貨成本",
    "種苗銷售":"進貨成本",
    "運費":    "物流費用",
    "材料耗材":"耗材費用",
    "水電費":  "水電費",
    "租金":    "租金",
    "薪資":    "人事費用",
    "換匯":    "換匯",
    "設備":    "設備投資",
    "農藥肥料":"農藥肥料",
}

# 銷售管道自動判斷
CHANNELS = {
    "蝦皮":"蝦皮","shopee":"蝦皮",
    "w上架":"蝦皮/網路","上架":"蝦皮/網路",
    "711":"7-11福袋","福袋":"7-11福袋",
    "大陸":"大陸出口","黃浩":"大陸出口","馬薩":"大陸出口",
    "韓國":"韓國出口",
    "植系":"植系",
}

# ══════════════════════════════════════════════════════════════
# 輔助函數
# ══════════════════════════════════════════════════════════════
def guess_category(text: str) -> str:
    t = text.lower()
    for k, v in sorted(CATEGORIES.items(), key=lambda item: len(item[0]), reverse=True):
        if k.lower() in t:
            return v
    return "其他"

def guess_channel(item: str, customer: str, raw: str) -> str:
    combined = f"{item} {customer} {raw}".lower()
    for k, v in sorted(CHANNELS.items(), key=lambda item: len(item[0]), reverse=True):
        if k.lower() in combined:
            return v
    return "直接客戶"

def guess_cost_structure(category: str, tx_type: str) -> str:
    if tx_type != TX_EXPENSE:
        return ""
    return COST_STRUCTURE.get(category, "其他費用")

def extract_qty_and_unit_price(item: str, amount: int) -> tuple:
    # $單價×數量
    m = re.search(r'\$(\d[\d,]*)\s*[*×x]\s*(\d[\d,]*)', item)
    if m:
        return to_int(m.group(2)), to_int(m.group(1))
    # 數量×單價
    m = re.search(r'(\d[\d,]*)\s*[*×x]\s*\$?(\d[\d,]*)', item)
    if m:
        qty = to_int(m.group(1))
        return qty, to_int(m.group(2))
    # 純數量+單位
    m = re.search(r'(\d[\d,]*)\s*[顆盒個株]', item)
    if m:
        qty = to_int(m.group(1))
        return qty, (round(amount / qty) if qty > 0 else 0)
    return 0, 0

# ══════════════════════════════════════════════════════════════
# 訊息解析
# 格式：收入/支出 金額 品項 [客戶] [進價X] [未付/已付/已收/部分收]
# ══════════════════════════════════════════════════════════════
def parse_message(text: str) -> dict | None:
    text = re.sub(r"\s+", " ", text.replace("　", " ")).strip()
    if text.startswith("+"):
        text = f"{TX_INCOME} " + text[1:].strip()
    elif text.startswith("-"):
        text = f"{TX_EXPENSE} " + text[1:].strip()

    m = re.match(rf"^({TX_INCOME}|{TX_EXPENSE})\s+([\d,]+)\s+(.+)$", text, re.IGNORECASE)
    if not m:
        return None

    tx_type    = m.group(1)
    amount     = to_int(m.group(2))
    remaining  = m.group(3).strip()
    if amount <= 0 or not remaining:
        return None

    status_pattern = "|".join(map(re.escape, STATUS_VALUES))
    status_match = re.search(rf"\s+({status_pattern})$", remaining)
    status = status_match.group(1) if status_match else (
        STATUS_UNPAID if tx_type == TX_INCOME else STATUS_PAID
    )
    if status_match:
        remaining = remaining[:status_match.start()].strip()

    # 進價解析（格式：進價XX）
    cost_per_unit = 0
    cm = re.search(r'進價\s*(\d[\d,]*)', remaining)
    if cm:
        cost_per_unit = to_int(cm.group(1))
        remaining = re.sub(r'進價\s*\d[\d,]*', '', remaining).strip()

    # 換匯金額（RMB）
    rmb = ""
    rmb_match = re.search(r'(\d[\d,]*)\s*rmb|rmb\s*(\d[\d,]*)|人民幣\s*(\d[\d,]*)', remaining, re.IGNORECASE)
    if rmb_match:
        rmb = next(group for group in rmb_match.groups() if group)

    item = remaining
    customer = ""
    if tx_type == TX_INCOME and " " in remaining:
        item, customer = [part.strip() for part in remaining.rsplit(" ", 1)]

    # 數量與售出單價
    qty, unit_price = extract_qty_and_unit_price(item, amount)

    # 毛利（需有進價和數量）
    gross_profit = ""
    if tx_type == TX_INCOME and cost_per_unit > 0 and qty > 0:
        gross_profit = amount - (cost_per_unit * qty)

    today    = today_tw()
    category = guess_category(f"{item} {text}")

    return {
        # ── 對應試算表欄位順序 ─────────────────────────────────
        "date":           today.strftime("%Y/%m/%d"),  # 日期
        "type":           tx_type,                      # 類型
        "amount":         amount,                       # 金額(NT$)
        "item":           item,                         # 品項說明
        "customer":       customer,                     # 客戶/廠商
        "status":         status,                       # 付款狀態
        "pay_date":       today.strftime("%Y/%m/%d") if status in PAID_STATUSES else "",  # 付款日期
        "qty":            qty,                          # 數量
        "unit_price":     unit_price,                   # 售出單價
        "cost_per_unit":  cost_per_unit,                # 進貨單價
        "gross_profit":   gross_profit,                 # 毛利
        "channel":        guess_channel(item, customer, text),  # 銷售管道
        "days_to_collect": "",                          # 收款天數（付款後計算）
        "note":           "",                           # 備註
        "category":       category,                     # 分類
        "cost_structure": guess_cost_structure(category, tx_type),  # 支出成本結構
        "month":          f"{today.month}月",           # 月份
        "rmb":            rmb,                          # 換匯金額(RMB)
        "raw":            text,                         # 原始備註
    }

# ══════════════════════════════════════════════════════════════
# 寫入交易記錄
# ══════════════════════════════════════════════════════════════
def append_transaction(wb, data: dict) -> int:
    sheet = wb.worksheet(SHEET_TRANSACTIONS)
    row = [
        data["date"], data["type"], data["amount"], data["item"],
        data["customer"], data["status"], data["pay_date"],
        data["qty"], data["unit_price"], data["cost_per_unit"],
        data["gross_profit"], data["channel"], data["days_to_collect"],
        data["note"], data["category"], data["cost_structure"],
        data["month"], data["rmb"], data["raw"],
    ]
    response = sheet.append_row(
        row,
        value_input_option="RAW",
        insert_data_option="INSERT_ROWS",
    )
    return parse_append_row_number(response, sheet)

# ══════════════════════════════════════════════════════════════
# 自動更新應收帳款
# 條件：收入 + 付款狀態為「未付」或「部分收」
# ══════════════════════════════════════════════════════════════
def update_receivables(wb, data: dict) -> bool:
    if data["type"] != TX_INCOME:
        return False
    if data["status"] not in (STATUS_UNPAID, STATUS_PARTIAL):
        return False

    sheet = wb.worksheet(SHEET_RECEIVABLES)

    # 已收金額：部分收時為 0（需手動更新），未付為 0
    collected = 0
    outstanding = data["amount"] - collected

    row = [
        data["date"],       # 交易日期
        data["customer"],   # 客戶
        data["item"],       # 品項
        data["amount"],     # 應收金額
        collected,          # 已收金額
        outstanding,        # 未收餘額
        "",                 # 付款期限（手動填）
        "",                 # 逾期天數
        data["raw"],        # 備註
    ]
    sheet.append_row(
        row,
        value_input_option="RAW",
        insert_data_option="INSERT_ROWS",
    )
    return True

# ══════════════════════════════════════════════════════════════
# Line 回覆格式
# ══════════════════════════════════════════════════════════════
def format_reply(data: dict, row_num: int, added_to_receivables: bool) -> str:
    icon        = "💰" if data["type"] == TX_INCOME else "💸"
    status_icon = {
        STATUS_COLLECTED: "✅",
        STATUS_PAID: "✅",
        STATUS_UNPAID: "⏳",
        STATUS_PARTIAL: "⚠️",
    }.get(data["status"], "")

    lines = [
        f"{icon} 已記錄到第 {row_num} 行",
        "─" * 22,
        f"日期｜{data['date']}",
        f"類型｜{data['type']}",
        f"金額｜NT$ {data['amount']:,}",
        f"品項｜{data['item']}",
    ]
    if data["customer"]:
        lines.append(f"客戶｜{data['customer']}")
    lines.append(f"狀態｜{status_icon} {data['status']}")
    lines.append(f"管道｜{data['channel']}")
    lines.append(f"分類｜{data['category']}")

    if data["qty"]:
        lines.append(f"數量｜{data['qty']:,} 顆/盒")
    if data["unit_price"]:
        lines.append(f"售價｜NT$ {data['unit_price']:,} / 顆")
    if data["cost_per_unit"]:
        lines.append(f"進價｜NT$ {data['cost_per_unit']:,} / 顆")
    if data["gross_profit"] != "":
        margin = round(data["gross_profit"] / data["amount"] * 100, 1) if data["amount"] else 0
        lines.append(f"毛利｜NT$ {data['gross_profit']:,}（{margin}%）")
    if data["cost_structure"]:
        lines.append(f"成本｜{data['cost_structure']}")
    if data["rmb"]:
        lines.append(f"換匯｜RMB {data['rmb']}")

    lines.append("─" * 22)
    if added_to_receivables:
        lines.append("📌 已同步到應收帳款")
    lines.append("✏️ 如需修改請直接更新試算表")
    return "\n".join(lines)

HELP_TEXT = """🌿 溫室帳目機器人

📥 基本格式：
收入 金額 品項 客戶
支出 金額 品項

📌 範例：
收入 50000 珍妮500顆 李淵男
支出 18000 大陸運費
收入 32000 爆米花1000顆 美琪 未付
+50000 侏儒黃月1000顆 吳政翰
-4200 淘寶悶箱

📊 加進價自動算毛利：
收入 50000 珍妮500顆 李淵男 進價30
→ 數量/售價/進價/毛利自動計算

🏪 管道自動判斷：
直接客戶／蝦皮網路／7-11福袋
大陸出口／韓國出口／植系

💡 付款狀態（選填）：
已收 / 已付 / 未付 / 部分收

📌 未付/部分收的收入會自動
   同步到「應收帳款」分頁

輸入「說明」再次查看"""

# ══════════════════════════════════════════════════════════════
# Flask 路由
# ══════════════════════════════════════════════════════════════
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
def handle_message(event):
    user_text   = event.message.text.strip()
    reply_token = event.reply_token
    message_id  = getattr(event.message, "id", None)

    with ApiClient(configuration) as api_client:
        line_api = MessagingApi(api_client)

        if user_text in ("說明", "help", "Help", "HELP", "?", "？"):
            reply = HELP_TEXT
        elif has_seen_message(message_id):
            reply = "這則訊息已經處理過，沒有重複寫入。"
        else:
            parsed = parse_message(user_text)
            if parsed is None:
                reply = (
                    "❌ 無法解析此訊息\n\n"
                    "格式：收入 50000 珍妮500顆 李淵男\n"
                    "　　　支出 18000 大陸運費\n\n"
                    "輸入「說明」查看完整格式"
                )
            else:
                try:
                    wb      = get_workbook()
                    row_num = append_transaction(wb, parsed)
                    remember_message(message_id)
                except Exception:
                    logger.exception("Failed to append transaction")
                    reply = "⚠️ 寫入交易記錄時發生錯誤，請稍後再試。"
                else:
                    added = False
                    receivable_error = False
                    try:
                        added = update_receivables(wb, parsed)
                    except Exception:
                        receivable_error = True
                        logger.exception("Failed to update receivables")

                    reply = format_reply(parsed, row_num, added)
                    if receivable_error:
                        reply += "\n⚠️ 交易已記錄，但應收帳款同步失敗，請手動確認應收表。"

        line_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=reply)]
            )
        )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
