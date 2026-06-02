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

SHEET_TRANSACTIONS = "📋 交易記錄"
SHEET_RECEIVABLES  = "💰 應收帳款"

TX_INCOME   = "收入"
TX_EXPENSE  = "支出"
STATUS_UNPAID     = "未付"
STATUS_PAID       = "已付"
STATUS_COLLECTED  = "已收"
STATUS_PARTIAL    = "部分收"
STATUS_VALUES     = (STATUS_UNPAID, STATUS_PAID, STATUS_COLLECTED, STATUS_PARTIAL)
PAID_STATUSES     = {STATUS_PAID, STATUS_COLLECTED}

# 交易記錄欄位對應（1-based，對應 Google Sheets 欄位）
COL_DATE        = 1   # 日期
COL_TYPE        = 2   # 類型
COL_AMOUNT      = 3   # 金額
COL_ITEM        = 4   # 品項說明
COL_CUSTOMER    = 5   # 客戶/廠商
COL_STATUS      = 6   # 付款狀態
COL_PAY_DATE    = 7   # 付款日期
COL_QTY         = 8   # 數量
COL_UNIT_PRICE  = 9   # 售出單價
COL_COST        = 10  # 進貨單價
COL_PROFIT      = 11  # 毛利
COL_CHANNEL     = 12  # 銷售管道
COL_DAYS        = 13  # 收款天數
COL_NOTE        = 14  # 備註
COL_CATEGORY    = 15  # 分類
COL_COST_STRUCT = 16  # 支出成本結構
COL_MONTH       = 17  # 月份
COL_RMB         = 18  # 換匯金額
COL_RAW         = 19  # 原始備註

# 應收帳款欄位
RECV_COL_DATE        = 1
RECV_COL_CUSTOMER    = 2
RECV_COL_ITEM        = 3
RECV_COL_AMOUNT      = 4
RECV_COL_COLLECTED   = 5
RECV_COL_OUTSTANDING = 6
RECV_COL_DUE         = 7
RECV_COL_OVERDUE     = 8
RECV_COL_NOTE        = 9

RECENT_MESSAGE_ID_LIMIT = 500
_recent_message_ids     = deque(maxlen=RECENT_MESSAGE_ID_LIMIT)
_recent_message_id_set: set[str] = set()
MIN_TRANSACTION_ROW = 3

# ══════════════════════════════════════════════════════════════
# Google Sheets 連線
# ══════════════════════════════════════════════════════════════
@lru_cache(maxsize=1)
def get_workbook():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        json.loads(GOOGLE_CREDS_JSON), scopes=scopes
    )
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def get_app_timezone():
    try:
        return ZoneInfo(APP_TIMEZONE)
    except ZoneInfoNotFoundError:
        return timezone(timedelta(hours=8))

def today_tw():
    return datetime.now(get_app_timezone()).date()

def to_int(s: str) -> int:
    return int(str(s).replace(",", ""))

def has_seen_message(mid):
    return bool(mid and mid in _recent_message_id_set)

def remember_message(mid):
    if not mid or mid in _recent_message_id_set:
        return
    if len(_recent_message_ids) == RECENT_MESSAGE_ID_LIMIT:
        _recent_message_id_set.discard(_recent_message_ids[0])
    _recent_message_ids.append(mid)
    _recent_message_id_set.add(mid)

# ══════════════════════════════════════════════════════════════
# 分類字典
# ══════════════════════════════════════════════════════════════
CATEGORIES = {
    "顆":"種苗銷售","盒":"種苗銷售","株":"種苗銷售",
    "珍妮":"種苗銷售","侏儒":"種苗銷售","斑葉":"種苗銷售",
    "黃月":"種苗銷售","神鉅":"種苗銷售","妙蛙":"種苗銷售",
    "火鶴":"種苗銷售","鹿角":"種苗銷售","爆米花":"種苗銷售",
    "白怪":"種苗銷售","紅水晶":"種苗銷售","聖靈":"種苗銷售",
    "粉斑":"種苗銷售","豆豆龍":"種苗銷售","nano":"種苗銷售",
    "omg":"種苗銷售","delta":"種苗銷售","戰鬥機":"種苗銷售",
    "種苗":"種苗銷售","植物":"種苗銷售",
    "運費":"運費","空軍":"運費","黑貓":"運費","郵局":"運費",
    "水苔":"材料耗材","紙箱":"材料耗材","膠膜":"材料耗材",
    "悶箱":"材料耗材","棉花":"材料耗材","垃圾袋":"材料耗材",
    "土":"材料耗材","盆":"材料耗材","木板":"材料耗材",
    "電費":"水電費","水電":"水電費",
    "租金":"租金","薪水":"薪資","薪資":"薪資",
    "換人民幣":"換匯","換rmb":"換匯","匯款":"換匯","人民幣":"換匯",
    "燈":"設備","機器":"設備","噴霧":"設備","冰箱":"設備","ro機":"設備",
    "農藥":"農藥肥料","肥料":"農藥肥料",
    "711":"活動銷售","福袋":"活動銷售",
}

COST_STRUCTURE_MAP = {
    "種苗銷售":"進貨成本","進貨":"進貨成本",
    "運費":"物流費用","材料耗材":"耗材費用",
    "水電費":"水電費","租金":"租金","薪資":"人事費用",
    "換匯":"換匯","設備":"設備投資","農藥肥料":"農藥肥料",
}

CHANNELS = {
    "蝦皮":"蝦皮","shopee":"蝦皮",
    "w上架":"蝦皮/網路","上架":"蝦皮/網路",
    "711":"7-11福袋","福袋":"7-11福袋",
    "大陸":"大陸出口","黃浩":"大陸出口","馬薩":"大陸出口",
    "韓國":"韓國出口","植系":"植系",
}

# ══════════════════════════════════════════════════════════════
# 輔助函數
# ══════════════════════════════════════════════════════════════
def guess_category(text: str) -> str:
    t = text.lower()
    for k, v in sorted(CATEGORIES.items(), key=lambda x: len(x[0]), reverse=True):
        if k.lower() in t:
            return v
    return "其他"

def guess_channel(item: str, customer: str, raw: str) -> str:
    combined = f"{item} {customer} {raw}".lower()
    for k, v in sorted(CHANNELS.items(), key=lambda x: len(x[0]), reverse=True):
        if k.lower() in combined:
            return v
    return "直接客戶"

def guess_cost_structure(category: str, tx_type: str) -> str:
    if tx_type != TX_EXPENSE:
        return ""
    return COST_STRUCTURE_MAP.get(category, "其他費用")

def extract_qty_and_unit_price(item: str, amount: int) -> tuple:
    m = re.search(r'\$(\d[\d,]*)\s*[*×x]\s*(\d[\d,]*)', item)
    if m:
        return to_int(m.group(2)), to_int(m.group(1))
    m = re.search(r'(\d[\d,]*)\s*[*×x]\s*\$?(\d[\d,]*)', item)
    if m:
        qty = to_int(m.group(1))
        return qty, to_int(m.group(2))
    m = re.search(r'(\d[\d,]*)\s*[顆盒個株]', item)
    if m:
        qty = to_int(m.group(1))
        return qty, (round(amount / qty) if qty > 0 else 0)
    return 0, 0

# ══════════════════════════════════════════════════════════════
# 新增交易：解析訊息
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

    tx_type   = m.group(1)
    amount    = to_int(m.group(2))
    remaining = m.group(3).strip()
    if amount <= 0 or not remaining:
        return None

    status_pat = "|".join(map(re.escape, STATUS_VALUES))
    sm = re.search(rf"\s+({status_pat})$", remaining)
    status    = sm.group(1) if sm else (STATUS_UNPAID if tx_type == TX_INCOME else STATUS_PAID)
    if sm:
        remaining = remaining[:sm.start()].strip()

    cost_per_unit = 0
    cm = re.search(r'進價\s*(\d[\d,]*)', remaining)
    if cm:
        cost_per_unit = to_int(cm.group(1))
        remaining = re.sub(r'進價\s*\d[\d,]*', '', remaining).strip()

    rmb = ""
    rm = re.search(r'(\d[\d,]*)\s*rmb|rmb\s*(\d[\d,]*)|人民幣\s*(\d[\d,]*)', remaining, re.IGNORECASE)
    if rm:
        rmb = next(g for g in rm.groups() if g)

    item = remaining
    customer = ""
    if tx_type == TX_INCOME and " " in remaining:
        item, customer = [p.strip() for p in remaining.rsplit(" ", 1)]

    qty, unit_price = extract_qty_and_unit_price(item, amount)

    gross_profit = ""
    if tx_type == TX_INCOME and cost_per_unit > 0 and qty > 0:
        gross_profit = amount - (cost_per_unit * qty)

    today    = today_tw()
    category = guess_category(f"{item} {text}")

    return {
        "date":           today.strftime("%Y/%m/%d"),
        "type":           tx_type,
        "amount":         amount,
        "item":           item,
        "customer":       customer,
        "status":         status,
        "pay_date":       today.strftime("%Y/%m/%d") if status in PAID_STATUSES else "",
        "qty":            qty,
        "unit_price":     unit_price,
        "cost_per_unit":  cost_per_unit,
        "gross_profit":   gross_profit,
        "channel":        guess_channel(item, customer, text),
        "days_to_collect": "",
        "note":           "",
        "category":       category,
        "cost_structure": guess_cost_structure(category, tx_type),
        "month":          f"{today.month}月",
        "rmb":            rmb,
        "raw":            text,
    }

# ══════════════════════════════════════════════════════════════
# 更新付款狀態：解析指令
# 格式：更新 行號 已收／已付／部分收 [已收金額]
# 範例：更新 21 已收
#       更新 21 部分收 30000
# ══════════════════════════════════════════════════════════════
def parse_update_command(text: str) -> dict | None:
    m = re.match(
        rf"^更新\s+(\d+)\s+({STATUS_COLLECTED}|{STATUS_PAID}|{STATUS_PARTIAL})"
        rf"(?:\s+([\d,]+))?$",
        text.strip()
    )
    if not m:
        return None
    status = m.group(2)
    collected = to_int(m.group(3)) if m.group(3) else None
    if status == STATUS_PARTIAL and (collected is None or collected <= 0):
        return None
    return {
        "row":       int(m.group(1)),
        "status":    status,
        "collected": collected,
    }

def receivable_note(raw: str, row_num: int) -> str:
    return f"[交易行號:{row_num}] {raw}".strip()

def find_receivable_row(all_recv: list[list[str]], row_num: int, customer: str, item: str, amount: int) -> int | None:
    marker = f"[交易行號:{row_num}]"
    fallback_matches = []
    for i, r in enumerate(all_recv, start=1):
        if i < MIN_TRANSACTION_ROW:
            continue

        note = r[RECV_COL_NOTE - 1] if len(r) >= RECV_COL_NOTE else ""
        if marker in note:
            return i

        r_customer = r[RECV_COL_CUSTOMER - 1] if len(r) >= RECV_COL_CUSTOMER else ""
        r_item     = r[RECV_COL_ITEM - 1]     if len(r) >= RECV_COL_ITEM     else ""
        r_amount   = r[RECV_COL_AMOUNT - 1]   if len(r) >= RECV_COL_AMOUNT   else ""
        try:
            same_amount = to_int(r_amount) == amount
        except ValueError:
            same_amount = False
        if r_customer == customer and r_item == item and same_amount:
            fallback_matches.append(i)

    return fallback_matches[0] if len(fallback_matches) == 1 else None

# ══════════════════════════════════════════════════════════════
# 更新付款狀態：寫入試算表
# ══════════════════════════════════════════════════════════════
def apply_status_update(wb, row_num: int, new_status: str, collected: int | None) -> dict:
    if row_num < MIN_TRANSACTION_ROW:
        return {"ok": False, "error": "請確認行號是否為交易資料列"}

    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    row_data  = tx_sheet.row_values(row_num)

    if not row_data or len(row_data) < COL_AMOUNT:
        return {"ok": False, "error": f"找不到第 {row_num} 行，請確認行號是否正確"}

    orig_date_str = row_data[COL_DATE - 1]
    tx_type  = row_data[COL_TYPE - 1] if len(row_data) >= COL_TYPE else ""
    item     = row_data[COL_ITEM - 1]     if len(row_data) >= COL_ITEM     else ""
    customer = row_data[COL_CUSTOMER - 1] if len(row_data) >= COL_CUSTOMER else ""
    amount_str = row_data[COL_AMOUNT - 1] if len(row_data) >= COL_AMOUNT   else "0"
    try:
        total_amount = to_int(amount_str) if amount_str else 0
    except ValueError:
        return {"ok": False, "error": "此列金額格式不正確，請先檢查試算表"}

    if tx_type not in (TX_INCOME, TX_EXPENSE):
        return {"ok": False, "error": "此列不是有效的交易資料列"}
    if new_status == STATUS_PARTIAL:
        if collected is None or collected <= 0:
            return {"ok": False, "error": "部分收必須填寫已收金額"}
        if collected >= total_amount:
            return {"ok": False, "error": "部分收金額需小於交易金額；若已全收請用「已收」"}

    today    = today_tw()
    pay_date = today.strftime("%Y/%m/%d")

    # 收款天數（只有完全收款才計算）
    days_to_collect = ""
    if new_status in PAID_STATUSES and orig_date_str:
        try:
            orig = datetime.strptime(orig_date_str, "%Y/%m/%d").date()
            days_to_collect = (today - orig).days
        except ValueError:
            pass

    # 批次更新三個欄位（減少 API 呼叫次數）
    tx_sheet.batch_update([
        {"range": f"F{row_num}", "values": [[new_status]]},
        {"range": f"G{row_num}", "values": [[pay_date]]},
        {"range": f"M{row_num}", "values": [[days_to_collect]]},
    ])

    # 同步應收帳款
    recv_updated = False
    try:
        recv_sheet = wb.worksheet(SHEET_RECEIVABLES)
        all_recv   = recv_sheet.get_all_values()
        recv_row = find_receivable_row(all_recv, row_num, customer, item, total_amount)

        if tx_type == TX_INCOME and recv_row:
            if new_status == STATUS_COLLECTED:
                # 全額已收
                recv_sheet.batch_update([
                    {"range": f"E{recv_row}", "values": [[total_amount]]},
                    {"range": f"F{recv_row}", "values": [[0]]},
                ])
                recv_updated = True
            elif new_status == STATUS_PARTIAL:
                # 部分收款
                outstanding = total_amount - collected
                recv_sheet.batch_update([
                    {"range": f"E{recv_row}", "values": [[collected]]},
                    {"range": f"F{recv_row}", "values": [[outstanding]]},
                ])
                recv_updated = True
    except Exception:
        logger.exception("Failed to update receivables during status update")

    return {
        "ok":           True,
        "row_num":      row_num,
        "type":         tx_type,
        "item":         item,
        "customer":     customer,
        "new_status":   new_status,
        "pay_date":     pay_date,
        "days":         days_to_collect,
        "recv_updated": recv_updated,
        "collected":    collected,
        "total_amount": amount_str,
    }

# ══════════════════════════════════════════════════════════════
# 寫入新交易
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
    response = sheet.append_row(row, value_input_option="RAW", insert_data_option="INSERT_ROWS")
    m = re.search(r"![A-Z]+(\d+)(?::[A-Z]+\d+)?$",
                  response.get("updates", {}).get("updatedRange", ""))
    return int(m.group(1)) if m else len(sheet.col_values(1))

def update_receivables(wb, data: dict, row_num: int) -> bool:
    if data["type"] != TX_INCOME:
        return False
    if data["status"] not in (STATUS_UNPAID, STATUS_PARTIAL):
        return False

    sheet      = wb.worksheet(SHEET_RECEIVABLES)
    all_values = sheet.get_all_values()

    # 找到「合計未收」那行，新資料插入在它上方
    insert_at = len(all_values) + 1  # 預設：找不到就附加在最後
    for i, row in enumerate(all_values, start=1):
        if row and "合計" in str(row[0]):
            insert_at = i
            break

    sheet.insert_row(
        [data["date"], data["customer"], data["item"],
         data["amount"], 0, data["amount"], "", "", receivable_note(data["raw"], row_num)],
        insert_at,
        value_input_option="RAW",
    )
    return True

# ══════════════════════════════════════════════════════════════
# 回覆格式
# ══════════════════════════════════════════════════════════════
def format_new_transaction_reply(data: dict, row_num: int, recv_added: bool) -> str:
    icon        = "💰" if data["type"] == TX_INCOME else "💸"
    status_icon = {"已收":"✅","已付":"✅","未付":"⏳","部分收":"⚠️"}.get(data["status"], "")
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
    lines += [
        f"狀態｜{status_icon} {data['status']}",
        f"管道｜{data['channel']}",
        f"分類｜{data['category']}",
    ]
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
    if recv_added:
        lines.append("📌 已同步到應收帳款")
    lines.append("✏️ 更新付款狀態請輸入：")
    lines.append(f"更新 {row_num} 已收")
    return "\n".join(lines)

def format_update_reply(result: dict) -> str:
    status_icon = {"已收":"✅","已付":"✅","部分收":"⚠️"}.get(result["new_status"], "")
    lines = [
        f"✅ 第 {result['row_num']} 行已更新",
        "─" * 22,
        f"品項｜{result['item']}",
    ]
    if result["customer"]:
        lines.append(f"客戶｜{result['customer']}")
    lines += [
        f"狀態｜{status_icon} {result['new_status']}",
        f"付款日｜{result['pay_date']}",
    ]
    if result["new_status"] == STATUS_PARTIAL and result["collected"]:
        lines.append(f"已收｜NT$ {result['collected']:,} / NT$ {result['total_amount']}")
    if result["days"] != "":
        lines.append(f"收款天數｜{result['days']} 天")
    lines.append("─" * 22)
    if result["recv_updated"]:
        lines.append("📌 應收帳款已同步更新")
    elif result.get("type") == TX_INCOME:
        lines.append("⚠️ 應收帳款請手動確認")
    return "\n".join(lines)

HELP_TEXT = """🌿 溫室帳目機器人

📥 新增交易：
收入 金額 品項 客戶
支出 金額 品項

📌 新增範例：
收入 50000 珍妮500顆 李淵男
支出 18000 大陸運費
收入 32000 爆米花1000顆 美琪 未付
+50000 侏儒黃月1000顆 吳政翰
-4200 淘寶悶箱
收入 50000 珍妮500顆 李淵男 進價30

🔄 更新付款狀態：
更新 行號 狀態

📌 更新範例：
更新 21 已收
更新 21 已付
更新 21 部分收 30000

💡 付款狀態：
已收 / 已付 / 未付 / 部分收

📌 未付/部分收收入會自動
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

        # ── 說明指令 ───────────────────────────────────────────
        if user_text in ("說明", "help", "Help", "HELP", "?", "？"):
            reply = HELP_TEXT

        # ── 更新付款狀態 ────────────────────────────────────────
        elif user_text.startswith("更新"):
            update_cmd = parse_update_command(user_text)
            if update_cmd is None:
                reply = (
                    "❌ 更新格式錯誤\n\n"
                    "正確格式：\n"
                    "更新 行號 已收\n"
                    "更新 行號 已付\n"
                    "更新 行號 部分收 金額\n\n"
                    "範例：更新 21 已收\n"
                    "　　　更新 21 部分收 30000"
                )
            else:
                try:
                    wb     = get_workbook()
                    result = apply_status_update(
                        wb,
                        update_cmd["row"],
                        update_cmd["status"],
                        update_cmd["collected"],
                    )
                    reply = format_update_reply(result) if result["ok"] else f"❌ {result['error']}"
                except Exception:
                    logger.exception("Update failed")
                    reply = "⚠️ 更新失敗，請稍後再試。"

        # ── 新增交易（防重複） ──────────────────────────────────
        elif has_seen_message(message_id):
            reply = "這則訊息已經處理過，沒有重複寫入。"

        else:
            parsed = parse_message(user_text)
            if parsed is None:
                reply = (
                    "❌ 無法解析此訊息\n\n"
                    "新增：收入 50000 珍妮500顆 李淵男\n"
                    "更新：更新 21 已收\n\n"
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
                    recv_added = False
                    recv_err   = False
                    try:
                        recv_added = update_receivables(wb, parsed, row_num)
                    except Exception:
                        recv_err = True
                        logger.exception("Failed to update receivables")

                    reply = format_new_transaction_reply(parsed, row_num, recv_added)
                    if recv_err:
                        reply += "\n⚠️ 交易已記錄，但應收帳款同步失敗。"

        line_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=reply)]
            )
        )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
