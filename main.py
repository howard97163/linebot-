import os
import re
import json
import logging
import calendar
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
CRON_SECRET               = os.environ.get("CRON_SECRET", "")

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler       = WebhookHandler(LINE_CHANNEL_SECRET)

SHEET_TRANSACTIONS = "📋 交易記錄"
SHEET_RECEIVABLES  = "💰 應收帳款"
SHEET_MONTHLY_OVERVIEW = "📊 月份總覽"
SHEET_CUSTOMER_ANALYSIS = "👥 客戶分析"

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
COL_EXCHANGE_RATE = 19  # 匯率
COL_RAW         = 20  # 原始備註
COL_TX_ID       = 21  # 交易編號（TX-0001，固定不隨排序變動）
COL_DELETED_AT  = 22  # 刪除日期（軟刪除標記）

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
RECV_COL_DELETED_AT  = 10

# 客戶分析欄位
CUST_COL_NAME        = 1
CUST_COL_TOTAL       = 2
CUST_COL_COLLECTED   = 3
CUST_COL_OUTSTANDING = 4
CUST_COL_COUNT       = 5
CUST_COL_RECENT      = 6

RECENT_MESSAGE_ID_LIMIT = 500
_recent_message_ids     = deque(maxlen=RECENT_MESSAGE_ID_LIMIT)
_recent_message_id_set: set[str] = set()
MIN_TRANSACTION_ROW = 3
TX_LAST_COL = COL_DELETED_AT
RECV_LAST_COL = RECV_COL_DELETED_AT
MONTHLY_LAST_COL = 19
ROW_COLOR_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
ROW_COLOR_GREEN = {"red": 0.91, "green": 0.97, "blue": 0.94}
ROW_COLOR_DELETED = {"red": 0.82, "green": 0.82, "blue": 0.82}
STATUS_COLOR_PAID = {"red": 0.72, "green": 0.95, "blue": 0.82}
STATUS_COLOR_UNPAID = {"red": 0.98, "green": 0.82, "blue": 0.82}
STATUS_COLOR_PARTIAL = {"red": 1.0, "green": 0.93, "blue": 0.68}
NUMBER_PATTERN = r"\d[\d,]*(?:\.\d+)?"
FIXED_COST_STRUCTURES = {"水電費", "租金", "人事費用"}

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

def worksheet_by_names(wb, *names):
    last_error = None
    for name in names:
        try:
            return wb.worksheet(name)
        except Exception as e:
            last_error = e
    raise last_error

def get_app_timezone():
    try:
        return ZoneInfo(APP_TIMEZONE)
    except ZoneInfoNotFoundError:
        return timezone(timedelta(hours=8))

def today_tw():
    return datetime.now(get_app_timezone()).date()

def to_int(s: str) -> int:
    cleaned = str(s).strip().replace(",", "").replace("NT$", "").replace("$", "")
    cleaned = re.sub(r"[^\d.-]", "", cleaned)
    if cleaned in ("", "-", ".", "-."):
        raise ValueError(f"Invalid number: {s}")
    return int(round(float(cleaned)))

def to_number(s: str):
    cleaned = str(s).strip().replace(",", "").replace("NT$", "").replace("$", "")
    cleaned = re.sub(r"[^\d.-]", "", cleaned)
    if cleaned in ("", "-", ".", "-."):
        raise ValueError(f"Invalid number: {s}")
    value = float(cleaned)
    return int(value) if value.is_integer() else value

def column_letter(col: int) -> str:
    letters = ""
    while col:
        col, remainder = divmod(col - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def cell_a1(row: int, col: int) -> str:
    return f"{column_letter(col)}{row}"

def range_a1(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    return f"{cell_a1(start_row, start_col)}:{cell_a1(end_row, end_col)}"

TX_ID_PATTERN = re.compile(r"^TX-(\d+)$", re.IGNORECASE)

def format_tx_id(n: int) -> str:
    return f"TX-{n:04d}"

def normalize_tx_id(text: str) -> str | None:
    t = str(text).strip().upper()
    if not t:
        return None
    if t.startswith("TX") and "-" not in t:
        t = t.replace("TX", "TX-", 1)
    m = TX_ID_PATTERN.match(t)
    if m:
        return format_tx_id(int(m.group(1)))
    if re.match(r"^\d+$", t):
        return format_tx_id(int(t))
    return None

def tx_id_number(tx_id: str) -> int:
    m = TX_ID_PATTERN.match(str(tx_id).strip())
    return int(m.group(1)) if m else 0

def next_tx_id(sheet) -> str:
    max_n = 0
    try:
        col_values = sheet.col_values(COL_TX_ID)
    except Exception:
        col_values = []
    for v in col_values:
        normalized = normalize_tx_id(v)
        if normalized:
            max_n = max(max_n, tx_id_number(normalized))
    return format_tx_id(max_n + 1)

def find_row_by_tx_id(sheet, tx_id: str) -> int:
    normalized = normalize_tx_id(tx_id)
    if not normalized:
        return 0
    try:
        col_values = sheet.col_values(COL_TX_ID)
    except Exception:
        return 0
    for row_num, value in enumerate(col_values, start=1):
        if normalize_tx_id(value) == normalized:
            return row_num
    return 0

def ensure_transaction_ids(sheet) -> int:
    values = sheet.get_all_values()
    next_n = tx_id_number(next_tx_id(sheet))
    updates = []
    for row_num in range(MIN_TRANSACTION_ROW, last_data_row(values) + 1):
        row = values[row_num - 1] if len(values) >= row_num else []
        existing = row[COL_TX_ID - 1] if len(row) >= COL_TX_ID else ""
        if normalize_tx_id(existing):
            continue
        updates.append({
            "range": cell_a1(row_num, COL_TX_ID),
            "values": [[format_tx_id(next_n)]],
        })
        next_n += 1
    if updates:
        sheet.batch_update(updates)
    return len(updates)

def parse_date_value(value: str, default_year: int | None = None):
    value = value.strip()
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", value)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
        except ValueError:
            return None

    m = re.match(r"^(\d{1,2})[/-](\d{1,2})$", value)
    if m and default_year is not None:
        try:
            return datetime(default_year, int(m.group(1)), int(m.group(2))).date()
        except ValueError:
            return None

    return None

def parse_optional_transaction_date(text: str):
    today = today_tw()
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})\s+(.+)$", text)
    if m:
        tx_date = parse_date_value(f"{m.group(1)}/{m.group(2)}/{m.group(3)}")
        if tx_date is None:
            return None, text
        return tx_date, m.group(4).strip()

    m = re.match(r"^(\d{1,2})[/-](\d{1,2})\s+(.+)$", text)
    if m:
        tx_date = parse_date_value(f"{m.group(1)}/{m.group(2)}", today.year)
        if tx_date is None:
            return None, text
        return tx_date, m.group(3).strip()

    return today, text

def add_months(d, months: int = 1):
    month_index = d.month - 1 + months
    year = d.year + month_index // 12
    month = month_index % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return datetime(year, month, day).date()

def extract_due_date(text: str, tx_date):
    pattern = r"(?:付款期限|期限|到期日|到期)\s*[:：]?\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}[/-]\d{1,2})"
    m = re.search(pattern, text)
    if not m:
        return add_months(tx_date, 1), text

    due_date = parse_date_value(m.group(1), tx_date.year)
    if due_date is None:
        return None, text

    text = (text[:m.start()] + text[m.end():]).strip()
    text = re.sub(r"\s+", " ", text)
    return due_date, text

def calculate_overdue_days(due_date, outstanding: int):
    if not due_date or outstanding <= 0:
        return ""
    days = (today_tw() - due_date).days
    return days if days > 0 else ""

def overdue_background_color(overdue_days):
    if not overdue_days:
        return None
    if overdue_days > 90:
        return {"red": 1.0, "green": 0.80, "blue": 0.80}
    if overdue_days > 60:
        return {"red": 1.0, "green": 0.88, "blue": 0.70}
    if overdue_days > 30:
        return {"red": 1.0, "green": 0.96, "blue": 0.65}
    return None

def alternating_color_for_row(row_num: int):
    return ROW_COLOR_GREEN if (row_num - MIN_TRANSACTION_ROW) % 2 else ROW_COLOR_WHITE

def is_soft_deleted_row(row: list[str], deleted_col: int) -> bool:
    return len(row) >= deleted_col and bool(str(row[deleted_col - 1]).strip())

def deleted_at_value(row: list[str], deleted_col: int) -> str:
    return str(row[deleted_col - 1]).strip() if len(row) >= deleted_col else ""

def deleted_date_from_row(row: list[str], deleted_col: int):
    value = deleted_at_value(row, deleted_col)
    return parse_date_value(value) if value else None

def append_row_format_request(sheet, row_num: int, last_col: int, color: dict, requests: list):
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet.id,
                "startRowIndex": row_num - 1,
                "endRowIndex": row_num,
                "startColumnIndex": 0,
                "endColumnIndex": last_col,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": color,
                }
            },
            "fields": "userEnteredFormat.backgroundColor",
        }
    })

def apply_row_background(sheet, row_num: int, last_col: int, color: dict):
    requests = []
    append_row_format_request(sheet, row_num, last_col, color, requests)
    sheet.spreadsheet.batch_update({"requests": requests})

def apply_deleted_row_format(sheet, row_num: int, last_col: int):
    apply_row_background(sheet, row_num, last_col, ROW_COLOR_DELETED)

def apply_soft_deleted_formats(sheet, deleted_col: int, last_col: int):
    values = sheet.get_all_values()
    end_row = last_data_row(values)
    total_row = find_total_row(values)
    if total_row:
        end_row = min(end_row, total_row - 1)
    requests = []
    for row_num in range(MIN_TRANSACTION_ROW, end_row + 1):
        row = values[row_num - 1] if len(values) >= row_num else []
        if is_soft_deleted_row(row, deleted_col):
            append_row_format_request(sheet, row_num, last_col, ROW_COLOR_DELETED, requests)
    if requests:
        sheet.spreadsheet.batch_update({"requests": requests})

def transaction_status_background(status: str):
    if status in (STATUS_PAID, STATUS_COLLECTED):
        return STATUS_COLOR_PAID
    if status == STATUS_UNPAID:
        return STATUS_COLOR_UNPAID
    if status == STATUS_PARTIAL:
        return STATUS_COLOR_PARTIAL
    return None

def find_total_row(values: list[list[str]]) -> int | None:
    for i, row in enumerate(values, start=1):
        if row and "合計" in str(row[0]):
            return i
    return None

def refresh_receivable_totals(sheet):
    values = sheet.get_all_values()
    total_row = find_total_row(values)
    if not total_row or total_row <= MIN_TRANSACTION_ROW:
        return

    last_detail_row = total_row - 1
    deleted_col_letter = column_letter(RECV_COL_DELETED_AT)
    sheet.batch_update([
        {"range": f"D{total_row}", "values": [[f"=SUMIF({deleted_col_letter}{MIN_TRANSACTION_ROW}:{deleted_col_letter}{last_detail_row},\"\",D{MIN_TRANSACTION_ROW}:D{last_detail_row})"]]},
        {"range": f"E{total_row}", "values": [[f"=SUMIF({deleted_col_letter}{MIN_TRANSACTION_ROW}:{deleted_col_letter}{last_detail_row},\"\",E{MIN_TRANSACTION_ROW}:E{last_detail_row})"]]},
        {"range": f"F{total_row}", "values": [[f"=SUMIF({deleted_col_letter}{MIN_TRANSACTION_ROW}:{deleted_col_letter}{last_detail_row},\"\",F{MIN_TRANSACTION_ROW}:F{last_detail_row})"]]},
    ], value_input_option="USER_ENTERED")

def last_data_row(values: list[list[str]], date_col: int = 1) -> int:
    last_row = 0
    for i, row in enumerate(values, start=1):
        if i < MIN_TRANSACTION_ROW:
            continue
        date_value = row[date_col - 1] if len(row) >= date_col else ""
        if parse_date_value(str(date_value)):
            last_row = i
    return last_row

def sort_sheet_by_date(sheet, last_col: int):
    values = sheet.get_all_values()
    end_row = last_data_row(values)
    if end_row <= MIN_TRANSACTION_ROW:
        return
    sheet.spreadsheet.batch_update({
        "requests": [{
            "sortRange": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": MIN_TRANSACTION_ROW - 1,
                    "endRowIndex": end_row,
                    "startColumnIndex": 0,
                    "endColumnIndex": last_col,
                },
                "sortSpecs": [{
                    "dimensionIndex": COL_DATE - 1,
                    "sortOrder": "ASCENDING",
                }],
            }
        }]
    })

def apply_alternating_row_colors(sheet, last_col: int):
    values = sheet.get_all_values()
    end_row = last_data_row(values)
    if end_row < MIN_TRANSACTION_ROW:
        return

    apply_alternating_row_colors_to(sheet, last_col, end_row)

def apply_transaction_status_formats(sheet):
    values = sheet.get_all_values()
    end_row = last_data_row(values)
    if end_row < MIN_TRANSACTION_ROW:
        return

    requests = []
    for row_num in range(MIN_TRANSACTION_ROW, end_row + 1):
        row = values[row_num - 1] if len(values) >= row_num else []
        status = row[COL_STATUS - 1].strip() if len(row) >= COL_STATUS and row[COL_STATUS - 1] else ""
        color = transaction_status_background(status)
        if not color:
            continue
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": row_num - 1,
                    "endRowIndex": row_num,
                    "startColumnIndex": COL_STATUS - 1,
                    "endColumnIndex": COL_STATUS,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": color,
                        "textFormat": {"bold": True},
                    }
                },
                "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold",
            }
        })
    if requests:
        sheet.spreadsheet.batch_update({"requests": requests})

def refresh_transaction_formats(sheet):
    apply_alternating_row_colors(sheet, TX_LAST_COL)
    apply_transaction_status_formats(sheet)
    apply_soft_deleted_formats(sheet, COL_DELETED_AT, TX_LAST_COL)

def apply_alternating_row_colors_to(sheet, last_col: int, end_row: int):
    if end_row < MIN_TRANSACTION_ROW:
        return

    requests = []
    for row_num in range(MIN_TRANSACTION_ROW, end_row + 1):
        color = alternating_color_for_row(row_num)
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": row_num - 1,
                    "endRowIndex": row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": last_col,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": color,
                    }
                },
                "fields": "userEnteredFormat.backgroundColor",
            }
        })

    if requests:
        sheet.spreadsheet.batch_update({"requests": requests})

def refresh_receivable_overdue_formats(sheet):
    apply_alternating_row_colors(sheet, RECV_LAST_COL)
    values = sheet.get_all_values()
    end_row = last_data_row(values)
    total_row = find_total_row(values)
    if total_row:
        end_row = min(end_row, total_row - 1)

    overdue_updates = []
    format_requests = []
    for row_num in range(MIN_TRANSACTION_ROW, end_row + 1):
        row = values[row_num - 1] if len(values) >= row_num else []
        if is_soft_deleted_row(row, RECV_COL_DELETED_AT):
            append_row_format_request(sheet, row_num, RECV_LAST_COL, ROW_COLOR_DELETED, format_requests)
            continue
        outstanding = 0
        due_date = None
        try:
            outstanding = to_int(row[RECV_COL_OUTSTANDING - 1]) if len(row) >= RECV_COL_OUTSTANDING and row[RECV_COL_OUTSTANDING - 1] else 0
        except ValueError:
            outstanding = 0
        if len(row) >= RECV_COL_DUE and row[RECV_COL_DUE - 1]:
            due_date = parse_date_value(row[RECV_COL_DUE - 1])
        if outstanding <= 0:
            existing_overdue = row[RECV_COL_OVERDUE - 1] if len(row) >= RECV_COL_OVERDUE else ""
            try:
                existing_overdue_days = to_int(existing_overdue) if existing_overdue else ""
            except ValueError:
                existing_overdue_days = ""
            if existing_overdue_days:
                color = overdue_background_color(existing_overdue_days)
                if color:
                    append_row_format_request(sheet, row_num, RECV_LAST_COL, color, format_requests)
            continue
        overdue_days = calculate_overdue_days(due_date, outstanding)
        if overdue_days:
            overdue_updates.append({
                "range": cell_a1(row_num, RECV_COL_OVERDUE),
                "values": [[overdue_days]],
            })
            color = overdue_background_color(overdue_days)
            if color:
                append_row_format_request(sheet, row_num, RECV_LAST_COL, color, format_requests)

    if overdue_updates:
        sheet.batch_update(overdue_updates)
    if format_requests:
        sheet.spreadsheet.batch_update({"requests": format_requests})

def apply_receivable_row_format(sheet, row_num: int, overdue_days=""):
    color = overdue_background_color(overdue_days) if overdue_days else None
    if not color:
        color = alternating_color_for_row(row_num)
    apply_row_background(sheet, row_num, RECV_LAST_COL, color)

def find_transaction_row(sheet, data: dict) -> int:
    values = sheet.get_all_values()
    candidates = []
    for row_num in range(MIN_TRANSACTION_ROW, last_data_row(values) + 1):
        row = values[row_num - 1] if len(values) >= row_num else []
        row_date = row[COL_DATE - 1] if len(row) >= COL_DATE else ""
        row_type = row[COL_TYPE - 1] if len(row) >= COL_TYPE else ""
        row_amount = row[COL_AMOUNT - 1] if len(row) >= COL_AMOUNT else ""
        row_item = row[COL_ITEM - 1] if len(row) >= COL_ITEM else ""
        row_customer = row[COL_CUSTOMER - 1] if len(row) >= COL_CUSTOMER else ""
        row_raw = row[COL_RAW - 1] if len(row) >= COL_RAW else ""
        try:
            amount_matches = to_int(row_amount) == data["amount"]
        except ValueError:
            amount_matches = False
        if (
            row_date == data["date"]
            and row_type == data["type"]
            and amount_matches
            and row_item == data["item"]
            and row_customer == data["customer"]
            and row_raw == data["raw"]
        ):
            candidates.append(row_num)
    return candidates[-1] if candidates else 0

def organize_transaction_sheet(sheet, data: dict) -> int:
    sort_sheet_by_date(sheet, TX_LAST_COL)
    refresh_transaction_formats(sheet)
    if data.get("tx_id"):
        return find_row_by_tx_id(sheet, data["tx_id"])
    return find_transaction_row(sheet, data)

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
    "愛心氣球":"種苗銷售","白斑犀牛皮":"種苗銷售",
    "斑葉橘柄":"種苗銷售","斑葉神巨":"種苗銷售",
    "種苗":"種苗銷售","植物":"種苗銷售",
    # 物流/包材
    "運費":"一般運費","空軍":"空軍物流","黑貓":"黑貓宅配","郵局":"郵局寄送",
    "宅配":"宅配","貨運":"貨運","冷藏":"冷藏配送",
    "水苔":"水苔","紙箱":"紙箱","膠膜":"膠膜","膠帶":"膠帶",
    "氣泡布":"氣泡布","保麗龍":"保麗龍","紙板":"紙板","包材":"包材",
    "悶箱":"悶箱","棉花":"棉花","垃圾袋":"垃圾袋",
    "土":"介質土","盆":"盆器","木板":"木板",
    "電費":"水電費","水電":"水電費",
    "租金":"租金","薪水":"薪資","薪資":"薪資",
    "工讀":"工讀","臨時工":"臨時工","獎金":"獎金","加班":"加班",
    "換人民幣":"人民幣換匯","換rmb":"人民幣換匯","匯款":"匯款","人民幣":"人民幣換匯",
    "燈":"燈具設備","機器":"機器設備","噴霧":"噴霧設備","冰箱":"冰箱設備","ro機":"RO設備",
    "維修":"設備維修","修理":"設備維修","零件":"設備零件",
    "馬達":"馬達","風扇":"風扇","水泵":"水泵","噴頭":"噴頭","管線":"管線",
    "農藥":"農藥","肥料":"肥料",
    # 進出口
    "檢疫":"檢疫","關稅":"關稅",
    "罰金":"罰金","出口":"出口費用",
    "標籤帶":"標籤帶",
    # 差旅
    "機票":"機票","住宿":"住宿","計程車":"計程車",
    "接機":"接機","機加酒":"機加酒",
    # 餐飲交際
    "晚餐":"晚餐","尾牙":"尾牙","吃飯":"吃飯",
    # 損耗退貨
    "退錢":"退款","退款":"退款",
    # 行銷廣告
    "廣告":"廣告","投放":"廣告投放","拍攝":"拍攝","設計":"設計",
    "印刷":"印刷","名片":"名片",
    # 銀行財務
    "轉帳費":"轉帳費","匯費":"匯費","銀行手續費":"銀行手續費","利息":"利息",
    # 稅務規費
    "營業稅":"營業稅","所得稅":"所得稅","規費":"規費","牌照":"牌照",
    # 場地活動
    "攤位":"攤位費","市集":"市集活動","展覽":"展覽","報名費":"報名費",
}

COST_STRUCTURE_MAP = {
    "種苗銷售":"進貨成本","進貨":"進貨成本",
    "一般運費":"物流費用","空軍物流":"物流費用","黑貓宅配":"物流費用",
    "郵局寄送":"物流費用","宅配":"物流費用","貨運":"物流費用","冷藏配送":"物流費用",
    "水苔":"耗材費用","紙箱":"耗材費用","膠膜":"耗材費用","膠帶":"耗材費用",
    "氣泡布":"耗材費用","保麗龍":"耗材費用","紙板":"耗材費用","包材":"耗材費用",
    "悶箱":"耗材費用","棉花":"耗材費用","垃圾袋":"耗材費用",
    "介質土":"耗材費用","盆器":"耗材費用","木板":"耗材費用",
    "水電費":"水電費","租金":"租金","薪資":"人事費用",
    "工讀":"人事費用","臨時工":"人事費用","獎金":"人事費用","加班":"人事費用",
    "人民幣換匯":"換匯","匯款":"換匯",
    "燈具設備":"設備投資","機器設備":"設備投資","噴霧設備":"設備投資",
    "冰箱設備":"設備投資","RO設備":"設備投資",
    "設備維修":"設備維修","設備零件":"設備維修","馬達":"設備維修",
    "風扇":"設備維修","水泵":"設備維修","噴頭":"設備維修","管線":"設備維修",
    "農藥":"農藥肥料","肥料":"農藥肥料",
    "檢疫":"進出口費用","關稅":"進出口費用","罰金":"進出口費用",
    "出口費用":"進出口費用","標籤帶":"進出口費用",
    "機票":"差旅費","住宿":"差旅費","計程車":"差旅費","接機":"差旅費","機加酒":"差旅費",
    "晚餐":"餐飲交際","尾牙":"餐飲交際","吃飯":"餐飲交際",
    "退款":"損耗退貨",
    "廣告":"行銷廣告","廣告投放":"行銷廣告","拍攝":"行銷廣告",
    "設計":"行銷廣告","印刷":"行銷廣告","名片":"行銷廣告",
    "轉帳費":"銀行費用","匯費":"銀行費用","銀行手續費":"銀行費用","利息":"銀行費用",
    "營業稅":"稅務規費","所得稅":"稅務規費","規費":"稅務規費","牌照":"稅務規費",
    "攤位費":"活動費用","市集活動":"活動費用","展覽":"活動費用","報名費":"活動費用",
}

CHANNELS = {
    "大陸":"大陸出口","鉑茵植造":"大陸出口",
    "閩卉園藝":"大陸出口","卉通園藝":"大陸出口",
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
    m = re.search(rf'\$({NUMBER_PATTERN})\s*[*×x]\s*(\d[\d,]*)', item)
    if m:
        return to_int(m.group(2)), to_number(m.group(1))
    m = re.search(rf'(\d[\d,]*)\s*[*×x]\s*\$?({NUMBER_PATTERN})', item)
    if m:
        qty = to_int(m.group(1))
        return qty, to_number(m.group(2))
    m = re.search(r'(\d[\d,]*)\s*[顆棵盒個株]', item)
    if m:
        qty = to_int(m.group(1))
        return qty, (round(amount / qty) if qty > 0 else 0)
    return 0, 0

def number_from_field(value: str):
    m = re.search(NUMBER_PATTERN, str(value))
    if not m:
        return ""
    return to_number(m.group(0))

def extract_labeled_customer(text: str) -> tuple[str, str]:
    m = re.search(r"\s+(?:客戶|廠商|供應商|對象)\s*[:：]\s*(.+)$", text)
    if not m:
        return text, ""
    customer = m.group(1).strip()
    text = text[:m.start()].strip()
    return text, customer

def clean_item_text(text: str) -> str:
    text = re.sub(r'(\d[\d,]*)\s*rmb|rmb\s*(\d[\d,]*)|人民幣\s*(\d[\d,]*)', ' ', text, flags=re.IGNORECASE)
    text = re.sub(rf'(?:total|總額)\s*[:：]?\s*{NUMBER_PATTERN}', ' ', text, flags=re.IGNORECASE)
    text = re.sub(rf'(?:進價|單價|一袋成本|成本)\s*{NUMBER_PATTERN}', ' ', text)
    text = re.sub(r'\d[\d,]*\s*[顆棵盒個株袋]', ' ', text)
    return re.sub(r"\s+", " ", text).strip()

# ══════════════════════════════════════════════════════════════
# 新增交易：解析訊息（支援標籤格式與原本空格格式）
# ══════════════════════════════════════════════════════════════
LABEL_ALIASES = {
    "類型": "type", "收支": "type",
    "金額": "amount", "金钱": "amount", "價格": "amount",
    "品項": "item", "品项": "item", "項目": "item", "商品": "item",
    "客戶": "customer", "客户": "customer", "廠商": "customer",
    "厂商": "customer", "供應商": "customer", "對象": "customer",
    "狀態": "status", "状态": "status", "付款狀態": "status",
    "日期": "date", "交易日期": "date",
    "期限": "due", "付款期限": "due", "到期日": "due",
    "數量": "qty", "数量": "qty",
    "售出單價": "unit_price", "售價": "unit_price",
    "進貨單價": "cost", "進價": "cost", "进价": "cost", "成本": "cost",
    "人民幣": "rmb", "人民币": "rmb", "rmb": "rmb", "RMB": "rmb",
    "已收": "collected", "已收金額": "collected",
    "備註": "note", "备注": "note",
}

_LABEL_TOKEN_RE = re.compile(
    r"(" + "|".join(sorted(map(re.escape, LABEL_ALIASES), key=len, reverse=True)) + r")\s*[：:]"
)

def is_labeled_format(text: str) -> bool:
    return bool(_LABEL_TOKEN_RE.search(text))

def extract_labeled_fields(text: str) -> dict:
    fields = {}
    matches = list(_LABEL_TOKEN_RE.finditer(text))
    for idx, m in enumerate(matches):
        label = m.group(1)
        start = m.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        value = text[start:end].strip().strip("，,、")
        key = LABEL_ALIASES.get(label)
        if key and value:
            fields[key] = value
    return fields

def parse_labeled_message(text: str) -> dict | None:
    fields = extract_labeled_fields(text)

    tx_type = fields.get("type", "").strip()
    if tx_type in ("+",):
        tx_type = TX_INCOME
    elif tx_type in ("-",):
        tx_type = TX_EXPENSE
    if tx_type not in (TX_INCOME, TX_EXPENSE):
        return None

    if "amount" not in fields or "item" not in fields:
        return None
    try:
        amount = to_int(fields["amount"])
    except ValueError:
        return None
    if amount <= 0:
        return None

    item = fields["item"].strip()
    if not item:
        return None
    customer = fields.get("customer", "").strip()

    today = today_tw()
    tx_date = today
    if "date" in fields:
        tx_date = parse_date_value(fields["date"], today.year)
        if tx_date is None:
            return None

    status = fields.get("status", "").strip()
    initial_collected = 0
    if "collected" in fields:
        try:
            initial_collected = to_int(fields["collected"])
        except ValueError:
            return None
        if initial_collected <= 0 or initial_collected >= amount:
            return None
        status = STATUS_PARTIAL
    if status and status not in STATUS_VALUES:
        return None
    if not status:
        status = STATUS_UNPAID if tx_type == TX_INCOME else STATUS_PAID
    if status == STATUS_PARTIAL and initial_collected == 0:
        return None

    due_date = add_months(tx_date, 1)
    if "due" in fields:
        due_date = parse_date_value(fields["due"], tx_date.year)
        if due_date is None:
            return None

    cost_per_unit = 0
    if "cost" in fields:
        try:
            cost_value = number_from_field(fields["cost"])
            cost_per_unit = cost_value if cost_value != "" else 0
        except ValueError:
            return None

    rmb = ""
    exchange_rate = ""
    if "rmb" in fields:
        try:
            rmb = to_int(fields["rmb"])
        except ValueError:
            return None
        if rmb > 0:
            exchange_rate = round(amount / rmb, 2)
        else:
            rmb = ""

    item_qty, item_unit_price = extract_qty_and_unit_price(item, amount)
    qty = item_qty
    if "qty" in fields:
        try:
            qty_value = number_from_field(fields["qty"])
        except ValueError:
            return None
        if qty_value != "":
            qty = int(qty_value)

    unit_price = item_unit_price
    if "unit_price" in fields:
        try:
            unit_price_value = number_from_field(fields["unit_price"])
        except ValueError:
            return None
        if unit_price_value != "":
            unit_price = unit_price_value
    elif qty and not unit_price:
        unit_price = round(amount / qty) if qty > 0 else 0

    gross_profit = ""
    if tx_type == TX_INCOME and cost_per_unit > 0 and qty > 0:
        gross_profit = amount - (cost_per_unit * qty)

    outstanding = amount - initial_collected if status == STATUS_PARTIAL else amount
    overdue_days = calculate_overdue_days(due_date, outstanding) if tx_type == TX_INCOME and status in (STATUS_UNPAID, STATUS_PARTIAL) else ""
    category = guess_category(f"{item} {text}")
    raw_single_line = " ".join(l.strip() for l in text.splitlines() if l.strip())

    return {
        "date":           tx_date.strftime("%Y/%m/%d"),
        "type":           tx_type,
        "amount":         amount,
        "item":           item,
        "customer":       customer,
        "status":         status,
        "pay_date":       tx_date.strftime("%Y/%m/%d") if status in PAID_STATUSES else "",
        "due_date":       due_date.strftime("%Y/%m/%d") if tx_type == TX_INCOME and status in (STATUS_UNPAID, STATUS_PARTIAL) else "",
        "initial_collected": initial_collected,
        "outstanding":    outstanding,
        "overdue_days":   overdue_days,
        "qty":            qty,
        "unit_price":     unit_price,
        "cost_per_unit":  cost_per_unit,
        "gross_profit":   gross_profit,
        "channel":        guess_channel(item, customer, text),
        "days_to_collect": "",
        "note":           fields.get("note", ""),
        "category":       category,
        "cost_structure": guess_cost_structure(category, tx_type),
        "month":          f"{tx_date.month}月",
        "rmb":            rmb,
        "exchange_rate":  exchange_rate,
        "raw":            raw_single_line,
    }

def parse_message(text: str) -> dict | None:
    if is_labeled_format(text):
        return parse_labeled_message(text)
    return None

BUTTON_TEMPLATE_TEXTS = {"記收入", "記支出", "更新收款"}

def is_blank_labeled_template(text: str) -> bool:
    if not is_labeled_format(text):
        return False
    fields = extract_labeled_fields(text)
    return not all(fields.get(k, "").strip() for k in ("type", "amount", "item"))

def should_ignore_template_prompt(text: str) -> bool:
    normalized = re.sub(r"\s+", "", text.replace("　", " ")).strip()
    if normalized in BUTTON_TEMPLATE_TEXTS:
        return True
    if re.fullmatch(r"更新TX-已收", normalized, re.IGNORECASE):
        return True
    if re.fullmatch(r"(刪除|恢復)TX-", normalized, re.IGNORECASE):
        return True
    return is_blank_labeled_template(text)

def parse_spaced_message(text: str) -> dict | None:
    text = re.sub(r"\s+", " ", text.replace("　", " ")).strip()
    tx_date, text = parse_optional_transaction_date(text)
    if tx_date is None:
        return None

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

    due_date, remaining = extract_due_date(remaining, tx_date)
    if due_date is None:
        return None

    initial_collected = 0
    partial_match = re.search(rf"\s+(?:{STATUS_PARTIAL}|部分付款)\s+([\d,]+)$", remaining)
    if partial_match:
        initial_collected = to_int(partial_match.group(1))
        if initial_collected <= 0 or initial_collected >= amount:
            return None
        status = STATUS_PARTIAL
        remaining = remaining[:partial_match.start()].strip()
    else:
        status = None

    status_pat = "|".join(map(re.escape, STATUS_VALUES))
    sm = re.search(rf"\s+({status_pat})$", remaining) if status is None else None
    status = status or (sm.group(1) if sm else (STATUS_UNPAID if tx_type == TX_INCOME else STATUS_PAID))
    if sm:
        remaining = remaining[:sm.start()].strip()

    remaining, customer = extract_labeled_customer(remaining)

    cost_per_unit = 0
    cm = re.search(rf'進價\s*({NUMBER_PATTERN})', remaining)
    if cm:
        cost_per_unit = to_number(cm.group(1))
        remaining = re.sub(rf'進價\s*{NUMBER_PATTERN}', '', remaining).strip()

    rmb = ""
    exchange_rate = ""
    rm = re.search(r'(\d[\d,]*)\s*rmb|rmb\s*(\d[\d,]*)|人民幣\s*(\d[\d,]*)', remaining, re.IGNORECASE)
    if rm:
        rmb = to_int(next(g for g in rm.groups() if g))
        if rmb > 0:
            exchange_rate = round(amount / rmb, 2)

    qty, unit_price = extract_qty_and_unit_price(remaining, amount)

    item = clean_item_text(remaining)
    has_detail_parts = bool(cost_per_unit or rmb or qty)
    if not customer and " " in item and (tx_type == TX_INCOME or has_detail_parts):
        item, customer = [p.strip() for p in item.rsplit(" ", 1)]
    if not item:
        return None

    gross_profit = ""
    if tx_type == TX_INCOME and cost_per_unit > 0 and qty > 0:
        gross_profit = amount - (cost_per_unit * qty)

    outstanding = amount - initial_collected if status == STATUS_PARTIAL else amount
    overdue_days = calculate_overdue_days(due_date, outstanding) if tx_type == TX_INCOME and status in (STATUS_UNPAID, STATUS_PARTIAL) else ""
    category = guess_category(f"{item} {text}")

    return {
        "date":           tx_date.strftime("%Y/%m/%d"),
        "type":           tx_type,
        "amount":         amount,
        "item":           item,
        "customer":       customer,
        "status":         status,
        "pay_date":       tx_date.strftime("%Y/%m/%d") if status in PAID_STATUSES else "",
        "due_date":       due_date.strftime("%Y/%m/%d") if tx_type == TX_INCOME and status in (STATUS_UNPAID, STATUS_PARTIAL) else "",
        "initial_collected": initial_collected,
        "outstanding":    outstanding,
        "overdue_days":   overdue_days,
        "qty":            qty,
        "unit_price":     unit_price,
        "cost_per_unit":  cost_per_unit,
        "gross_profit":   gross_profit,
        "channel":        guess_channel(item, customer, text),
        "days_to_collect": "",
        "note":           "",
        "category":       category,
        "cost_structure": guess_cost_structure(category, tx_type),
        "month":          f"{tx_date.month}月",
        "rmb":            rmb,
        "exchange_rate":  exchange_rate,
        "raw":            text,
    }

# ══════════════════════════════════════════════════════════════
# 更新付款狀態：解析指令
# 格式：更新 交易編號 已收／已付／部分收 [本次收款金額]
# 範例：更新 TX-0021 已收
#       更新 21 部分收 30000
# ══════════════════════════════════════════════════════════════
def parse_update_command(text: str) -> dict | None:
    m = re.match(
        rf"^更新\s+(TX-?\d+|\d+)\s+({STATUS_COLLECTED}|{STATUS_PAID}|{STATUS_PARTIAL})"
        rf"(?:\s+([\d,]+))?$",
        text.strip(),
        re.IGNORECASE,
    )
    if not m:
        return None
    tx_id = normalize_tx_id(m.group(1))
    if tx_id is None:
        return None
    status = m.group(2)
    collected = to_int(m.group(3)) if m.group(3) else None
    if status == STATUS_PARTIAL and (collected is None or collected <= 0):
        return None
    return {
        "tx_id":     tx_id,
        "status":    status,
        "collected": collected,
    }

def parse_delete_command(text: str) -> dict | None:
    m = re.match(r"^刪除(?:交易(?:紀錄|記錄)?)?\s+(TX-?\d+|\d+)$", text.strip(), re.IGNORECASE)
    if not m:
        return None
    tx_id = normalize_tx_id(m.group(1))
    if tx_id is None:
        return None
    return {"tx_id": tx_id}

def parse_restore_command(text: str) -> dict | None:
    m = re.match(r"^恢復(?:交易(?:紀錄|記錄)?)?\s+(TX-?\d+|\d+)$", text.strip(), re.IGNORECASE)
    if not m:
        return None
    tx_id = normalize_tx_id(m.group(1))
    if tx_id is None:
        return None
    return {"tx_id": tx_id}

def receivable_note(raw: str, tx_id: str) -> str:
    return f"[{tx_id}] {raw}".strip()

def find_receivable_row(all_recv: list[list[str]], tx_id: str, customer: str, item: str, amount: int, raw: str = "") -> int | None:
    marker = f"[{tx_id}]"
    raw_matches = []
    fallback_matches = []
    for i, r in enumerate(all_recv, start=1):
        if i < MIN_TRANSACTION_ROW:
            continue

        note = r[RECV_COL_NOTE - 1] if len(r) >= RECV_COL_NOTE else ""
        if marker in note:
            return i
        if raw and raw in note:
            raw_matches.append(i)

        r_customer = r[RECV_COL_CUSTOMER - 1] if len(r) >= RECV_COL_CUSTOMER else ""
        r_item     = r[RECV_COL_ITEM - 1]     if len(r) >= RECV_COL_ITEM     else ""
        r_amount   = r[RECV_COL_AMOUNT - 1]   if len(r) >= RECV_COL_AMOUNT   else ""
        try:
            same_amount = to_int(r_amount) == amount
        except ValueError:
            same_amount = False
        if r_customer == customer and r_item == item and same_amount:
            fallback_matches.append(i)

    if len(raw_matches) == 1:
        return raw_matches[0]
    return fallback_matches[0] if len(fallback_matches) == 1 else None

# ══════════════════════════════════════════════════════════════
# 更新付款狀態：寫入試算表
# ══════════════════════════════════════════════════════════════
def apply_status_update(wb, tx_id: str, new_status: str, collected: int | None) -> dict:
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    row_num = find_row_by_tx_id(tx_sheet, tx_id)
    if not row_num or row_num < MIN_TRANSACTION_ROW:
        return {"ok": False, "error": f"找不到交易編號 {tx_id}，請確認編號是否正確"}
    row_data  = tx_sheet.row_values(row_num)

    if not row_data or len(row_data) < COL_AMOUNT:
        return {"ok": False, "error": f"找不到交易編號 {tx_id} 的資料，請確認試算表"}
    if is_soft_deleted_row(row_data, COL_DELETED_AT):
        return {"ok": False, "error": f"交易 {tx_id} 已標記刪除，若要更新請先輸入「恢復 {tx_id}」"}

    orig_date_str = row_data[COL_DATE - 1]
    tx_type  = row_data[COL_TYPE - 1] if len(row_data) >= COL_TYPE else ""
    item     = row_data[COL_ITEM - 1]     if len(row_data) >= COL_ITEM     else ""
    customer = row_data[COL_CUSTOMER - 1] if len(row_data) >= COL_CUSTOMER else ""
    amount_str = row_data[COL_AMOUNT - 1] if len(row_data) >= COL_AMOUNT   else "0"
    raw = row_data[COL_RAW - 1] if len(row_data) >= COL_RAW else ""
    try:
        total_amount = to_int(amount_str) if amount_str else 0
    except ValueError:
        return {"ok": False, "error": "此列金額格式不正確，請先檢查試算表"}

    if tx_type not in (TX_INCOME, TX_EXPENSE):
        return {"ok": False, "error": "此列不是有效的交易資料列"}
    if new_status == STATUS_PARTIAL:
        if collected is None or collected <= 0:
            return {"ok": False, "error": "部分收必須填寫本次收到的金額"}

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
        {"range": cell_a1(row_num, COL_STATUS), "values": [[new_status]]},
        {"range": cell_a1(row_num, COL_PAY_DATE), "values": [[pay_date]]},
        {"range": cell_a1(row_num, COL_DAYS), "values": [[days_to_collect]]},
    ])
    apply_transaction_status_formats(tx_sheet)

    # 同步應收帳款
    recv_updated = False
    total_collected = collected
    outstanding = ""
    try:
        recv_sheet = wb.worksheet(SHEET_RECEIVABLES)
        all_recv   = recv_sheet.get_all_values()
        recv_row = find_receivable_row(all_recv, tx_id, customer, item, total_amount, raw)

        if tx_type == TX_INCOME and recv_row:
            recv_data = all_recv[recv_row - 1] if len(all_recv) >= recv_row else []
            due_date_str = recv_data[RECV_COL_DUE - 1] if len(recv_data) >= RECV_COL_DUE else ""
            due_date = parse_date_value(due_date_str) if due_date_str else None
            existing_overdue_text = recv_data[RECV_COL_OVERDUE - 1] if len(recv_data) >= RECV_COL_OVERDUE else ""
            try:
                existing_overdue_days = to_int(existing_overdue_text) if existing_overdue_text else ""
            except ValueError:
                existing_overdue_days = ""
            if new_status == STATUS_COLLECTED:
                # 全額已收
                recv_sheet.batch_update([
                    {"range": f"E{recv_row}", "values": [[total_amount]]},
                    {"range": f"F{recv_row}", "values": [[0]]},
                    {"range": f"I{recv_row}", "values": [[receivable_note(raw, tx_id)]]},
                ])
                if existing_overdue_days:
                    apply_receivable_row_format(recv_sheet, recv_row, existing_overdue_days)
                recv_updated = True
            elif new_status == STATUS_PARTIAL:
                # 部分收款：輸入金額視為「本次收款」，需累加到既有已收金額。
                previous_collected_text = recv_data[RECV_COL_COLLECTED - 1] if len(recv_data) >= RECV_COL_COLLECTED else "0"
                try:
                    previous_collected = to_int(previous_collected_text) if previous_collected_text else 0
                except ValueError:
                    previous_collected = 0
                total_collected = previous_collected + collected
                if total_collected >= total_amount:
                    total_collected = total_amount
                    outstanding = 0
                    new_status = STATUS_COLLECTED
                    days_to_collect = ""
                    if orig_date_str:
                        try:
                            orig = datetime.strptime(orig_date_str, "%Y/%m/%d").date()
                            days_to_collect = (today - orig).days
                        except ValueError:
                            pass
                    tx_sheet.batch_update([
                        {"range": cell_a1(row_num, COL_STATUS), "values": [[new_status]]},
                        {"range": cell_a1(row_num, COL_DAYS), "values": [[days_to_collect]]},
                    ])
                    apply_transaction_status_formats(tx_sheet)
                else:
                    outstanding = total_amount - total_collected
                overdue_days = calculate_overdue_days(due_date, outstanding)
                receivable_updates = [
                    {"range": f"E{recv_row}", "values": [[total_collected]]},
                    {"range": f"F{recv_row}", "values": [[outstanding]]},
                    {"range": f"I{recv_row}", "values": [[receivable_note(raw, tx_id)]]},
                ]
                if outstanding > 0:
                    receivable_updates.append({"range": f"H{recv_row}", "values": [[overdue_days]]})
                recv_sheet.batch_update(receivable_updates)
                apply_receivable_row_format(
                    recv_sheet,
                    recv_row,
                    overdue_days if outstanding > 0 else existing_overdue_days,
                )
                recv_updated = True
    except Exception:
        logger.exception("Failed to update receivables during status update")

    return {
        "ok":           True,
        "tx_id":        tx_id,
        "row_num":      row_num,
        "type":         tx_type,
        "item":         item,
        "customer":     customer,
        "new_status":   new_status,
        "pay_date":     pay_date,
        "days":         days_to_collect,
        "recv_updated": recv_updated,
        "collected":    total_collected,
        "received_now": collected,
        "outstanding":  outstanding,
        "total_amount": amount_str,
    }

def apply_delete_transaction(wb, tx_id: str) -> dict:
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    row_num = find_row_by_tx_id(tx_sheet, tx_id)
    if not row_num or row_num < MIN_TRANSACTION_ROW:
        return {"ok": False, "error": f"找不到交易編號 {tx_id}，請確認編號是否正確"}
    row_data = tx_sheet.row_values(row_num)
    if not row_data or len(row_data) < COL_AMOUNT:
        return {"ok": False, "error": f"找不到交易編號 {tx_id} 的資料，請確認試算表"}

    tx_type = row_data[COL_TYPE - 1] if len(row_data) >= COL_TYPE else ""
    item = row_data[COL_ITEM - 1] if len(row_data) >= COL_ITEM else ""
    customer = row_data[COL_CUSTOMER - 1] if len(row_data) >= COL_CUSTOMER else ""
    amount_str = row_data[COL_AMOUNT - 1] if len(row_data) >= COL_AMOUNT else "0"
    raw = row_data[COL_RAW - 1] if len(row_data) >= COL_RAW else ""
    if is_soft_deleted_row(row_data, COL_DELETED_AT):
        return {"ok": False, "error": f"交易 {tx_id} 已經是灰色刪除狀態"}
    try:
        amount = to_int(amount_str) if amount_str else 0
    except ValueError:
        return {"ok": False, "error": "此列金額格式不正確，請先檢查試算表"}

    if tx_type not in (TX_INCOME, TX_EXPENSE):
        return {"ok": False, "error": "此列不是有效的交易資料列"}

    deleted_at = today_tw().strftime("%Y/%m/%d")
    recv_deleted = False
    recv_row = None
    if tx_type == TX_INCOME:
        try:
            recv_sheet = wb.worksheet(SHEET_RECEIVABLES)
            all_recv = recv_sheet.get_all_values()
            recv_row = find_receivable_row(all_recv, tx_id, customer, item, amount, raw)
            if recv_row:
                recv_sheet.batch_update([
                    {"range": cell_a1(recv_row, RECV_COL_DELETED_AT), "values": [[deleted_at]]},
                ])
                apply_deleted_row_format(recv_sheet, recv_row, RECV_LAST_COL)
                refresh_receivable_totals(recv_sheet)
                recv_deleted = True
        except Exception:
            logger.exception("Failed to delete related receivable")
            return {"ok": False, "error": "找到交易資料，但標記對應應收帳款時失敗，請稍後再試"}

    tx_sheet.batch_update([
        {"range": cell_a1(row_num, COL_DELETED_AT), "values": [[deleted_at]]},
    ])
    apply_deleted_row_format(tx_sheet, row_num, TX_LAST_COL)

    return {
        "ok": True,
        "tx_id": tx_id,
        "row_num": row_num,
        "deleted_at": deleted_at,
        "type": tx_type,
        "item": item,
        "customer": customer,
        "amount": amount,
        "recv_deleted": recv_deleted,
        "recv_row": recv_row,
    }

def apply_restore_transaction(wb, tx_id: str) -> dict:
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    row_num = find_row_by_tx_id(tx_sheet, tx_id)
    if not row_num or row_num < MIN_TRANSACTION_ROW:
        return {"ok": False, "error": f"找不到交易編號 {tx_id}，請確認編號是否正確"}
    row_data = tx_sheet.row_values(row_num)
    if not row_data or len(row_data) < COL_AMOUNT:
        return {"ok": False, "error": f"找不到交易編號 {tx_id} 的資料，請確認試算表"}
    if not is_soft_deleted_row(row_data, COL_DELETED_AT):
        return {"ok": False, "error": f"交易 {tx_id} 目前不是灰色刪除狀態"}

    tx_type = row_data[COL_TYPE - 1] if len(row_data) >= COL_TYPE else ""
    item = row_data[COL_ITEM - 1] if len(row_data) >= COL_ITEM else ""
    customer = row_data[COL_CUSTOMER - 1] if len(row_data) >= COL_CUSTOMER else ""
    amount_str = row_data[COL_AMOUNT - 1] if len(row_data) >= COL_AMOUNT else "0"
    raw = row_data[COL_RAW - 1] if len(row_data) >= COL_RAW else ""
    try:
        amount = to_int(amount_str) if amount_str else 0
    except ValueError:
        return {"ok": False, "error": "此列金額格式不正確，請先檢查試算表"}

    tx_sheet.batch_update([
        {"range": cell_a1(row_num, COL_DELETED_AT), "values": [[""]]},
    ])
    refresh_transaction_formats(tx_sheet)

    recv_restored = False
    recv_row = None
    if tx_type == TX_INCOME:
        try:
            recv_sheet = wb.worksheet(SHEET_RECEIVABLES)
            all_recv = recv_sheet.get_all_values()
            recv_row = find_receivable_row(all_recv, tx_id, customer, item, amount, raw)
            if recv_row:
                recv_data = all_recv[recv_row - 1] if len(all_recv) >= recv_row else []
                overdue = recv_data[RECV_COL_OVERDUE - 1] if len(recv_data) >= RECV_COL_OVERDUE else ""
                recv_sheet.batch_update([
                    {"range": cell_a1(recv_row, RECV_COL_DELETED_AT), "values": [[""]]},
                ])
                apply_receivable_row_format(recv_sheet, recv_row, overdue)
                refresh_receivable_totals(recv_sheet)
                recv_restored = True
        except Exception:
            logger.exception("Failed to restore related receivable")
            return {"ok": False, "error": "交易已恢復，但恢復對應應收帳款時失敗，請稍後再輸入「整理」確認"}

    return {
        "ok": True,
        "tx_id": tx_id,
        "row_num": row_num,
        "type": tx_type,
        "item": item,
        "customer": customer,
        "amount": amount,
        "recv_restored": recv_restored,
        "recv_row": recv_row,
    }

def purge_soft_deleted_rows(sheet, deleted_col: int, retention_days: int = 30) -> int:
    values = sheet.get_all_values()
    total_row = find_total_row(values)
    end_row = last_data_row(values)
    if total_row:
        end_row = min(end_row, total_row - 1)
    cutoff = today_tw() - timedelta(days=retention_days)
    rows_to_delete = []
    for row_num in range(MIN_TRANSACTION_ROW, end_row + 1):
        row = values[row_num - 1] if len(values) >= row_num else []
        deleted_date = deleted_date_from_row(row, deleted_col)
        if deleted_date and deleted_date <= cutoff:
            rows_to_delete.append(row_num)

    for row_num in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_num)
    return len(rows_to_delete)

# ══════════════════════════════════════════════════════════════
# 寫入新交易
# ══════════════════════════════════════════════════════════════
def append_transaction(wb, data: dict) -> tuple[int, str]:
    sheet = wb.worksheet(SHEET_TRANSACTIONS)
    tx_id = next_tx_id(sheet)
    data["tx_id"] = tx_id
    row = [
        data["date"], data["type"], data["amount"], data["item"],
        data["customer"], data["status"], data["pay_date"],
        data["qty"], data["unit_price"], data["cost_per_unit"],
        data["gross_profit"], data["channel"], data["days_to_collect"],
        data["note"], data["category"], data["cost_structure"],
        data["month"], data["rmb"], data["exchange_rate"], data["raw"], tx_id, "",
    ]
    sheet.append_row(row, value_input_option="RAW", insert_data_option="INSERT_ROWS")
    sorted_row_num = organize_transaction_sheet(sheet, data)
    return (sorted_row_num if sorted_row_num else len(sheet.col_values(1)), tx_id)

def update_receivables(wb, data: dict, tx_id: str) -> bool:
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
         data["amount"], data["initial_collected"], data["outstanding"],
         data["due_date"], data["overdue_days"], receivable_note(data["raw"], tx_id), ""],
        insert_at,
        value_input_option="RAW",
    )
    apply_receivable_row_format(sheet, insert_at, data["overdue_days"])
    refresh_receivable_totals(sheet)
    return True

def customer_analysis_stats(wb) -> dict:
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    tx_values = tx_sheet.get_all_values()
    stats = {}

    for row in tx_values[MIN_TRANSACTION_ROW - 1:]:
        if is_soft_deleted_row(row, COL_DELETED_AT):
            continue
        tx_type = row[COL_TYPE - 1] if len(row) >= COL_TYPE else ""
        if tx_type != TX_INCOME:
            continue

        customer = row[COL_CUSTOMER - 1].strip() if len(row) >= COL_CUSTOMER and row[COL_CUSTOMER - 1] else ""
        if not customer:
            continue

        amount_text = row[COL_AMOUNT - 1] if len(row) >= COL_AMOUNT else "0"
        date_text = row[COL_DATE - 1] if len(row) >= COL_DATE else ""
        try:
            amount = to_int(amount_text)
        except ValueError:
            amount = 0

        tx_date = parse_date_value(date_text) if date_text else None
        entry = stats.setdefault(customer, {
            "total": 0,
            "outstanding": 0,
            "count": 0,
            "recent": None,
        })
        entry["total"] += amount
        entry["count"] += 1
        if tx_date and (entry["recent"] is None or tx_date > entry["recent"]):
            entry["recent"] = tx_date

    try:
        recv_sheet = wb.worksheet(SHEET_RECEIVABLES)
        recv_values = recv_sheet.get_all_values()
        total_row = find_total_row(recv_values)
        end_index = (total_row - 1) if total_row else len(recv_values)
        for row in recv_values[MIN_TRANSACTION_ROW - 1:end_index]:
            if is_soft_deleted_row(row, RECV_COL_DELETED_AT):
                continue
            customer = row[RECV_COL_CUSTOMER - 1].strip() if len(row) >= RECV_COL_CUSTOMER and row[RECV_COL_CUSTOMER - 1] else ""
            if not customer:
                continue
            outstanding_text = row[RECV_COL_OUTSTANDING - 1] if len(row) >= RECV_COL_OUTSTANDING else "0"
            try:
                outstanding = to_int(outstanding_text) if outstanding_text else 0
            except ValueError:
                outstanding = 0
            entry = stats.setdefault(customer, {
                "total": 0,
                "outstanding": 0,
                "count": 0,
                "recent": None,
            })
            entry["outstanding"] += outstanding
    except Exception:
        logger.exception("Failed to read receivables while building customer analysis")

    return stats

def refresh_customer_analysis(wb) -> bool:
    sheet = worksheet_by_names(wb, SHEET_CUSTOMER_ANALYSIS, "客戶分析")
    values = sheet.get_all_values()
    stats = customer_analysis_stats(wb)

    existing_names = []
    for row in values[MIN_TRANSACTION_ROW - 1:]:
        name = row[CUST_COL_NAME - 1].strip() if len(row) >= CUST_COL_NAME and row[CUST_COL_NAME - 1] else ""
        if name and name not in existing_names:
            existing_names.append(name)

    all_names = existing_names[:]
    for name in sorted(stats.keys()):
        if name not in all_names:
            all_names.append(name)

    old_row_count = max(len(values), MIN_TRANSACTION_ROW - 1)
    if old_row_count >= MIN_TRANSACTION_ROW:
        sheet.batch_clear([f"A{MIN_TRANSACTION_ROW}:F{old_row_count}"])

    rows = []
    for name in all_names:
        entry = stats.get(name, {
            "total": 0,
            "outstanding": 0,
            "count": 0,
            "recent": None,
        })
        total = entry["total"]
        outstanding = entry["outstanding"]
        collected = max(total - outstanding, 0)
        recent = entry["recent"].strftime("%Y/%m/%d") if entry["recent"] else ""
        rows.append([name, total, collected, outstanding, entry["count"], recent])

    if rows:
        sheet.update(f"A{MIN_TRANSACTION_ROW}:F{MIN_TRANSACTION_ROW + len(rows) - 1}", rows, value_input_option="RAW")
        apply_alternating_row_colors_to(sheet, CUST_COL_RECENT, MIN_TRANSACTION_ROW + len(rows) - 1)
    return True

def update_customer_analysis(wb, data: dict | None = None) -> bool:
    if data is not None and data["type"] != TX_INCOME:
        return False
    return refresh_customer_analysis(wb)

def monthly_summary_empty_stats() -> dict:
    return {
        "income": 0,
        "cash": 0,
        "purchase": 0,
        "fixed": 0,
        "variable": 0,
        "expense": 0,
        "receivable": 0,
        "exchange": 0,
        "direct_income": 0,
        "export_income": 0,
        "max_transaction": 0,
        "count": 0,
        "collection_days": [],
    }

def format_rate(numerator: int | float, denominator: int | float) -> str:
    if not denominator:
        return "0.00%"
    return f"{(numerator / denominator * 100):.2f}%"

def rounded_amount(value: int | float) -> int | float:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return round(value, 2) if isinstance(value, float) else value

def refresh_monthly_overview(wb) -> bool:
    sheet = worksheet_by_names(wb, SHEET_MONTHLY_OVERVIEW, "月份總覽", "月分總覽", "月份財務總覽")
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    tx_values = tx_sheet.get_all_values()
    monthly = {month: monthly_summary_empty_stats() for month in range(1, 13)}

    for row in tx_values[MIN_TRANSACTION_ROW - 1:]:
        if is_soft_deleted_row(row, COL_DELETED_AT):
            continue
        date_text = row[COL_DATE - 1] if len(row) >= COL_DATE else ""
        tx_date = parse_date_value(str(date_text)) if date_text else None
        if not tx_date:
            continue

        month = tx_date.month
        entry = monthly[month]
        tx_type = row[COL_TYPE - 1] if len(row) >= COL_TYPE else ""
        amount_text = row[COL_AMOUNT - 1] if len(row) >= COL_AMOUNT else "0"
        status = row[COL_STATUS - 1] if len(row) >= COL_STATUS else ""
        pay_date_text = row[COL_PAY_DATE - 1] if len(row) >= COL_PAY_DATE else ""
        channel = row[COL_CHANNEL - 1] if len(row) >= COL_CHANNEL else ""
        cost_structure = row[COL_COST_STRUCT - 1] if len(row) >= COL_COST_STRUCT else ""
        days_text = row[COL_DAYS - 1] if len(row) >= COL_DAYS else ""

        try:
            amount = to_int(amount_text) if amount_text else 0
        except ValueError:
            amount = 0
        if amount <= 0:
            continue

        entry["max_transaction"] = max(entry["max_transaction"], amount)
        entry["count"] += 1

        if tx_type == TX_INCOME:
            entry["income"] += amount
            pay_date = parse_date_value(str(pay_date_text)) if pay_date_text else None
            if status in PAID_STATUSES and pay_date:
                monthly[pay_date.month]["cash"] += amount
            if channel == "大陸出口":
                entry["export_income"] += amount
            else:
                entry["direct_income"] += amount
            try:
                if days_text != "":
                    entry["collection_days"].append(to_number(days_text))
            except ValueError:
                pass
        elif tx_type == TX_EXPENSE:
            entry["expense"] += amount
            if cost_structure == "進貨成本":
                entry["purchase"] += amount
            elif cost_structure in FIXED_COST_STRUCTURES:
                entry["fixed"] += amount
            elif cost_structure == "換匯":
                entry["exchange"] += amount
            else:
                entry["variable"] += amount

    try:
        recv_sheet = wb.worksheet(SHEET_RECEIVABLES)
        recv_values = recv_sheet.get_all_values()
        total_row = find_total_row(recv_values)
        end_index = (total_row - 1) if total_row else len(recv_values)
        for row in recv_values[MIN_TRANSACTION_ROW - 1:end_index]:
            if is_soft_deleted_row(row, RECV_COL_DELETED_AT):
                continue
            date_text = row[RECV_COL_DATE - 1] if len(row) >= RECV_COL_DATE else ""
            recv_date = parse_date_value(str(date_text)) if date_text else None
            if not recv_date:
                continue
            try:
                outstanding = to_int(row[RECV_COL_OUTSTANDING - 1]) if len(row) >= RECV_COL_OUTSTANDING and row[RECV_COL_OUTSTANDING - 1] else 0
            except ValueError:
                outstanding = 0
            try:
                collected = to_int(row[RECV_COL_COLLECTED - 1]) if len(row) >= RECV_COL_COLLECTED and row[RECV_COL_COLLECTED - 1] else 0
            except ValueError:
                collected = 0
            monthly[recv_date.month]["receivable"] += outstanding
            if outstanding > 0 and collected > 0:
                monthly[recv_date.month]["cash"] += collected
    except Exception:
        logger.exception("Failed to read receivables while building monthly overview")

    rows = []
    all_days = []
    totals = monthly_summary_empty_stats()
    for month in range(1, 13):
        entry = monthly[month]
        all_days.extend(entry["collection_days"])
        for key in ("income", "cash", "purchase", "fixed", "variable", "expense",
                    "receivable", "exchange", "direct_income", "export_income", "count"):
            totals[key] += entry[key]
        totals["max_transaction"] = max(totals["max_transaction"], entry["max_transaction"])

        gross_profit = entry["income"] - entry["purchase"]
        net_profit = entry["income"] - entry["expense"]
        avg_days = round(sum(entry["collection_days"]) / len(entry["collection_days"]), 1) if entry["collection_days"] else 0
        rows.append([
            f"{month}月",
            entry["income"],
            entry["cash"],
            entry["purchase"],
            entry["fixed"],
            entry["variable"],
            entry["expense"],
            gross_profit,
            format_rate(gross_profit, entry["income"]),
            net_profit,
            entry["receivable"],
            entry["cash"],
            format_rate(entry["cash"], entry["income"]),
            avg_days,
            entry["exchange"],
            entry["direct_income"],
            entry["export_income"],
            entry["max_transaction"],
            entry["count"],
        ])

    total_gross_profit = totals["income"] - totals["purchase"]
    total_net_profit = totals["income"] - totals["expense"]
    total_avg_days = round(sum(all_days) / len(all_days), 1) if all_days else 0
    rows.append([
        "全年合計",
        totals["income"],
        totals["cash"],
        totals["purchase"],
        totals["fixed"],
        totals["variable"],
        totals["expense"],
        total_gross_profit,
        format_rate(total_gross_profit, totals["income"]),
        total_net_profit,
        totals["receivable"],
        totals["cash"],
        format_rate(totals["cash"], totals["income"]),
        total_avg_days,
        totals["exchange"],
        totals["direct_income"],
        totals["export_income"],
        totals["max_transaction"],
        totals["count"],
    ])

    sheet.update(
        range_a1(MIN_TRANSACTION_ROW, 1, MIN_TRANSACTION_ROW + len(rows) - 1, MONTHLY_LAST_COL),
        rows,
        value_input_option="RAW",
    )
    apply_alternating_row_colors_to(sheet, MONTHLY_LAST_COL, MIN_TRANSACTION_ROW + 11)
    return True

# ══════════════════════════════════════════════════════════════
# 回覆格式
# ══════════════════════════════════════════════════════════════
def format_new_transaction_reply(data: dict, tx_id: str, recv_added: bool, customer_added: bool = False) -> str:
    icon        = "💰" if data["type"] == TX_INCOME else "💸"
    status_icon = {"已收":"✅","已付":"✅","未付":"⏳","部分收":"⚠️"}.get(data["status"], "")
    lines = [
        f"{icon} 已記錄交易 {tx_id}",
        "─" * 22,
        f"編號｜{tx_id}",
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
        lines.append(f"數量｜{data['qty']:,}")
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
    if data["exchange_rate"]:
        lines.append(f"匯率｜{data['exchange_rate']}")
    if data["due_date"]:
        lines.append(f"付款期限｜{data['due_date']}")
    if data["initial_collected"]:
        lines.append(f"已收｜NT$ {data['initial_collected']:,}")
        lines.append(f"未收｜NT$ {data['outstanding']:,}")
    if data["overdue_days"]:
        lines.append(f"逾期｜{data['overdue_days']} 天")
    lines.append("─" * 22)
    if recv_added:
        lines.append("📌 已同步到應收帳款")
    if customer_added:
        lines.append("👥 已新增到客戶分析")
    lines.append("✏️ 更新付款狀態請輸入：")
    lines.append(f"更新 {tx_id} 已收")
    return "\n".join(lines)

def format_update_reply(result: dict) -> str:
    status_icon = {"已收":"✅","已付":"✅","部分收":"⚠️"}.get(result["new_status"], "")
    lines = [
        f"✅ 交易 {result['tx_id']} 已更新",
        "─" * 22,
        f"編號｜{result['tx_id']}",
        f"品項｜{result['item']}",
    ]
    if result["customer"]:
        lines.append(f"客戶｜{result['customer']}")
    lines += [
        f"狀態｜{status_icon} {result['new_status']}",
        f"付款日｜{result['pay_date']}",
    ]
    if result["received_now"]:
        lines.append(f"本次收款｜NT$ {result['received_now']:,}")
    if result["collected"]:
        lines.append(f"累計已收｜NT$ {result['collected']:,} / NT$ {result['total_amount']}")
    if result["outstanding"] != "":
        lines.append(f"未收餘額｜NT$ {result['outstanding']:,}")
    if result["days"] != "":
        lines.append(f"收款天數｜{result['days']} 天")
    lines.append("─" * 22)
    if result["recv_updated"]:
        lines.append("📌 應收帳款已同步更新")
    elif result.get("type") == TX_INCOME:
        lines.append("⚠️ 應收帳款請手動確認")
    return "\n".join(lines)

def format_delete_reply(result: dict) -> str:
    lines = [
        f"🗑️ 交易 {result['tx_id']} 已標記刪除",
        "─" * 22,
        f"編號｜{result['tx_id']}",
        f"刪除日期｜{result['deleted_at']}",
        f"類型｜{result['type']}",
        f"金額｜NT$ {result['amount']:,}",
        f"品項｜{result['item']}",
    ]
    if result["customer"]:
        lines.append(f"客戶｜{result['customer']}")
    if result["recv_deleted"]:
        lines.append("📌 對應應收帳款已一併標記刪除")
    lines.append("30 天後會由「整理」或每日排程自動移除")
    lines.append(f"若要復原請輸入：恢復 {result['tx_id']}")
    return "\n".join(lines)

def format_restore_reply(result: dict) -> str:
    lines = [
        f"↩️ 交易 {result['tx_id']} 已恢復",
        "─" * 22,
        f"編號｜{result['tx_id']}",
        f"類型｜{result['type']}",
        f"金額｜NT$ {result['amount']:,}",
        f"品項｜{result['item']}",
    ]
    if result["customer"]:
        lines.append(f"客戶｜{result['customer']}")
    if result["recv_restored"]:
        lines.append("📌 對應應收帳款已一併恢復")
    lines.append("報表會在「整理」或每日排程後重新納入統計")
    return "\n".join(lines)

HELP_TEXT = """溫室帳目機器人

【紀錄收入】
請使用按鈕模板或標籤格式：
記收入
日期：7/1
類型：收入
金額：50000
品項：珍妮
數量：500棵
售出單價：100
進貨單價：30
客戶：李淵男
狀態：未付
期限：7/31

【紀錄支出】
請使用按鈕模板或標籤格式：
記支出
日期：7/1
類型：支出
金額：3000
品項：罰款
數量：
廠商：
狀態：已付
期限：

【可選欄位說明】
付款狀態：已收 / 已付 / 未付 / 部分收
日期留空：預設今天
收入期限留空：預設交易日起1個月
部分收款：已收：30000
數量：數量：500棵 / 數量：500
售出單價：售出單價：100
進貨單價：進貨單價：30
外匯：人民幣：4000 / RMB：4000

可用標籤：日期、類型、金額、品項、數量、售出單價、進貨單價、客戶、廠商、狀態、期限、人民幣、已收、備註
標籤格式必填：類型、金額、品項
舊格式如「收入 50000 品項 客戶」已停用

【按鈕文字】
記收入 / 記支出 / 更新收款：只作為輸入模板，不會回覆也不會記帳
整理：執行整理
刪除 TX-0001：軟刪除
恢復 TX-0001：恢復軟刪除

【自動規則】
未輸入日期：預設今天
收入未輸入狀態：預設未付
支出未輸入狀態：預設已付
已收/已付：付款日期預設為交易日期
未付/部分收收入：自動同步應收帳款
客戶分析與月份總覽：輸入「整理」或每日排程時更新
每筆交易會自動產生固定交易編號，例如 TX-0001
更新與刪除請用交易編號，排序後也不會更新錯筆

【更新付款狀態】
更新 交易編號 已收
更新 交易編號 已付
更新 交易編號 部分收 本次收款金額

更新範例：
更新 TX-0021 已收
更新 TX-0021 已付
更新 TX-0021 部分收 30000
更新 21 已收
部分收會累加到既有已收金額；累計達全額時會自動結清

【刪除交易】
刪除 交易編號
刪除交易 交易編號
刪除交易紀錄 交易編號

刪除範例：
刪除 TX-0025
刪除交易 TX-0025
刪除 25

刪除會先標成灰色並填入刪除日期，30天後由「整理」或每日排程移除。
30天內可輸入：恢復 TX-0025"""

def refresh_receivables_job() -> dict:
    wb = get_workbook()
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    sheet = wb.worksheet(SHEET_RECEIVABLES)
    ids_added = ensure_transaction_ids(tx_sheet)
    purged_transactions = purge_soft_deleted_rows(tx_sheet, COL_DELETED_AT)
    purged_receivables = purge_soft_deleted_rows(sheet, RECV_COL_DELETED_AT)
    sort_sheet_by_date(tx_sheet, TX_LAST_COL)
    refresh_transaction_formats(tx_sheet)
    sort_sheet_by_date(sheet, RECV_LAST_COL)
    refresh_receivable_overdue_formats(sheet)
    refresh_receivable_totals(sheet)
    refresh_customer_analysis(wb)
    refresh_monthly_overview(wb)
    return {
        "ok": True,
        "date": today_tw().strftime("%Y/%m/%d"),
        "ids_added": ids_added,
        "purged_transactions": purged_transactions,
        "purged_receivables": purged_receivables,
    }

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

@app.route("/refresh-receivables", methods=["POST", "GET"])
def refresh_receivables_route():
    expected_secret = CRON_SECRET
    provided_secret = (
        request.headers.get("X-Cron-Secret")
        or request.args.get("secret")
        or ""
    )
    if expected_secret and provided_secret != expected_secret:
        abort(403)
    if not expected_secret:
        logger.warning("CRON_SECRET is not set; /refresh-receivables is unprotected")

    try:
        result = refresh_receivables_job()
        return {
            "status": "ok",
            "date": result["date"],
            "ids_added": result.get("ids_added", 0),
            "purged_transactions": result.get("purged_transactions", 0),
            "purged_receivables": result.get("purged_receivables", 0),
        }
    except Exception:
        logger.exception("Scheduled receivables refresh failed")
        return {"status": "error"}, 500

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_text   = event.message.text.strip()
    reply_token = event.reply_token
    message_id  = getattr(event.message, "id", None)

    if should_ignore_template_prompt(user_text):
        return

    with ApiClient(configuration) as api_client:
        line_api = MessagingApi(api_client)

        # ── 說明指令 ───────────────────────────────────────────
        if user_text in ("說明", "help", "Help", "HELP", "?", "？"):
            reply = HELP_TEXT

        # ── 手動刷新應收帳款 ───────────────────────────────────
        elif user_text == "整理":
            try:
                result = refresh_receivables_job()
                reply = (
                    "✅ 交易格式、帳款、月份總覽與客戶分析已整理\n"
                    f"日期｜{result['date']}\n"
                    f"補上交易編號｜{result.get('ids_added', 0)} 筆\n"
                    f"移除逾期刪除交易｜{result.get('purged_transactions', 0)} 筆\n"
                    f"移除逾期刪除應收｜{result.get('purged_receivables', 0)} 筆"
                )
            except Exception:
                logger.exception("Manual receivables refresh failed")
                reply = "⚠️ 整理失敗，請稍後再試。"

        # ── 更新付款狀態 ────────────────────────────────────────
        elif user_text.startswith("更新"):
            update_cmd = parse_update_command(user_text)
            if update_cmd is None:
                reply = (
                    "❌ 更新格式錯誤\n\n"
                    "正確格式：\n"
                    "更新 交易編號 已收\n"
                    "更新 交易編號 已付\n"
                    "更新 交易編號 部分收 金額\n\n"
                    "範例：更新 TX-0021 已收\n"
                    "　　　更新 TX-0021 部分收 30000\n"
                    "也可簡寫：更新 21 已收"
                )
            else:
                try:
                    wb     = get_workbook()
                    result = apply_status_update(
                        wb,
                        update_cmd["tx_id"],
                        update_cmd["status"],
                        update_cmd["collected"],
                    )
                    reply = format_update_reply(result) if result["ok"] else f"❌ {result['error']}"
                except Exception:
                    logger.exception("Update failed")
                    reply = "⚠️ 更新失敗，請稍後再試。"

        # ── 恢復軟刪除交易 ──────────────────────────────────────
        elif user_text.startswith("恢復"):
            restore_cmd = parse_restore_command(user_text)
            if restore_cmd is None:
                reply = (
                    "❌ 恢復格式錯誤\n\n"
                    "正確格式：\n"
                    "恢復 交易編號\n"
                    "恢復交易 交易編號\n\n"
                    "範例：恢復 TX-0025\n"
                    "也可簡寫：恢復 25"
                )
            else:
                try:
                    wb = get_workbook()
                    result = apply_restore_transaction(wb, restore_cmd["tx_id"])
                    reply = format_restore_reply(result) if result["ok"] else f"❌ {result['error']}"
                except Exception:
                    logger.exception("Restore failed")
                    reply = "⚠️ 恢復失敗，請稍後再試。"

        # ── 刪除交易紀錄 ────────────────────────────────────────
        elif user_text.startswith("刪除"):
            delete_cmd = parse_delete_command(user_text)
            if delete_cmd is None:
                reply = (
                    "❌ 刪除格式錯誤\n\n"
                    "正確格式：\n"
                    "刪除 交易編號\n"
                    "刪除交易 交易編號\n\n"
                    "範例：刪除 TX-0025\n"
                    "也可簡寫：刪除 25"
                )
            else:
                try:
                    wb = get_workbook()
                    result = apply_delete_transaction(wb, delete_cmd["tx_id"])
                    reply = format_delete_reply(result) if result["ok"] else f"❌ {result['error']}"
                except Exception:
                    logger.exception("Delete failed")
                    reply = "⚠️ 刪除失敗，請稍後再試。"

        # ── 新增交易（防重複） ──────────────────────────────────
        elif has_seen_message(message_id):
            reply = "這則訊息已經處理過，沒有重複寫入。"

        else:
            parsed = parse_message(user_text)
            if parsed is None:
                reply = (
                    "❌ 無法解析此訊息\n\n"
                    "新增交易請使用標籤格式：\n"
                    "記收入\n"
                    "日期：7/1\n"
                    "類型：收入\n"
                    "金額：50000\n"
                    "品項：珍妮\n"
                    "數量：500棵\n"
                    "售出單價：100\n"
                    "進貨單價：30\n"
                    "客戶：李淵男\n\n"
                    "更新收款請輸入：更新 TX-0001 已收\n\n"
                    "輸入「說明」查看完整格式"
                )
            else:
                try:
                    wb      = get_workbook()
                    row_num, tx_id = append_transaction(wb, parsed)
                    remember_message(message_id)
                except Exception:
                    logger.exception("Failed to append transaction")
                    reply = "⚠️ 寫入交易記錄時發生錯誤，請稍後再試。"
                else:
                    recv_added = False
                    recv_err   = False
                    try:
                        recv_added = update_receivables(wb, parsed, tx_id)
                    except Exception:
                        recv_err = True
                        logger.exception("Failed to update receivables")

                    reply = format_new_transaction_reply(parsed, tx_id, recv_added)
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
