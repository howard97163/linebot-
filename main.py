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
COL_WHOLESALE_PRICE = 10  # 售出批發價
COL_COST        = 11  # 進貨單價
COL_PROFIT      = 12  # 毛利
COL_CHANNEL     = 13  # 銷售管道
COL_DAYS        = 14  # 收款天數
COL_NOTE        = 15  # 備註
COL_CATEGORY    = 16  # 分類
COL_COST_STRUCT = 17  # 支出成本結構
COL_MONTH       = 18  # 月份
COL_RMB         = 19  # 換匯金額
COL_EXCHANGE_RATE = 20  # 匯率
COL_RAW         = 21  # 原始備註

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
TX_LAST_COL = COL_RAW
RECV_LAST_COL = RECV_COL_NOTE
MONTHLY_LAST_COL = 19
ROW_COLOR_WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
ROW_COLOR_GREEN = {"red": 0.91, "green": 0.97, "blue": 0.94}
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

def apply_receivable_overdue_format(sheet, row_num: int, overdue_days):
    color = overdue_background_color(overdue_days)
    if not color:
        return
    sheet.format(f"A{row_num}:I{row_num}", {"backgroundColor": color})

def alternating_color_for_row(row_num: int):
    return ROW_COLOR_GREEN if (row_num - MIN_TRANSACTION_ROW) % 2 else ROW_COLOR_WHITE

def clear_receivable_overdue_format(sheet, row_num: int):
    sheet.format(
        f"A{row_num}:I{row_num}",
        {"backgroundColor": alternating_color_for_row(row_num)},
    )

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
    sheet.batch_update([
        {"range": f"D{total_row}", "values": [[f"=SUM(D{MIN_TRANSACTION_ROW}:D{last_detail_row})"]]},
        {"range": f"E{total_row}", "values": [[f"=SUM(E{MIN_TRANSACTION_ROW}:E{last_detail_row})"]]},
        {"range": f"F{total_row}", "values": [[f"=SUM(F{MIN_TRANSACTION_ROW}:F{last_detail_row})"]]},
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
    for row_num in range(MIN_TRANSACTION_ROW, end_row + 1):
        row = values[row_num - 1] if len(values) >= row_num else []
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
                apply_receivable_overdue_format(sheet, row_num, existing_overdue_days)
            continue
        overdue_days = calculate_overdue_days(due_date, outstanding)
        if overdue_days:
            sheet.update_cell(row_num, RECV_COL_OVERDUE, overdue_days)
            apply_receivable_overdue_format(sheet, row_num, overdue_days)
        else:
            clear_receivable_overdue_format(sheet, row_num)
    refresh_receivable_totals(sheet)

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
    "711":"7-11活動","福袋":"福袋活動",
    # 進出口
    "檢疫":"檢疫","關稅":"關稅",
    "罰金":"罰金","出口":"出口費用",
    "標籤帶":"標籤帶",
    # 差旅
    "機票":"機票","住宿":"住宿","計程車":"計程車",
    "接機":"接機","機加酒":"機加酒",
    # 餐飲交際
    "晚餐":"晚餐","尾牙":"尾牙","吃飯":"吃飯",
    # 平台費用
    "蝦皮手續費":"蝦皮手續費","手續費":"手續費","服務費":"服務費",
    "刷卡":"刷卡費","金流":"金流費","平台抽成":"平台抽成",
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
    "7-11活動":"活動銷售","福袋活動":"活動銷售",
    "檢疫":"進出口費用","關稅":"進出口費用","罰金":"進出口費用",
    "出口費用":"進出口費用","標籤帶":"進出口費用",
    "機票":"差旅費","住宿":"差旅費","計程車":"差旅費","接機":"差旅費","機加酒":"差旅費",
    "晚餐":"餐飲交際","尾牙":"餐飲交際","吃飯":"餐飲交際",
    "蝦皮手續費":"平台費用","手續費":"平台費用","服務費":"平台費用",
    "刷卡費":"平台費用","金流費":"平台費用","平台抽成":"平台費用",
    "退款":"損耗退貨",
    "廣告":"行銷廣告","廣告投放":"行銷廣告","拍攝":"行銷廣告",
    "設計":"行銷廣告","印刷":"行銷廣告","名片":"行銷廣告",
    "轉帳費":"銀行費用","匯費":"銀行費用","銀行手續費":"銀行費用","利息":"銀行費用",
    "營業稅":"稅務規費","所得稅":"稅務規費","規費":"稅務規費","牌照":"稅務規費",
    "攤位費":"活動費用","市集活動":"活動費用","展覽":"活動費用","報名費":"活動費用",
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
    m = re.search(rf'\$({NUMBER_PATTERN})\s*[*×x]\s*(\d[\d,]*)', item)
    if m:
        return to_int(m.group(2)), to_number(m.group(1))
    m = re.search(rf'(\d[\d,]*)\s*[*×x]\s*\$?({NUMBER_PATTERN})', item)
    if m:
        qty = to_int(m.group(1))
        return qty, to_number(m.group(2))
    m = re.search(r'(\d[\d,]*)\s*[顆盒個株]', item)
    if m:
        qty = to_int(m.group(1))
        return qty, (round(amount / qty) if qty > 0 else 0)
    return 0, 0

def extract_labeled_customer(text: str) -> tuple[str, str]:
    m = re.search(r"\s+(?:客戶|廠商|供應商|對象)\s*[:：]\s*(.+)$", text)
    if not m:
        return text, ""
    customer = m.group(1).strip()
    text = text[:m.start()].strip()
    return text, customer

def extract_wholesale_price(text: str) -> tuple[str, float | int | str]:
    pattern = rf"(?:售出批發價|批發售價|批發價|批價)\s*[:：]?\s*({NUMBER_PATTERN})"
    m = re.search(pattern, text)
    if not m:
        return text, ""
    wholesale_price = to_number(m.group(1))
    text = (text[:m.start()] + " " + text[m.end():]).strip()
    return re.sub(r"\s+", " ", text), wholesale_price

def clean_item_text(text: str) -> str:
    text = re.sub(r'(\d[\d,]*)\s*rmb|rmb\s*(\d[\d,]*)|人民幣\s*(\d[\d,]*)', ' ', text, flags=re.IGNORECASE)
    text = re.sub(rf'(?:total|總額)\s*[:：]?\s*{NUMBER_PATTERN}', ' ', text, flags=re.IGNORECASE)
    text = re.sub(rf'(?:售出批發價|批發售價|批發價|批價)\s*[:：]?\s*{NUMBER_PATTERN}', ' ', text)
    text = re.sub(rf'(?:進價|單價|一袋成本|成本)\s*{NUMBER_PATTERN}', ' ', text)
    text = re.sub(r'\d[\d,]*\s*[顆盒個株袋]', ' ', text)
    return re.sub(r"\s+", " ", text).strip()

# ══════════════════════════════════════════════════════════════
# 新增交易：解析訊息
# ══════════════════════════════════════════════════════════════
def parse_message(text: str) -> dict | None:
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

    remaining, wholesale_price = extract_wholesale_price(remaining)
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
    has_detail_parts = bool(cost_per_unit or wholesale_price or rmb or qty)
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
        "wholesale_price": wholesale_price,
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

def parse_delete_command(text: str) -> dict | None:
    m = re.match(r"^刪除(?:交易(?:紀錄|記錄)?)?\s+(\d+)$", text.strip())
    if not m:
        return None
    return {"row": int(m.group(1))}

def receivable_note(raw: str, row_num: int) -> str:
    return f"[交易行號:{row_num}] {raw}".strip()

def find_receivable_row(all_recv: list[list[str]], row_num: int, customer: str, item: str, amount: int, raw: str = "") -> int | None:
    marker = f"[交易行號:{row_num}]"
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
        recv_row = find_receivable_row(all_recv, row_num, customer, item, total_amount, raw)

        if tx_type == TX_INCOME and recv_row:
            recv_data = all_recv[recv_row - 1] if len(all_recv) >= recv_row else []
            due_date_str = recv_data[RECV_COL_DUE - 1] if len(recv_data) >= RECV_COL_DUE else ""
            due_date = parse_date_value(due_date_str) if due_date_str else None
            if new_status == STATUS_COLLECTED:
                # 全額已收
                recv_sheet.batch_update([
                    {"range": f"E{recv_row}", "values": [[total_amount]]},
                    {"range": f"F{recv_row}", "values": [[0]]},
                    {"range": f"I{recv_row}", "values": [[receivable_note(raw, row_num)]]},
                ])
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
                    {"range": f"I{recv_row}", "values": [[receivable_note(raw, row_num)]]},
                ]
                if outstanding > 0:
                    receivable_updates.append({"range": f"H{recv_row}", "values": [[overdue_days]]})
                recv_sheet.batch_update(receivable_updates)
                if outstanding > 0:
                    if overdue_days:
                        apply_receivable_overdue_format(recv_sheet, recv_row, overdue_days)
                    else:
                        clear_receivable_overdue_format(recv_sheet, recv_row)
                recv_updated = True
        refresh_receivable_overdue_formats(recv_sheet)
    except Exception:
        logger.exception("Failed to update receivables during status update")
    if tx_type == TX_INCOME:
        try:
            refresh_customer_analysis(wb)
        except Exception:
            logger.exception("Failed to refresh customer analysis during status update")
    try:
        refresh_monthly_overview(wb)
    except Exception:
        logger.exception("Failed to refresh monthly overview during status update")

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
        "collected":    total_collected,
        "received_now": collected,
        "outstanding":  outstanding,
        "total_amount": amount_str,
    }

def apply_delete_transaction(wb, row_num: int) -> dict:
    if row_num < MIN_TRANSACTION_ROW:
        return {"ok": False, "error": "請確認行號是否為交易資料列"}

    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    row_data = tx_sheet.row_values(row_num)
    if not row_data or len(row_data) < COL_AMOUNT:
        return {"ok": False, "error": f"找不到第 {row_num} 行，請確認行號是否正確"}

    tx_type = row_data[COL_TYPE - 1] if len(row_data) >= COL_TYPE else ""
    item = row_data[COL_ITEM - 1] if len(row_data) >= COL_ITEM else ""
    customer = row_data[COL_CUSTOMER - 1] if len(row_data) >= COL_CUSTOMER else ""
    amount_str = row_data[COL_AMOUNT - 1] if len(row_data) >= COL_AMOUNT else "0"
    raw = row_data[COL_RAW - 1] if len(row_data) >= COL_RAW else ""
    try:
        amount = to_int(amount_str) if amount_str else 0
    except ValueError:
        return {"ok": False, "error": "此列金額格式不正確，請先檢查試算表"}

    if tx_type not in (TX_INCOME, TX_EXPENSE):
        return {"ok": False, "error": "此列不是有效的交易資料列"}

    recv_deleted = False
    recv_row = None
    if tx_type == TX_INCOME:
        try:
            recv_sheet = wb.worksheet(SHEET_RECEIVABLES)
            all_recv = recv_sheet.get_all_values()
            recv_row = find_receivable_row(all_recv, row_num, customer, item, amount, raw)
            if recv_row:
                recv_sheet.delete_rows(recv_row)
                recv_deleted = True
        except Exception:
            logger.exception("Failed to delete related receivable")
            return {"ok": False, "error": "找到交易資料，但刪除對應應收帳款時失敗，請稍後再試"}

    tx_sheet.delete_rows(row_num)
    sort_sheet_by_date(tx_sheet, TX_LAST_COL)
    refresh_transaction_formats(tx_sheet)
    if tx_type == TX_INCOME:
        try:
            refresh_receivable_overdue_formats(wb.worksheet(SHEET_RECEIVABLES))
        except Exception:
            logger.exception("Failed to refresh receivables after deletion")
        try:
            refresh_customer_analysis(wb)
        except Exception:
            logger.exception("Failed to refresh customer analysis after deletion")
    try:
        refresh_monthly_overview(wb)
    except Exception:
        logger.exception("Failed to refresh monthly overview after deletion")

    return {
        "ok": True,
        "row_num": row_num,
        "type": tx_type,
        "item": item,
        "customer": customer,
        "amount": amount,
        "recv_deleted": recv_deleted,
        "recv_row": recv_row,
    }

# ══════════════════════════════════════════════════════════════
# 寫入新交易
# ══════════════════════════════════════════════════════════════
def append_transaction(wb, data: dict) -> int:
    sheet = wb.worksheet(SHEET_TRANSACTIONS)
    row = [
        data["date"], data["type"], data["amount"], data["item"],
        data["customer"], data["status"], data["pay_date"],
        data["qty"], data["unit_price"], data["wholesale_price"], data["cost_per_unit"],
        data["gross_profit"], data["channel"], data["days_to_collect"],
        data["note"], data["category"], data["cost_structure"],
        data["month"], data["rmb"], data["exchange_rate"], data["raw"],
    ]
    sheet.append_row(row, value_input_option="RAW", insert_data_option="INSERT_ROWS")
    sorted_row_num = organize_transaction_sheet(sheet, data)
    return sorted_row_num if sorted_row_num else len(sheet.col_values(1))

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
         data["amount"], data["initial_collected"], data["outstanding"],
         data["due_date"], data["overdue_days"], receivable_note(data["raw"], row_num)],
        insert_at,
        value_input_option="RAW",
    )
    sort_sheet_by_date(sheet, RECV_LAST_COL)
    refresh_receivable_overdue_formats(sheet)
    return True

def customer_analysis_stats(wb) -> dict:
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    tx_values = tx_sheet.get_all_values()
    stats = {}

    for row in tx_values[MIN_TRANSACTION_ROW - 1:]:
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
def format_new_transaction_reply(data: dict, row_num: int, recv_added: bool, customer_added: bool = False) -> str:
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
    if data["wholesale_price"]:
        lines.append(f"批發售價｜NT$ {data['wholesale_price']:,} / 顆")
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
        f"🗑️ 已刪除第 {result['row_num']} 行交易",
        "─" * 22,
        f"類型｜{result['type']}",
        f"金額｜NT$ {result['amount']:,}",
        f"品項｜{result['item']}",
    ]
    if result["customer"]:
        lines.append(f"客戶｜{result['customer']}")
    if result["recv_deleted"]:
        lines.append("📌 對應應收帳款已一併刪除")
    return "\n".join(lines)

HELP_TEXT = """溫室帳目機器人

【紀錄收入】
基本格式：
收入 金額 品項 客戶 [付款狀態] [期限日期] [進價] [批發價]

可用簡寫：
+金額 品項 客戶 [付款狀態]

收入範例：
收入 50000 珍妮500顆 李淵男
收入 32000 爆米花1000顆 美琪 未付
收入 50000 珍妮500顆 李淵男 已收
收入 50000 珍妮500顆 李淵男 未付 期限7/15
收入 50000 珍妮500顆 李淵男 部分收 30000
收入 50000 珍妮500顆 李淵男 進價30
收入 20000 鹿角蕨40顆 姜孟學 批發價450 已收
+50000 侏儒黃月1000顆 吳政翰

【紀錄支出】
基本格式：
支出 金額 品項 [對象/備註] [付款狀態]

可用簡寫：
-金額 品項 [對象/備註]

支出範例：
支出 18000 大陸運費
支出 400 測試 姜孟學 已付
支出 18114 換人民幣4000 黃浩
-4200 淘寶悶箱

【補登以前日期】
日期可放在最前面：
2026/05/29 收入 20000 鹿角蕨40顆 姜孟學 未付
2026-05-29 支出 400 測試 姜孟學 已付
5/29 收入 20000 鹿角蕨40顆 姜孟學
5-29 支出 400 測試

【可選欄位說明】
付款狀態：已收 / 已付 / 未付 / 部分收
期限日期：期限7/15 或 付款期限2026/07/15
部分收款：部分收 30000
進價：進價30
批發售價：批發價120 / 批發售價120 / 批價120
外匯：人民幣4000 / RMB4000 / 4000 RMB

【自動規則】
未輸入日期：預設今天
收入未輸入狀態：預設未付
支出未輸入狀態：預設已付
已收/已付：付款日期預設為交易日期
未付/部分收收入：自動同步應收帳款
收入的新客戶：自動加入客戶分析

【更新付款狀態】
更新 行號 已收
更新 行號 已付
更新 行號 部分收 本次收款金額

更新範例：
更新 21 已收
更新 21 已付
更新 21 部分收 30000
部分收會累加到既有已收金額；累計達全額時會自動結清

【刪除交易】
刪除 行號
刪除交易 行號
刪除交易紀錄 行號

刪除範例：
刪除 25
刪除交易 25
刪除交易紀錄 25"""

def refresh_receivables_job() -> dict:
    wb = get_workbook()
    tx_sheet = wb.worksheet(SHEET_TRANSACTIONS)
    sheet = wb.worksheet(SHEET_RECEIVABLES)
    refresh_transaction_formats(tx_sheet)
    refresh_receivable_overdue_formats(sheet)
    refresh_customer_analysis(wb)
    refresh_monthly_overview(wb)
    return {
        "ok": True,
        "date": today_tw().strftime("%Y/%m/%d"),
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
        }
    except Exception:
        logger.exception("Scheduled receivables refresh failed")
        return {"status": "error"}, 500

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

        # ── 手動刷新應收帳款 ───────────────────────────────────
        elif user_text == "整理":
            try:
                result = refresh_receivables_job()
                reply = f"✅ 交易格式、帳款、月份總覽與客戶分析已整理\n日期｜{result['date']}"
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

        # ── 刪除交易紀錄 ────────────────────────────────────────
        elif user_text.startswith("刪除"):
            delete_cmd = parse_delete_command(user_text)
            if delete_cmd is None:
                reply = (
                    "❌ 刪除格式錯誤\n\n"
                    "正確格式：\n"
                    "刪除 行號\n"
                    "刪除交易 行號\n\n"
                    "範例：刪除 25"
                )
            else:
                try:
                    wb = get_workbook()
                    result = apply_delete_transaction(wb, delete_cmd["row"])
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
                    customer_added = False
                    customer_err = False
                    monthly_err = False
                    try:
                        recv_added = update_receivables(wb, parsed, row_num)
                    except Exception:
                        recv_err = True
                        logger.exception("Failed to update receivables")
                    try:
                        customer_added = update_customer_analysis(wb, parsed)
                    except Exception:
                        customer_err = True
                        logger.exception("Failed to update customer analysis")
                    try:
                        refresh_monthly_overview(wb)
                    except Exception:
                        monthly_err = True
                        logger.exception("Failed to refresh monthly overview")

                    reply = format_new_transaction_reply(parsed, row_num, recv_added, customer_added)
                    if recv_err:
                        reply += "\n⚠️ 交易已記錄，但應收帳款同步失敗。"
                    if customer_err:
                        reply += "\n⚠️ 交易已記錄，但客戶分析同步失敗。"
                    if monthly_err:
                        reply += "\n⚠️ 交易已記錄，但月份總覽同步失敗。"

        line_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=reply)]
            )
        )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
