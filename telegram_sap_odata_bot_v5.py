import os
import time
import logging
from datetime import datetime

import openpyxl
import requests
import telebot
import urllib3
from dotenv import load_dotenv
from openai import OpenAI
from requests.auth import HTTPBasicAuth

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ALLOWED_CHAT_ID = int(os.getenv("ALLOWED_CHAT_ID", "0"))
SAP_BASE_URL = os.getenv("SAP_BASE_URL", "").strip()
SAP_USERNAME = os.getenv("SAP_USERNAME", "").strip()
SAP_PASSWORD = os.getenv("SAP_PASSWORD", "").strip()
EXCEL_PATH = os.getenv("EXCEL_PATH", "").strip()
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1").strip()
DEFAULT_UOM = os.getenv("DEFAULT_UOM", "EA").strip()
LOG_FILE = os.getenv("LOG_FILE", os.path.join(BASE_DIR, "telegram_sap_odata_bot.log")).strip()
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "").strip()
GROQ_MODEL = os.getenv("GROQ_MODEL", "llama-3.1-8b-instant").strip()

REQUEST_TIMEOUT = 60

if ":" not in BOT_TOKEN:
    raise Exception("BOT_TOKEN in .env is invalid.")

if not SAP_BASE_URL.startswith("http"):
    raise Exception("SAP_BASE_URL in .env is invalid.")

if not GROQ_API_KEY:
    raise Exception("GROQ_API_KEY in .env is missing.")

client = OpenAI(
    api_key=GROQ_API_KEY,
    base_url="https://api.groq.com/openai/v1"
)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

bot = telebot.TeleBot(BOT_TOKEN)
AUTH = HTTPBasicAuth(SAP_USERNAME, SAP_PASSWORD)
user_sessions = {}

MANUAL_ORDER_FIELDS = [
    ("order_type", "📋 Step 1/8 — Order Type\nEnter Order Type (e.g. OR):"),
    ("sales_org", "🏢 Step 2/8 — Sales Organization\nEnter Sales Organization (e.g. 1010):"),
    ("dist_channel", "📦 Step 3/8 — Distribution Channel\nEnter Distribution Channel (e.g. 10):"),
    ("division", "🗂️ Step 4/8 — Division\nEnter Division (e.g. 00):"),
    ("sold_to", "👤 Step 5/8 — Sold-to Party\nEnter Sold-to Party (Customer No.):"),
    ("ship_to", "🚚 Step 6/8 — Ship-to Party\nEnter Ship-to Party (or type SKIP to use Sold-to):"),
    ("cust_ref", "📝 Step 7/8 — Customer Reference\nEnter Customer PO / Reference No. (or SKIP):"),
    ("cust_ref_date", "📅 Step 8/8 — Customer Reference Date\nEnter date (DD.MM.YYYY) or SKIP:"),
]

LINE_ITEM_FIELDS = [
    ("material", "🔩 Line Item {item_no}/2 — Material\nEnter Material Number:"),
    ("quantity", "🔢 Line Item {item_no}/2 — Order Quantity\nEnter Order Quantity (numbers only):"),
]

INTENT_PROMPT = """
Classify the user's message into exactly one label from this list:

start_help
process_pending
retry_errors
send_excel
bot_status
stop_bot
new_order
unknown

Rules:
- Reply with exactly one label only.
- No explanation.
- No punctuation.
- No quotes.

Meaning:
- start_help: greeting, hi, hello, start, menu, help, what can you do
- process_pending: create sales orders from excel, process orders from excel, run orders from excel, create pending orders
- retry_errors: retry failed rows, retry errors, process failed records
- send_excel: send excel, send file, download excel, share workbook
- bot_status: show status, bot status, configuration, current setup
- stop_bot: stop bot, exit, close bot, shut down bot
- new_order: create new order, new sales order, manual order, enter order, create order manually, single order
- unknown: anything else

User message:
{message}
"""


def log_info(msg):
    logging.info(msg)
    print(msg)


def log_error(msg):
    logging.error(msg)
    print(msg)


def safe_str(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_header(text):
    if text is None:
        return ""
    return str(text).strip().lower().replace(".", "").replace(" ", "").replace("-", "")


def is_allowed(message):
    return message.chat.id == ALLOWED_CHAT_ID


def format_excel_date(value):
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT00:00:00")
    text = str(value).strip()
    formats = ["%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"]
    for fmt in formats:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%Y-%m-%dT00:00:00")
        except Exception:
            pass
    return None


def classify_intent(user_message):
    text = (user_message or "").strip().lower()
    fallback_map = [
        (["hi", "hello", "start", "help", "menu"], "start_help"),
        (["new order", "new sales order", "manual order", "enter order", "create order manually", "single order"], "new_order"),
        (["create sales order", "create sales orders", "process", "run orders", "run sales orders"], "process_pending"),
        (["retry", "failed rows", "retry errors", "retry failed"], "retry_errors"),
        (["send excel", "send file", "download excel", "send workbook"], "send_excel"),
        (["status", "bot status", "configuration", "setup"], "bot_status"),
        (["stop", "exit", "close bot", "shut down"], "stop_bot"),
    ]

    try:
        prompt = INTENT_PROMPT.format(message=user_message)
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[
                {"role": "system", "content": "You are an intent classifier. Reply with only one label."},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )
        intent = (response.choices[0].message.content or "").strip().lower()
        valid_intents = {
            "start_help", "process_pending", "retry_errors",
            "send_excel", "bot_status", "stop_bot", "new_order", "unknown"
        }
        if intent in valid_intents:
            return intent
        for item in valid_intents:
            if item in intent:
                return item
    except Exception as e:
        logging.error(f"Groq intent error: {str(e)}")
        print("Groq intent error:", str(e))

    for keywords, intent in fallback_map:
        if any(k in text for k in keywords):
            return intent
    return "unknown"


def get_csrf_token():
    response = requests.get(
        f"{SAP_BASE_URL}/$metadata",
        auth=AUTH,
        headers={"X-CSRF-Token": "Fetch", "Accept": "application/xml"},
        verify=False,
        timeout=REQUEST_TIMEOUT
    )
    response.raise_for_status()
    token = response.headers.get("X-CSRF-Token")
    if not token:
        raise Exception("Unable to fetch CSRF token from SAP.")
    return token, response.cookies


def extract_sap_error(response):
    try:
        data = response.json()
        if "error" in data:
            err = data["error"]
            if isinstance(err, dict):
                msg = err.get("message")
                if isinstance(msg, dict):
                    return safe_str(msg.get("value"))
                if msg:
                    return safe_str(msg)
                inner = err.get("innererror", {})
                if isinstance(inner, dict):
                    details = inner.get("errordetails", [])
                    if details and isinstance(details, list):
                        first = details[0]
                        if isinstance(first, dict):
                            return safe_str(first.get("message"))
        return safe_str(response.text)[:300]
    except Exception:
        return safe_str(response.text)[:300]


def create_sales_order(payload):
    token, cookies = get_csrf_token()
    response = requests.post(
        f"{SAP_BASE_URL}/A_SalesOrder",
        auth=AUTH,
        headers={
            "X-CSRF-Token": token,
            "Content-Type": "application/json",
            "Accept": "application/json"
        },
        cookies=cookies,
        json=payload,
        verify=False,
        timeout=REQUEST_TIMEOUT
    )
    if response.status_code in [200, 201]:
        try:
            data = response.json()
            if "d" in data and isinstance(data["d"], dict):
                so_number = safe_str(data["d"].get("SalesOrder"))
            else:
                so_number = safe_str(data.get("SalesOrder"))
            if so_number:
                return True, so_number, "Success"
            return False, "", "Sales order created but number not returned."
        except Exception:
            return False, "", "Success response received but JSON parsing failed."
    return False, "", extract_sap_error(response)


def get_workbook_sheet_headers():
    if not os.path.exists(EXCEL_PATH):
        raise Exception("Excel file not found.")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    if SHEET_NAME not in wb.sheetnames:
        raise Exception(f"Sheet '{SHEET_NAME}' not found.")
    ws = wb[SHEET_NAME]
    headers = {}
    for col in range(1, ws.max_column + 1):
        headers[normalize_header(ws.cell(row=1, column=col).value)] = col
    required = {
        "ordertype": "Order Type",
        "salesorganization": "Sales Organization",
        "distributionchannel": "Distribution Channel",
        "division": "Division",
        "soldtoparty": "Sold-to Party",
        "shiptoparty": "Ship-to Party",
        "custreference": "Cust.Reference",
        "custrefdate": "Cust. Ref. Date",
        "material": "Material",
        "orderquantity": "Order Quantity",
        "status": "Status",
        "salesordernumber": "Sales Order Number"
    }
    missing = [v for k, v in required.items() if k not in headers]
    if missing:
        raise Exception("Missing Excel columns: " + ", ".join(missing))
    return wb, ws, headers


def build_payload_from_row(ws, row, h):
    order_type = safe_str(ws.cell(row=row, column=h["ordertype"]).value)
    sales_org = safe_str(ws.cell(row=row, column=h["salesorganization"]).value)
    dist_channel = safe_str(ws.cell(row=row, column=h["distributionchannel"]).value)
    division = safe_str(ws.cell(row=row, column=h["division"]).value)
    sold_to = safe_str(ws.cell(row=row, column=h["soldtoparty"]).value)
    ship_to = safe_str(ws.cell(row=row, column=h["shiptoparty"]).value)
    cust_ref = safe_str(ws.cell(row=row, column=h["custreference"]).value)
    cust_ref_date = ws.cell(row=row, column=h["custrefdate"]).value
    material = safe_str(ws.cell(row=row, column=h["material"]).value)
    quantity = safe_str(ws.cell(row=row, column=h["orderquantity"]).value)

    missing = []
    if not order_type:
        missing.append("Order Type")
    if not sales_org:
        missing.append("Sales Organization")
    if not dist_channel:
        missing.append("Distribution Channel")
    if not division:
        missing.append("Division")
    if not sold_to:
        missing.append("Sold-to Party")
    if not material:
        missing.append("Material")
    if not quantity:
        missing.append("Order Quantity")
    if missing:
        raise Exception("Mandatory fields missing: " + ", ".join(missing))

    payload = {
        "SalesOrderType": order_type,
        "SalesOrganization": sales_org,
        "DistributionChannel": dist_channel,
        "OrganizationDivision": division,
        "SoldToParty": sold_to,
        "PurchaseOrderByCustomer": cust_ref,
        "to_Item": {
            "results": [
                {
                    "Material": material,
                    "RequestedQuantity": quantity,
                    "RequestedQuantityUnit": DEFAULT_UOM
                }
            ]
        }
    }

    sap_date = format_excel_date(cust_ref_date)
    if sap_date:
        payload["CustomerPurchaseOrderDate"] = sap_date

    if ship_to:
        payload["to_Partner"] = {
            "results": [
                {
                    "PartnerFunction": "WE",
                    "Customer": ship_to
                }
            ]
        }
    return payload


def process_excel_orders(mode="pending"):
    wb, ws, h = get_workbook_sheet_headers()
    status_col = h["status"]
    so_col = h["salesordernumber"]
    success_count = 0
    error_count = 0
    skipped_count = 0
    created_orders = []

    for row in range(2, ws.max_row + 1):
        status_value = safe_str(ws.cell(row=row, column=status_col).value)
        so_value = safe_str(ws.cell(row=row, column=so_col).value)
        status_lower = status_value.lower()

        if mode == "pending":
            if status_lower == "success" and so_value:
                skipped_count += 1
                continue
            if status_value != "":
                skipped_count += 1
                continue

        if mode == "retryerrors":
            if not status_lower.startswith("error"):
                skipped_count += 1
                continue

        try:
            payload = build_payload_from_row(ws, row, h)
            ok, so_number, message = create_sales_order(payload)
            if ok:
                ws.cell(row=row, column=status_col).value = "Success"
                ws.cell(row=row, column=so_col).value = so_number
                success_count += 1
                created_orders.append(so_number)
                log_info(f"Row {row}: Success - SO {so_number}")
            else:
                ws.cell(row=row, column=status_col).value = f"Error - {message[:200]}"
                error_count += 1
                log_error(f"Row {row}: Error - {message[:200]}")
        except Exception as e:
            ws.cell(row=row, column=status_col).value = f"Error - {str(e)[:200]}"
            error_count += 1
            log_error(f"Row {row}: Exception - {str(e)[:200]}")

    wb.save(EXCEL_PATH)
    return success_count, error_count, skipped_count, created_orders


def create_empty_item():
    return {"material": "", "quantity": ""}


def get_current_item(session):
    index = session.get("current_item_index", 0)
    items = session.setdefault("items", [])
    while len(items) <= index:
        items.append(create_empty_item())
    return items[index]


def start_manual_order_session(chat_id):
    user_sessions[chat_id] = {
        "mode": "header",
        "step": 0,
        "data": {},
        "items": [create_empty_item()],
        "current_item_index": 0,
        "item_step": 0,
    }
    _, question = MANUAL_ORDER_FIELDS[0]
    bot.send_message(
        chat_id,
        "🆕 *New Sales Order — Manual Entry*\n\n"
        "You can type CANCEL at any time to abort.\n\n" + question,
        parse_mode="Markdown"
    )


def handle_manual_order_step(chat_id, user_input):
    session = user_sessions.get(chat_id)
    if not session:
        return

    text = user_input.strip()
    if text.upper() == "CANCEL":
        del user_sessions[chat_id]
        bot.send_message(chat_id, "❌ Order entry cancelled.")
        return

    mode = session.get("mode", "header")
    if mode == "header":
        handle_header_step(chat_id, text, session)
    elif mode == "item":
        handle_item_step(chat_id, text, session)
    elif mode == "more_items":
        handle_more_items_step(chat_id, text, session)


def handle_header_step(chat_id, text, session):
    step = session["step"]
    field_key, _ = MANUAL_ORDER_FIELDS[step]
    optional_fields = {"ship_to", "cust_ref", "cust_ref_date"}

    if text.upper() == "SKIP" and field_key in optional_fields:
        session["data"][field_key] = ""
    else:
        if field_key == "cust_ref_date":
            parsed = format_excel_date(text)
            if not parsed:
                bot.send_message(chat_id, "⚠️ Invalid date format. Please use DD.MM.YYYY (e.g. 15.06.2026) or type SKIP.")
                return
        session["data"][field_key] = text

    session["step"] += 1
    next_step = session["step"]
    if next_step < len(MANUAL_ORDER_FIELDS):
        _, next_question = MANUAL_ORDER_FIELDS[next_step]
        bot.send_message(chat_id, next_question)
    else:
        session["mode"] = "item"
        session["item_step"] = 0
        ask_current_item_question(chat_id, session)


def ask_current_item_question(chat_id, session):
    item_no = session["current_item_index"] + 1
    field_key, template = LINE_ITEM_FIELDS[session["item_step"]]
    bot.send_message(chat_id, template.format(item_no=item_no, field_key=field_key))


def handle_item_step(chat_id, text, session):
    field_key, _ = LINE_ITEM_FIELDS[session["item_step"]]
    current_item = get_current_item(session)

    if field_key == "quantity":
        try:
            float(text)
        except ValueError:
            bot.send_message(chat_id, "⚠️ Please enter a valid number for Order Quantity.")
            return

    if field_key == "material" and not text:
        bot.send_message(chat_id, "⚠️ Material cannot be blank.")
        return

    current_item[field_key] = text
    session["item_step"] += 1

    if session["item_step"] < len(LINE_ITEM_FIELDS):
        ask_current_item_question(chat_id, session)
        return

    item_no = session["current_item_index"] + 1
    bot.send_message(
        chat_id,
        f"✅ Line item {item_no} captured.\n\nDo you have more line item? Reply with *YES* or *NO*.",
        parse_mode="Markdown"
    )
    session["mode"] = "more_items"


def handle_more_items_step(chat_id, text, session):
    answer = text.strip().upper()
    if answer in ["YES", "Y"]:
        session["current_item_index"] += 1
        session["items"].append(create_empty_item())
        session["item_step"] = 0
        session["mode"] = "item"
        ask_current_item_question(chat_id, session)
        return

    if answer in ["NO", "N"]:
        session["mode"] = "confirm"
        show_order_confirmation(chat_id, session["data"], session["items"])
        return

    bot.send_message(chat_id, "Please reply with YES to add another line item or NO to continue.")


def show_order_confirmation(chat_id, data, items):
    ship_to_display = data.get("ship_to") or "(same as Sold-to)"
    cust_ref_display = data.get("cust_ref") or "(none)"
    cust_ref_date_display = data.get("cust_ref_date") or "(none)"

    item_lines = []
    for idx, item in enumerate(items, start=1):
        item_lines.append(
            f"{idx}. Material: `{item.get('material')}` | Qty: `{item.get('quantity')} {DEFAULT_UOM}`"
        )
    items_text = "\n".join(item_lines)

    summary = (
        "✅ *Please confirm the Sales Order details:*\n\n"
        f"📋 Order Type : `{data.get('order_type')}`\n"
        f"🏢 Sales Organization : `{data.get('sales_org')}`\n"
        f"📦 Distribution Channel: `{data.get('dist_channel')}`\n"
        f"🗂️ Division : `{data.get('division')}`\n"
        f"👤 Sold-to Party : `{data.get('sold_to')}`\n"
        f"🚚 Ship-to Party : `{ship_to_display}`\n"
        f"📝 Customer Reference : `{cust_ref_display}`\n"
        f"📅 Reference Date : `{cust_ref_date_display}`\n\n"
        f"📦 *Line Items:*\n{items_text}\n\n"
        "Type *CONFIRM* to create the order or *CANCEL* to abort."
    )
    user_sessions[chat_id]["step"] = "confirm"
    bot.send_message(chat_id, summary, parse_mode="Markdown")


def confirm_and_create_manual_order(chat_id, user_input):
    text = user_input.strip().upper()
    session = user_sessions.get(chat_id)

    if text == "CANCEL":
        del user_sessions[chat_id]
        bot.send_message(chat_id, "❌ Order creation cancelled.")
        return

    if text != "CONFIRM":
        bot.send_message(chat_id, "Please type CONFIRM to create or CANCEL to abort.")
        return

    data = session["data"]
    items = session["items"]
    bot.send_message(chat_id, "⏳ Creating sales order in SAP, please wait...")

    try:
        payload = build_manual_payload(data, items)
        ok, so_number, message = create_sales_order(payload)
        if ok:
            del user_sessions[chat_id]
            bot.send_message(
                chat_id,
                f"✅ *Sales Order Created Successfully!*\n\n📄 Sales Order Number: `{so_number}`",
                parse_mode="Markdown"
            )
            log_info(f"Manual order created: SO {so_number}")
        else:
            del user_sessions[chat_id]
            bot.send_message(
                chat_id,
                f"❌ *Failed to create Sales Order*\n\nSAP Error: {message}",
                parse_mode="Markdown"
            )
            log_error(f"Manual order failed: {message}")
    except Exception as e:
        del user_sessions[chat_id]
        bot.send_message(chat_id, f"❌ Error: {str(e)}")
        log_error(f"Manual order exception: {str(e)}")


def build_manual_payload(data, items):
    order_type = data.get("order_type", "")
    sales_org = data.get("sales_org", "")
    dist_channel = data.get("dist_channel", "")
    division = data.get("division", "")
    sold_to = data.get("sold_to", "")
    ship_to = data.get("ship_to", "")
    cust_ref = data.get("cust_ref", "")
    cust_ref_date = data.get("cust_ref_date", "")

    missing = []
    if not order_type:
        missing.append("Order Type")
    if not sales_org:
        missing.append("Sales Organization")
    if not dist_channel:
        missing.append("Distribution Channel")
    if not division:
        missing.append("Division")
    if not sold_to:
        missing.append("Sold-to Party")
    if not items:
        missing.append("Line Items")
    if missing:
        raise Exception("Mandatory fields missing: " + ", ".join(missing))

    item_results = []
    for item in items:
        material = safe_str(item.get("material"))
        quantity = safe_str(item.get("quantity"))
        if not material or not quantity:
            raise Exception("Each line item must have Material and Order Quantity.")
        item_results.append(
            {
                "Material": material,
                "RequestedQuantity": quantity,
                "RequestedQuantityUnit": DEFAULT_UOM
            }
        )

    payload = {
        "SalesOrderType": order_type,
        "SalesOrganization": sales_org,
        "DistributionChannel": dist_channel,
        "OrganizationDivision": division,
        "SoldToParty": sold_to,
        "PurchaseOrderByCustomer": cust_ref,
        "to_Item": {"results": item_results}
    }

    sap_date = format_excel_date(cust_ref_date)
    if sap_date:
        payload["CustomerPurchaseOrderDate"] = sap_date

    if ship_to:
        payload["to_Partner"] = {
            "results": [
                {
                    "PartnerFunction": "WE",
                    "Customer": ship_to
                }
            ]
        }
    return payload


def send_help(chat_id):
    bot.send_message(
        chat_id,
        "👋 *Hello! I am your SAP Sales Order Bot.*\n\n"
        "*📥 Manual Order Entry (Multi Line Items):*\n"
        "- 'New sales order'\n"
        "- 'Create order manually'\n"
        "- /neworder\n\n"
        "*📊 Excel Batch Processing:*\n"
        "- 'Create sales orders' (from Excel)\n"
        "- 'Retry failed rows'\n"
        "- /process\n"
        "- /retryerrors\n\n"
        "*📁 Other Commands:*\n"
        "- 'Send me the Excel file' / /sendexcel\n"
        "- 'Show bot status' / /status\n"
        "- 'Stop the bot' / /stop",
        parse_mode="Markdown"
    )


def handle_intent(chat_id, intent, message):
    if intent == "start_help":
        send_help(chat_id)
    elif intent == "new_order":
        start_manual_order_session(chat_id)
    elif intent == "process_pending":
        bot.send_message(chat_id, "⏳ Processing pending sales orders from Excel...")
        try:
            success, errors, skipped, orders = process_excel_orders(mode="pending")
            reply = f"✅ Process completed.\nSuccess: {success}, Errors: {errors}, Skipped: {skipped}"
            if orders:
                reply += "\n\n📄 Created Sales Order Numbers:\n" + "\n".join(orders)
            bot.send_message(chat_id, reply)
        except Exception as e:
            bot.send_message(chat_id, f"❌ Error: {str(e)}")
    elif intent == "retry_errors":
        bot.send_message(chat_id, "🔄 Retrying rows with Error status...")
        try:
            success, errors, skipped, orders = process_excel_orders(mode="retryerrors")
            reply = f"✅ Retry completed.\nSuccess: {success}, Errors: {errors}, Skipped: {skipped}"
            if orders:
                reply += "\n\n📄 Created Sales Order Numbers:\n" + "\n".join(orders)
            bot.send_message(chat_id, reply)
        except Exception as e:
            bot.send_message(chat_id, f"❌ Error: {str(e)}")
    elif intent == "send_excel":
        if os.path.exists(EXCEL_PATH):
            with open(EXCEL_PATH, "rb") as f:
                bot.send_document(chat_id, f)
        else:
            bot.send_message(chat_id, "❌ Excel file not found.")
    elif intent == "bot_status":
        text = (
            "🤖 *Bot Status:*\n"
            f"Excel Path : `{EXCEL_PATH}`\n"
            f"Sheet Name : `{SHEET_NAME}`\n"
            f"SAP URL : `{SAP_BASE_URL}`\n"
            f"Default UOM: `{DEFAULT_UOM}`\n"
            f"Log File : `{LOG_FILE}`\n"
            f"Groq Model : `{GROQ_MODEL}`"
        )
        bot.send_message(chat_id, text, parse_mode="Markdown")
    elif intent == "stop_bot":
    	if chat_id in user_sessions:
        	del user_sessions[chat_id]
    	bot.send_message(chat_id, "🛑 Session ended. Say 'Hi' anytime to start again.")
    	log_info("User session stopped by command (bot still running).")

    else:
        bot.send_message(
            chat_id,
            "🤔 I did not understand that.\n\n"
            "Try:\n"
            "- 'New sales order'\n"
            "- 'Create sales orders' (Excel)\n"
            "- 'Retry failed rows'\n"
            "- 'Send me the Excel'\n"
            "- 'Show status'\n"
            "- 'Stop bot'"
        )


@bot.message_handler(commands=["start"])
def start_command(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return
    send_help(message.chat.id)


@bot.message_handler(commands=["neworder"])
def new_order_command(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return
    start_manual_order_session(message.chat.id)


@bot.message_handler(commands=["process"])
def process_command(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return
    handle_intent(message.chat.id, "process_pending", message.text)


@bot.message_handler(commands=["retryerrors"])
def retry_errors_command(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return
    handle_intent(message.chat.id, "retry_errors", message.text)


@bot.message_handler(commands=["sendexcel"])
def send_excel_command(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return
    handle_intent(message.chat.id, "send_excel", message.text)


@bot.message_handler(commands=["status"])
def status_command(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return
    handle_intent(message.chat.id, "bot_status", message.text)


@bot.message_handler(commands=["stop"])
def stop_command(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return
    handle_intent(message.chat.id, "stop_bot", message.text)


@bot.message_handler(func=lambda m: True)
def all_text_handler(message):
    if not is_allowed(message):
        bot.reply_to(message, "Unauthorized user.")
        return

    chat_id = message.chat.id
    user_text = safe_str(message.text)
    log_info(f"Received: {user_text}")

    if chat_id in user_sessions:
        session = user_sessions[chat_id]
        if session.get("mode") == "confirm" or session.get("step") == "confirm":
            confirm_and_create_manual_order(chat_id, user_text)
        else:
            handle_manual_order_step(chat_id, user_text)
        return

    bot.send_message(chat_id, "🔍 Let me understand your request...")
    intent = classify_intent(user_text)
    log_info(f"Groq intent: {intent}")
    handle_intent(chat_id, intent, user_text)


log_info("Telegram SAP OData Bot V6 (Groq + Multi Line Manual Order) is running...")
while True:
    try:
        bot.infinity_polling(skip_pending=True, timeout=60, long_polling_timeout=30)
    except Exception as e:
        logging.error(f"Bot crashed, restarting in 10s: {e}")
        time.sleep(10)
