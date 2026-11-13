import os
import sys
import io
import gc
import re
import asyncio
import threading
import time
from datetime import datetime
from urllib.parse import quote
import pytz
import requests
import msal
import openpyxl
from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes

# ================== CẤU HÌNH (100% TỪ ENV) ==================
TOKEN = os.getenv("BOT_TOKEN")
BOT_PASSWORD = os.getenv("BOT_PASSWORD", "123123")
EXPENSE_PASSWORD = os.getenv("EXPENSE_PASSWORD", "0939")

# --- Microsoft Graph / OneDrive ---
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")
GRAPH_AUTHORITY = os.getenv("GRAPH_AUTHORITY", "https://login.microsoftonline.com/consumers")
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read", "offline_access"]
GRAPH_REFRESH_TOKEN = os.getenv("GRAPH_REFRESH_TOKEN", "")
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "")

# --- SePay / Telegram ---
SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")

# --- Excel ---
EXCEL_DATA_START_ROW = int(os.getenv("EXCEL_DATA_START_ROW", "2"))
EXCEL_DATA_END_ROW = int(os.getenv("EXCEL_DATA_END_ROW", "27"))

# --- API test ---
API_TEST_TOKEN = os.getenv("API_TEST_TOKEN", "")

app_web = Flask(__name__)
logged_in_users = {}
awaiting_expense_password = set()
expense_mode_users = set()
excel_write_lock = threading.Lock()

AMOUNT_PATTERN = re.compile(r'^([+-]?)\s*([\d.,]+)\s*([kK]?)$')

BANKS = [
    {"name": "Vietcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "96886693059121", "qr_url": "https://api.vietqr.io/image/MSB-96886693059121-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "OCB (Phương Đông)", "account_name": "HUYNH NGOC NGHIA", "account_number": "OCB-SEPHN48935", "qr_url": "https://api.vietqr.io/image/OCB-SEPHN48935-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "LPBank", "account_name": "HUYNH NGOC NGHIA", "account_number": "0916856322", "qr_url": "https://api.vietqr.io/image/970449-0916856322-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "Techcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "3838396852", "qr_url": "https://api.vietqr.io/image/970407-3838396852-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"}
]

# ================== CLEAR CHAT ==================
async def clear_chat_history(context, chat_id, from_message_id, max_delete=2000, max_consecutive_fail=40):
    msg_id = from_message_id
    consecutive_fail = 0
    deleted = 0
    while msg_id > 0 and deleted < max_delete and consecutive_fail < max_consecutive_fail:
        try:
            ok = await context.bot.delete_message(chat_id=chat_id, message_id=msg_id)
            if ok:
                deleted += 1
                consecutive_fail = 0
            else:
                consecutive_fail += 1
        except Exception:
            consecutive_fail += 1
        msg_id -= 1
    return deleted

# ================== MICROSOFT GRAPH AUTH ==================
def _validate_graph_config():
    if not GRAPH_CLIENT_ID:
        raise Exception("GRAPH_CLIENT_ID đang trống! Cần set env trên Render.")
    if not GRAPH_REFRESH_TOKEN:
        raise Exception("GRAPH_REFRESH_TOKEN đang trống! Cần set env trên Render.")
    if not ONEDRIVE_FILE_PATH:
        raise Exception("ONEDRIVE_FILE_PATH đang trống! Cần set env trên Render.")

def get_graph_access_token():
    _validate_graph_config()

    # Cách 1: token cache serialized
    try:
        cache = msal.SerializableTokenCache()
        cache.deserialize(GRAPH_REFRESH_TOKEN)
        app = msal.PublicClientApplication(GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY, token_cache=cache)
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]
    except Exception as e:
        print(f"⚠️ acquire_token_silent thất bại: {e}")

    # Cách 2: refresh_token thô
    try:
        app = msal.PublicClientApplication(GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY)
        result = app.acquire_token_by_refresh_token(GRAPH_REFRESH_TOKEN, scopes=GRAPH_SCOPES)
        if result and "access_token" in result:
            new_refresh = result.get("refresh_token")
            if new_refresh:
                print("🔑 Refresh token mới (copy vào env GRAPH_REFRESH_TOKEN nếu muốn):", new_refresh[:20] + "...")
            return result["access_token"]
    except Exception as e:
        raise Exception(f"Không lấy được access token: {e}")

    raise Exception(
        "Không lấy được access token. Kiểm tra GRAPH_CLIENT_ID / GRAPH_REFRESH_TOKEN. "
        "(refresh_token có thể đã hết hạn — cần login lại bằng script local.)"
    )

def _encoded_file_path():
    path = ONEDRIVE_FILE_PATH.strip().strip("/")
    if not path:
        raise Exception("ONEDRIVE_FILE_PATH đang trống!")
    if "/" in path:
        folder, filename = path.rsplit("/", 1)
        return f"{quote(folder)}/{quote(filename)}"
    return quote(path)

# ================== ĐỌC / GHI EXCEL ==================
def download_excel_for_write():
    token = get_graph_access_token()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{_encoded_file_path()}:/content"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if resp.status_code != 200:
        raise Exception(f"Graph API tải file lỗi {resp.status_code}: {resp.text[:300]}")
    return resp.content

def get_excel_data():
    try:
        with excel_write_lock:
            content = download_excel_for_write()

        workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        sheet = workbook.active

        thu_list, chi_list = [], []
        for i in range(EXCEL_DATA_START_ROW, EXCEL_DATA_END_ROW + 1):
            v_thu = sheet.cell(row=i, column=1).value
            v_chi = sheet.cell(row=i, column=2).value
            if v_thu is not None:
                thu_list.append(v_thu)
            if v_chi is not None:
                chi_list.append(v_chi)

        thu_total = sum(v for v in thu_list if isinstance(v, (int, float)))
        chi_total = sum(v for v in chi_list if isinstance(v, (int, float)))
        remain = thu_total - chi_total

        msg = "📊 <b>BÁO CÁO SỔ THU CHI (ONEDRIVE)</b>\n" + "-" * 40 + "\n"
        msg += "📥 <b>DANH SÁCH THU:</b>\n"
        for val in thu_list:
            msg += f"  • {val:,.0f} VNĐ\n" if isinstance(val, (int, float)) else f"  • {val}\n"
        msg += "\n📤 <b>DANH SÁCH CHI:</b>\n"
        for val in chi_list:
            msg += f"  • {val:,.0f} VNĐ\n" if isinstance(val, (int, float)) else f"  • {val}\n"
        msg += "-" * 40 + "\n"
        msg += f"🟢 <b>Tổng Tiền Nhận Vào:</b> <code>{thu_total:,.0f}</code> VNĐ\n"
        msg += f"🔴 <b>Tổng Tiền Đã Chi:</b> <code>{chi_total:,.0f}</code> VNĐ\n"
        msg += f"💰 <b>SỐ TIỀN CÒN LẠI:</b> <code>{remain:,.0f}</code> VNĐ\n"

        del workbook, sheet
        gc.collect()
        return msg
    except Exception as e:
        return f"❌ Lỗi đọc Excel: {e}"

def upload_excel(content_bytes, max_retries=4):
    token = get_graph_access_token()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{_encoded_file_path()}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    last_error = "Upload thất bại"
    for attempt in range(1, max_retries + 1):
        resp = requests.put(url, headers=headers, data=content_bytes, timeout=30)
        if resp.status_code in (200, 201):
            return resp.json()
        if resp.status_code == 423:
            last_error = "File đang bị khóa trên OneDrive (423 Locked)."
            print(f"⚠️ 423 Locked, thử lại {attempt}/{max_retries}...")
            time.sleep(2 * attempt)
            continue
        print(f"❌ Graph upload lỗi {resp.status_code}: {resp.text[:300]}")
        resp.raise_for_status()

    raise Exception(
        f"{last_error} Đã thử {max_retries} lần. Đóng hẳn file Excel trên mọi thiết bị rồi gửi lại."
    )

def append_transaction_and_upload(amount, is_income):
    with excel_write_lock:
        content = download_excel_for_write()
        workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
        sheet = workbook.active

        col = 1 if is_income else 2
        target_row = None
        for i in range(EXCEL_DATA_START_ROW, EXCEL_DATA_END_ROW + 1):
            if sheet.cell(row=i, column=col).value is None:
                target_row = i
                break
        if target_row is None:
            raise Exception(f"Hết chỗ trống trong bảng (dòng {EXCEL_DATA_START_ROW}-{EXCEL_DATA_END_ROW}).")

        sheet.cell(row=target_row, column=col).value = amount

        buf = io.BytesIO()
        workbook.save(buf)
        buf.seek(0)
        upload_excel(buf.read())

        del workbook, sheet, content, buf
        gc.collect()

# ================== PARSE AMOUNT ==================
def parse_amount_message(text):
    if not text:
        return None
    match = AMOUNT_PATTERN.match(text.strip())
    if not match:
        return None
    _sign, number_part, k_suffix = match.groups()

    if not re.fullmatch(r'\d{1,3}(\.\d{3})*|\d+', number_part):
        return None
    cleaned = number_part.replace('.', '')
    if not cleaned.isdigit():
        return None

    amount = int(cleaned)
    if k_suffix:
        amount *= 1000
    if amount <= 0:
        return None
    return amount

# ================== API TEST ==================
def _check_api_auth():
    if not API_TEST_TOKEN:
        return True
    return request.headers.get("X-Api-Token") == API_TEST_TOKEN

@app_web.route("/ping", methods=["GET"])
def ping():
    return "OK", 200

@app_web.route("/api/graph/token", methods=["GET"])
def api_graph_token():
    if not _check_api_auth():
        return jsonify({"success": False, "message": "Unauthorized"}), 401
    try:
        token = get_graph_access_token()
        return jsonify({"success": True, "token_prefix": token[:20] + "..."})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app_web.route("/api/excel/read", methods=["GET"])
def api_excel_read():
    if not _check_api_auth():
        return jsonify({"success": False, "message": "Unauthorized"}), 401
    try:
        with excel_write_lock:
            content = download_excel_for_write()
        workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        sheet = workbook.active

        rows = []
        for i in range(EXCEL_DATA_START_ROW, EXCEL_DATA_END_ROW + 1):
            thu = sheet.cell(row=i, column=1).value
            chi = sheet.cell(row=i, column=2).value
            if thu is not None or chi is not None:
                rows.append({"row": i, "thu": thu, "chi": chi})

        thu_total = sum(r["thu"] for r in rows if isinstance(r["thu"], (int, float)))
        chi_total = sum(r["chi"] for r in rows if isinstance(r["chi"], (int, float)))

        return jsonify({
            "success": True,
            "file_path": ONEDRIVE_FILE_PATH,
            "rows": rows,
            "thu_total": thu_total,
            "chi_total": chi_total,
            "remain": thu_total - chi_total,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app_web.route("/api/excel/append", methods=["POST"])
def api_excel_append():
    if not _check_api_auth():
        return jsonify({"success": False, "message": "Unauthorized"}), 401
    data = request.get_json(force=True, silent=True) or {}
    try:
        amount = int(data.get("amount", 0))
        if amount <= 0:
            return jsonify({"success": False, "error": "amount phải > 0"}), 400
        is_income = data.get("type", "expense") == "income"
        append_transaction_and_upload(amount, is_income)
        return jsonify({
            "success": True,
            "amount": amount,
            "type": "income" if is_income else "expense",
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

# ================== SEPAY WEBHOOK ==================
@app_web.route("/sepay-webhook", methods=["POST"])
def sepay_webhook():
    auth_header = request.headers.get("Authorization", "")
    if not SEPAY_API_KEY or SEPAY_API_KEY not in auth_header:
        return jsonify({"success": False, "message": "Unauthorized"}), 401
    data = request.get_json(force=True, silent=True) or {}
    threading.Thread(target=process_transaction, args=(data,), daemon=True).start()
    return jsonify({"success": True}), 200

def process_transaction(data):
    try:
        noi_dung = data.get("content", "")
        so_tien = data.get("transferAmount", 0)
        loai_gd = data.get("transferType")
        is_income = (loai_gd == "in")

        loi_chi_tiet = ""
        try:
            append_transaction_and_upload(so_tien, is_income)
            ghi_thanh_cong = True
        except Exception as e:
            ghi_thanh_cong = False
            loi_chi_tiet = str(e)
            print(f"❌ Lỗi ghi Excel: {loi_chi_tiet}")

        loai_text = "💰 Nhận tiền (in)" if is_income else "💸 Chi tiền (out)"

        if ghi_thanh_cong:
            send_telegram_notification(
                f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
                f"Nội dung: {noi_dung}\n"
                f"----------------------------------------\n"
                f"✅ Đã ghi vào Excel thành công!"
            )
        else:
            is_lock_error = ("423" in loi_chi_tiet) or ("khóa" in loi_chi_tiet.lower())
            msg_id = send_telegram_notification(
                f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
                f"Nội dung: {noi_dung}\n"
                + (
                    f"⏳ File đang bị khóa, bot sẽ tự thử lại nền trong vài phút..."
                    if is_lock_error else
                    f"❌ Lỗi ghi file:\n<code>{loi_chi_tiet}</code>"
                )
            )
            if is_lock_error and msg_id and TELEGRAM_CHAT_ID:
                background_retry_write(so_tien, is_income, TELEGRAM_CHAT_ID, msg_id)
    except Exception as e:
        print(f"❌ LỖI XỬ LÝ GIAO DỊCH SEPAY: {str(e)}")

def send_telegram_notification(text):
    if not TOKEN or not TELEGRAM_CHAT_ID:
        return None
    try:
        resp = requests.post(
            f"https://api.telegram.org/bot{TOKEN}/sendMessage",
            json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"},
            timeout=15,
        )
        return resp.json().get("result", {}).get("message_id")
    except Exception:
        print("Không gửi được tin nhắn Telegram.")
        return None

def edit_telegram_message(chat_id, message_id, text):
    if not TOKEN:
        return
    try:
        requests.post(
            f"https://api.telegram.org/bot{TOKEN}/editMessageText",
            json={"chat_id": chat_id, "message_id": message_id, "text": text, "parse_mode": "HTML"},
            timeout=15,
        )
    except Exception as e:
        print(f"Không sửa được tin nhắn Telegram: {e}")

def background_retry_write(amount, is_income, chat_id, message_id, attempt=1, max_attempts=5, delay_seconds=90):
    loai_text = "💰 Thu" if is_income else "💸 Chi"

    def _attempt():
        try:
            append_transaction_and_upload(amount, is_income)
            edit_telegram_message(
                chat_id, message_id,
                f"{loai_text}: <code>{amount:,.0f}</code> VNĐ\n✅ Đã ghi vào Excel thành công! (tự thử lại nền)"
            )
        except Exception as e:
            if attempt < max_attempts:
                edit_telegram_message(
                    chat_id, message_id,
                    f"{loai_text}: <code>{amount:,.0f}</code> VNĐ\n"
                    f"⏳ File vẫn đang bị khóa, bot đang tự thử lại nền (lần {attempt}/{max_attempts})..."
                )
                threading.Timer(
                    delay_seconds, background_retry_write,
                    args=(amount, is_income, chat_id, message_id, attempt + 1, max_attempts, delay_seconds)
                ).start()
            else:
                edit_telegram_message(
                    chat_id, message_id,
                    f"{loai_text}: <code>{amount:,.0f}</code> VNĐ\n"
                    f"❌ Vẫn lỗi sau nhiều lần thử nền:\n<code>{str(e)}</code>\n"
                    f"Vui lòng kiểm tra file Excel trên OneDrive rồi gửi lại số tiền."
                )

    threading.Thread(target=_attempt, daemon=True).start()

# ================== TELEGRAM BOT ==================
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📱 Lấy mã QR", callback_data="get_qr")],
        [InlineKeyboardButton("💰 Kiểm tra tiền", callback_data="check_money")],
        [InlineKeyboardButton("📝 Ghi sổ chi", callback_data="add_expense")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        "🤖 Xin chào!\nNhập số <b>1</b> hoặc bấm menu dưới đây để chọn chức năng.\nNhập số <b>2</b> để đăng xuất.\n\n"
        "📝 Muốn ghi khoản <b>chi</b> vào Excel, bấm nút <b>Ghi sổ chi</b> — cần nhập thêm mật khẩu riêng.",
        reply_markup=reply_markup, parse_mode="HTML"
    )

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if logged_in_users.get(user_id):
        await show_menu(update, context)
    else:
        await update.message.reply_text("🔐 <