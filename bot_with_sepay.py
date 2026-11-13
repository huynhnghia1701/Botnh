import os
import sys
import io
import gc
import re
import asyncio
import threading
import time
import multiprocessing
import html
from datetime import datetime
from urllib.parse import quote
import pytz
import requests
import msal
import openpyxl
from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes

# ================== CẤU HÌNH ==================
TOKEN = os.getenv("BOT_TOKEN")
BOT_PASSWORD = os.getenv("BOT_PASSWORD", "123123")
EXPENSE_PASSWORD = os.getenv("EXPENSE_PASSWORD", "0939")
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")
GRAPH_AUTHORITY = os.getenv("GRAPH_AUTHORITY", "https://login.microsoftonline.com/consumers")
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]
GRAPH_REFRESH_TOKEN = os.getenv("GRAPH_REFRESH_TOKEN", "")
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "")
SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")
EXCEL_DATA_START_ROW = int(os.getenv("EXCEL_DATA_START_ROW", "2"))
EXCEL_DATA_END_ROW = int(os.getenv("EXCEL_DATA_END_ROW", "27"))
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

# ================== GRAPH AUTH ==================
def _validate_graph_config():
    if not GRAPH_CLIENT_ID:
        raise Exception("GRAPH_CLIENT_ID đang trống!")
    if not GRAPH_REFRESH_TOKEN:
        raise Exception("GRAPH_REFRESH_TOKEN đang trống!")
    if not ONEDRIVE_FILE_PATH:
        raise Exception("ONEDRIVE_FILE_PATH đang trống!")

def get_graph_access_token():
    _validate_graph_config()
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
    try:
        app = msal.PublicClientApplication(GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY)
        result = app.acquire_token_by_refresh_token(GRAPH_REFRESH_TOKEN, scopes=GRAPH_SCOPES)
        if result and "access_token" in result:
            return result["access_token"]
    except Exception as e:
        raise Exception(f"Không lấy được access token: {e}")
    raise Exception("Không lấy được access token. Kiểm tra GRAPH_CLIENT_ID / GRAPH_REFRESH_TOKEN.")

def _encoded_file_path():
    path = ONEDRIVE_FILE_PATH.strip().strip("/")
    if not path:
        raise Exception("ONEDRIVE_FILE_PATH đang trống!")
    if "/" in path:
        folder, filename = path.rsplit("/", 1)
        return f"{quote(folder)}/{quote(filename)}"
    return quote(path)

# ================== ĐỌC EXCEL ==================
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

# ================== GHI EXCEL ==================
def upload_excel(content_bytes, max_retries=4):
    token = get_graph_access_token()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{_encoded_file_path()}:/content"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
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
    raise Exception(f"{last_error} Đã thử {max_retries} lần. Đóng hẳn file Excel trên mọi thiết bị rồi gửi lại.")

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
        return jsonify({"success": True, "file_path": ONEDRIVE_FILE_PATH, "rows": rows, "thu_total": thu_total, "chi_total": chi_total, "remain": thu_total - chi_total})
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
        return jsonify({"success": True, "amount": amount, "type": "income" if is_income else "expense"})
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
        is_income = (data.get("transferType") == "in")
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
            send_telegram_notification(f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\nNội dung: {noi_dung}\n----------------------------------------\n✅ Đã ghi vào Excel thành công!")
        else:
            is_lock_error = ("423" in loi_chi_tiet) or ("khóa" in loi_chi_tiet.lower())
            msg_id = send_telegram_notification(f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\nNội dung: {noi_dung}\n" + ("⏳ File đang bị khóa, bot sẽ tự thử lại nền..." if is_lock_error else f"❌ Lỗi ghi file:\n<code>{loi_chi_tiet}</code>"))
            if is_lock_error and msg_id and TELEGRAM_CHAT_ID:
                background_retry_write(so_tien, is_income, TELEGRAM_CHAT_ID, msg_id)
    except Exception as e:
        print(f"❌ LỖI XỬ LÝ GIAO DỊCH SEPAY: {str(e)}")

def send_telegram_notification(text):
    if not TOKEN or not TELEGRAM_CHAT_ID:
        return None
    try:
        resp = requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=15)
        return resp.json().get("result", {}).get("message_id")
    except Exception:
        return None

def edit_telegram_message(chat_id, message_id, text):
    if not TOKEN:
        return
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/editMessageText", json={"chat_id": chat_id, "message_id": message_id, "text": text, "parse_mode": "HTML"}, timeout=15)
    except Exception as e:
        print(f"Không sửa được tin nhắn Telegram: {e}")

def background_retry_write(amount, is_income, chat_id, message_id, attempt=1, max_attempts=5, delay_seconds=90):
    loai_text = "💰 Thu" if is_income else "💸 Chi"
    def _attempt():
        try:
            append_transaction_and_upload(amount, is_income)
            edit_telegram_message(chat_id, message_id, f"{loai_text}: <code>{amount:,.0f}</code> VNĐ\n✅ Đã ghi vào Excel thành công! (tự thử lại nền)")
        except Exception as e:
            if attempt < max_attempts:
                edit_telegram_message(chat_id, message_id, f"{loai_text}: <code>{amount:,.0f}</code> VNĐ\n⏳ File vẫn đang bị khóa, bot đang tự thử lại nền (lần {attempt}/{max_attempts})...")
                threading.Timer(delay_seconds, background_retry_write, args=(amount, is_income, chat_id, message_id, attempt + 1, max_attempts, delay_seconds)).start()
            else:
                edit_telegram_message(chat_id, message_id, f"{loai_text}: <code>{amount:,.0f}</code> VNĐ\n❌ Vẫn lỗi sau nhiều lần thử nền:\n<code>{str(e)}</code>\nVui lòng kiểm tra file Excel trên OneDrive rồi gửi lại số tiền.")
    threading.Thread(target=_attempt, daemon=True).start()

# ================== TELEGRAM BOT ==================
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[InlineKeyboardButton("📱 Lấy mã QR", callback_data="get_qr")], [InlineKeyboardButton("💰 Kiểm tra tiền", callback_data="check_money")], [InlineKeyboardButton("📝 Ghi sổ chi", callback_data="add_expense")]]
    await update.message.reply_text("🤖 Xin chào!\nNhập số <b>1</b> hoặc bấm menu dưới đây để chọn chức năng.\nNhập số <b>2</b> để đăng xuất.\n\n📝 Muốn ghi khoản <b>chi</b> vào Excel, bấm nút <b>Ghi sổ chi</b> — cần nhập thêm mật khẩu riêng.", reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="HTML")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if logged_in_users.get(user_id):
        await show_menu(update, context)
    else:
        await update.message.reply_text("🔐 <b>Menu được bảo vệ.</b>\nVui lòng nhập mật khẩu để tiếp tục:", parse_mode="HTML")

async def do_clear_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    try:
        member = await context.bot.get_chat_member(chat.id, update.effective_user.id)
        is_admin = member.status in ("administrator", "creator")
    except Exception:
        is_admin = False
    if not is_admin:
        await update.message.reply_text("❌ Chỉ admin của nhóm mới được dùng lệnh này.")
        return
    try:
        bot_member = await context.bot.get_chat_member(chat.id, context.bot.id)
        if not getattr(bot_member, "can_delete_messages", False):
            await update.message.reply_text("❌ Bot chưa có quyền xóa tin nhắn trong nhóm này.")
            return
    except Exception:
        pass
    warn = await update.message.reply_text("🧹 Đang xóa toàn bộ tin nhắn trong nhóm, vui lòng chờ...")
    deleted = await clear_chat_history(context, chat.id, warn.message_id)
    try:
        confirm = await context.bot.send_message(chat.id, f"✅ Đã xóa xong (khoảng {deleted} tin nhắn).")
        async def _self_delete():
            await asyncio.sleep(4)
            try:
                await context.bot.delete_message(chat_id=chat.id, message_id=confirm.message_id)
            except Exception:
                pass
        asyncio.create_task(_self_delete())
    except Exception:
        pass

async def clearall_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type not in ("group", "supergroup"):
        await update.message.reply_text("⚠️ Lệnh này chỉ dùng trong nhóm.")
        return
    await do_clear_group(update, context)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip() if update.message and update.message.text else ""
    if update.effective_chat.type in ("group", "supergroup") and text == "2":
        await do_clear_group(update, context)
        return
    if not logged_in_users.get(user_id):
        if text == BOT_PASSWORD:
            logged_in_users[user_id] = True
            await update.message.reply_text("✅ <b>Đăng nhập thành công!</b>", parse_mode="HTML")
            await show_menu(update, context)
        else:
            await update.message.reply_text("❌ <b>Sai mật khẩu.</b> Vui lòng thử lại:", parse_mode="HTML")
        return
    if text == "1":
        await show_menu(update, context)
        return
    if text == "2":
        logged_in_users[user_id] = False
        awaiting_expense_password.discard(user_id)
        expense_mode_users.discard(user_id)
        await clear_chat_history(context, update.effective_chat.id, update.message.message_id)
        sent = await context.bot.send_message(update.effective_chat.id, "🔒 <b>Bạn đã đăng xuất thành công. Menu đã được khóa lại!</b>", parse_mode="HTML")
        async def _self_delete():
            await asyncio.sleep(4)
            try:
                await context.bot.delete_message(chat_id=update.effective_chat.id, message_id=sent.message_id)
            except Exception:
                pass
        asyncio.create_task(_self_delete())
        return
    if user_id in awaiting_expense_password:
        if text == EXPENSE_PASSWORD:
            awaiting_expense_password.discard(user_id)
            expense_mode_users.add(user_id)
            await update.message.reply_text("✅ <b>Mở khóa sổ chi thành công!</b>\nGiờ hãy gõ số tiền để ghi khoản chi, ví dụ: <code>2.550.000</code>", parse_mode="HTML")
        else:
            await update.message.reply_text("❌ <b>Sai mật khẩu sổ chi.</b> Vui lòng thử lại:", parse_mode="HTML")
        return
    if user_id in expense_mode_users:
        amount = parse_amount_message(text)
        if amount is not None:
            status_msg = await update.message.reply_text(f"⏳ Đang ghi khoản chi <code>{amount:,.0f}</code> VNĐ vào Excel...", parse_mode="HTML")
            try:
                await asyncio.to_thread(append_transaction_and_upload, amount, False)
                await status_msg.edit_text(f"💸 Chi: <code>{amount:,.0f}</code> VNĐ\n✅ Đã ghi vào Excel thành công!", parse_mode="HTML")
            except Exception as e:
                error_text = str(e)
                if ("423" in error_text) or ("khóa" in error_text.lower()):
                    await status_msg.edit_text(f"💸 Chi: <code>{amount:,.0f}</code> VNĐ\n⏳ File đang bị khóa, bot sẽ tự thử lại nền...", parse_mode="HTML")
                    background_retry_write(amount, False, update.effective_chat.id, status_msg.message_id)
                else:
                    await status_msg.edit_text(f"💸 Chi: <code>{amount:,.0f}</code> VNĐ\n❌ Lỗi ghi file:\n<code>{error_text}</code>", parse_mode="HTML")
            return
    await update.message.reply_text("💡 Nếu muốn tìm Menu ấn số 1\nHoặc bấm nút \"Ghi sổ chi\" trong menu để ghi khoản chi.")

async def handle_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user_id = update.effective_user.id
    if not logged_in_users.get(user_id):
        await query.answer("Bạn chưa đăng nhập!", show_alert=True)
        return
    await query.answer()
    if query.data == "get_qr":
        tz_vn = pytz.timezone('Asia/Ho_Chi_Minh')
        current_time = datetime.now(tz_vn).strftime("%d/%m/%Y %H:%M:%S")
        media_group = []
        info_text = f"📅 <b>NGÀY TẠO MÃ:</b> <code>{current_time}</code>\n👤 <b>HỌ TÊN:</b> <code>HUYNH NGOC NGHIA</code>\n\n📌 <b>DANH SÁCH SỐ TÀI KHOẢN:</b>\n\n"
        for i, bank in enumerate(BANKS, 1):
            media_group.append(InputMediaPhoto(media=bank['qr_url'], caption=f"{bank['name']}\nSTK: {bank['account_number']}\nTên: {bank['account_name']}"))
            info_text += f"{i}. <b>{bank['name']}</b>\n   - STK: <code>{bank['account_number']}</code>\n   - Tên: <code>{bank['account_name']}</code>\n\n"
        await query.message.reply_media_group(media=media_group)
        await query.message.reply_text(info_text, parse_mode="HTML")
    elif query.data == "check_money":
        await query.message.reply_text("⏳ Đang tải dữ liệu từ OneDrive...")
        try:
            report = await asyncio.to_thread(get_excel_data)
            await query.message.reply_text(report, parse_mode="HTML")
        except Exception as e:
            await query.message.reply_text(f"❌ Lỗi lấy dữ liệu: {e}")
    elif query.data == "add_expense":
        if user_id in expense_mode_users:
            await query.message.reply_text("📝 Sổ chi đã mở khóa sẵn.\nGõ số tiền để ghi khoản chi, ví dụ: <code>2.550.000</code>", parse_mode="HTML")
        else:
            awaiting_expense_password.add(user_id)
            await query.message.reply_text("🔐 <b>Vui lòng nhập mật khẩu riêng để mở khóa sổ chi:</b>", parse_mode="HTML")

# ================== ERROR HANDLER ==================
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    print(f"⚠️ Exception: {context.error}")
    import traceback
    traceback.print_exception(type(context.error), context.error, context.error.__traceback__)

# ================== HÀM CHẠY BOT TELEGRAM ==================
def run_telegram_bot():
    if not TOKEN:
        print("LỖI: Chưa cài BOT_TOKEN!")
        return
    try:
        _validate_graph_config()
        print("✅ Env Graph API đầy đủ.")
    except Exception as e:
        print(f"❌ {e}")
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/deleteWebhook")
        print("Đã xóa webhook cũ!")
    except Exception:
        pass
    application = Application.builder().token(TOKEN).concurrent_updates(True).build()
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("xoahet", clearall_command))
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    application.add_handler(CallbackQueryHandler(handle_button))
    application.add_error_handler(error_handler)
    print("Bot đang chạy...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

# ================== MAIN ==================
# ================== MAIN ==================
def main():
    if not TOKEN:
        print("LỖI: Chưa cài BOT_TOKEN!")
        sys.exit(1)

    try:
        _validate_graph_config()
        print("✅ Env Graph API đầy đủ.")
    except Exception as e:
        print(f"❌ {e}")

    # 1. Chạy Flask trong một luồng (thread) nền
    port = int(os.environ.get("PORT", 10000))
    from waitress import serve
    flask_thread = threading.Thread(
        target=lambda: serve(app_web, host="0.0.0.0", port=port),
        daemon=True
    )
    flask_thread.start()
    print(f"Web server chạy cổng {port}")

    # 2. Xóa webhook cũ
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/deleteWebhook")
        print("Đã xóa webhook cũ!")
    except Exception:
        pass

    # 3. Tạo một event loop mới cho tiến trình chính (Fix cho Python 3.14)
    try:
        loop = asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    # 4. Khởi tạo và chạy Bot Telegram
    application = Application.builder().token(TOKEN).concurrent_updates(True).build()
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("xoahet", clearall_command))
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    application.add_handler(CallbackQueryHandler(handle_button))
    application.add_error_handler(error_handler)

    print("Bot đang chạy...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
