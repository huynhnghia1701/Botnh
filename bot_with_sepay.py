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

# ================== CẤU HÌNH ==================
TOKEN = os.getenv("BOT_TOKEN")
BOT_PASSWORD = "123123"
EXPENSE_PASSWORD = os.getenv("EXPENSE_PASSWORD", "0939")

GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "32a9cfe8-da08-4233-8446-42b0a4da1547")          
ONEDRIVE_TOKEN_CACHE = os.getenv("ONEDRIVE_TOKEN_CACHE", "")  
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "Sổ Thu Chi Mới.xlsx")      
SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")  
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")  

GRAPH_AUTHORITY = "https://login.microsoftonline.com/consumers"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]

# Biến lưu cache trong RAM để giữ Refresh Token mới nhất suốt vòng đời server
runtime_token_cache_str = ONEDRIVE_TOKEN_CACHE

app_web = Flask(__name__)
logged_in_users = {}
awaiting_expense_password = set()
expense_mode_users = set()
excel_write_lock = threading.Lock()

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

AMOUNT_PATTERN = re.compile(r'^([+-]?)\s*([\d.,]+)\s*([kK]?)$')

BANKS = [
    {"name": "Vietcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "96886693059121", "qr_url": "https://api.vietqr.io/image/MSB-96886693059121-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "OCB (Phương Đông)", "account_name": "HUYNH NGOC NGHIA", "account_number": "OCB-SEPHN48935", "qr_url": "https://api.vietqr.io/image/OCB-SEPHN48935-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "LPBank", "account_name": "HUYNH NGOC NGHIA", "account_number": "0916856322", "qr_url": "https://api.vietqr.io/image/970449-0916856322-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "Techcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "3838396852", "qr_url": "https://api.vietqr.io/image/970407-3838396852-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"}
]

# ================== HÀM MICROSOFT GRAPH API ==================
def get_graph_access_token():
    global runtime_token_cache_str
    if not runtime_token_cache_str:
        raise Exception("ONEDRIVE_TOKEN_CACHE đang trống!")
    
    cache = msal.SerializableTokenCache()
    cache.deserialize(runtime_token_cache_str)
    app = msal.PublicClientApplication(GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if not accounts:
        raise Exception("Token cache không có tài khoản nào!")
    
    result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise Exception("Token hết hạn hoặc bị thu hồi! Cần tạo lại ONEDRIVE_TOKEN_CACHE.")
    
    if cache.has_state_changed:
        runtime_token_cache_str = cache.serialize()
        
    return result["access_token"]

def get_encoded_file_path():
    path = ONEDRIVE_FILE_PATH.strip().strip('/')
    if not path:
        raise Exception("ONEDRIVE_FILE_PATH đang trống!")
    if '/' in path:
        folder, filename = path.rsplit('/', 1)
        return f"{quote(folder)}:/{quote(filename)}"
    return quote(path)

def download_excel_for_write():
    token = get_graph_access_token()
    file_path = get_encoded_file_path()
    
    # Lấy metadata và trường downloadUrl trực tiếp để tránh lỗi 401 khi redirect sang SharePoint
    meta_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}"
    meta_resp = requests.get(meta_url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
    
    if meta_resp.status_code != 200:
        raise Exception(f"Không tìm thấy file '{ONEDRIVE_FILE_PATH}' ({meta_resp.status_code}): {meta_resp.text}")
        
    download_url = meta_resp.json().get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise Exception("Không lấy được link tải trực tiếp từ Graph API.")
        
    resp = requests.get(download_url, timeout=20)
    resp.raise_for_status()
    return resp.content

def upload_excel(content_bytes, max_retries=4):
    token = get_graph_access_token()
    file_path = get_encoded_file_path()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.put(url, headers=headers, data=content_bytes, timeout=30)
            if resp.status_code == 423:
                last_error = "File đang bị khóa trên OneDrive (423 Locked) — có thể đang mở trong Excel."
                print(f"⚠️ 423 Locked, thử lại lần {attempt}/{max_retries}...")
                time.sleep(2 * attempt)
                continue
            if resp.status_code not in (200, 201):
                print(f"❌ Lỗi Graph API chi tiết: {resp.text}")
                resp.raise_for_status()
            return resp.json()
        except requests.exceptions.HTTPError:
            raise
        except Exception as e:
            last_error = str(e)
            raise Exception(f"Lỗi upload: {last_error}")

    raise Exception(
        f"{last_error} Đã thử lại {max_retries} lần nhưng vẫn bị khóa. "
        f"Vui lòng đóng hẳn file Excel trên mọi thiết bị rồi gửi lại số tiền."
    )

# ================== HÀM ĐỌC / GHI EXCEL ==================
def get_excel_data():
    try:
        content = download_excel_for_write()
        excel_file = io.BytesIO(content)
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = workbook.active

        thu_list = [sheet.cell(row=i, column=1).value for i in range(2, 28) if sheet.cell(row=i, column=1).value is not None]
        chi_list = [sheet.cell(row=i, column=2).value for i in range(2, 28) if sheet.cell(row=i, column=2).value is not None]
        thu_total = sum(v for v in thu_list if isinstance(v, (int, float)))
        chi_total = sum(v for v in chi_list if isinstance(v, (int, float)))
        remain = thu_total - chi_total

        msg = "📊 <b>BÁO CÁO SỔ THU CHI (ONEDRIVE)</b>\n----------------------------------------\n📥 <b>DANH SÁCH THU:</b>\n"
        for val in thu_list:
            msg += f"  • {val:,.0f} VNĐ\n" if isinstance(val, (int, float)) else f"  • {val}\n"
        msg += "\n📤 <b>DANH SÁCH CHI:</b>\n"
        for val in chi_list:
            msg += f"  • {val:,.0f} VNĐ\n" if isinstance(val, (int, float)) else f"  • {val}\n"
        msg += "----------------------------------------\n"
        msg += f"🟢 <b>Tổng Tiền Nhận Vào:</b> <code>{thu_total:,.0f}</code> VNĐ\n"
        msg += f"🔴 <b>Tổng Tiền Đã Chi:</b> <code>{chi_total:,.0f}</code> VNĐ\n"
        msg += f"💰 <b>SỐ TIỀN CÒN LẠI:</b> <code>{remain:,.0f}</code> VNĐ\n"
        
        del workbook, sheet, excel_file
        gc.collect()
        return msg
    except requests.exceptions.Timeout:
        return "❌ Lỗi: Tải file Excel quá chậm (Timeout)."
    except Exception as e:
        return f"❌ Lỗi xử lý file Excel: {str(e)}"

def append_transaction_and_upload(amount, is_income):
    with excel_write_lock:
        content = download_excel_for_write()
        workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
        sheet = workbook.active
        col = 1 if is_income else 2
        
        target_row = None
        for i in range(2, 28):
            if sheet.cell(row=i, column=col).value is None:
                target_row = i
                break
        if target_row is None:
            raise Exception("Hết chỗ trống trong bảng (đủ 26 dòng)! Cần dọn bớt Excel.")

        sheet.cell(row=target_row, column=col).value = amount

        buf = io.BytesIO()
        workbook.save(buf)
        buf.seek(0)
        upload_excel(buf.read())

        del workbook, sheet, content, buf
        gc.collect()
        return True

# ================== HÀM PHÂN TÍCH SỐ TIỀN TỪ TIN NHẮN CHAT ==================
def parse_amount_message(text):
    if not text:
        return None
    match = AMOUNT_PATTERN.match(text.strip())
    if not match:
        return None
    _sign, number_part, k_suffix = match.groups()

    cleaned = number_part.replace('.', '').replace(',', '')
    if not cleaned.isdigit():
        return None

    amount = int(cleaned)
    if k_suffix:
        amount *= 1000

    if amount <= 0:
        return None

    return amount

# ================== WEBHOOK SEPAY ==================
@app_web.route('/ping', methods=['GET'])
def ping():
    return "OK", 200

@app_web.route('/sepay-webhook', methods=['GET', 'POST'])
def sepay_webhook():
    if request.method == 'GET':
        return jsonify({"status": "alive", "message": "Server Render đang hoạt động!"}), 200
        
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
                    "⏳ File đang bị khóa, bot sẽ tự thử lại nền trong vài phút, không cần làm gì thêm..."
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
            timeout=15
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
            timeout=15
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
                f"{loai_text}: <code>{amount:,.0f}</code> VNĐ\n"
                f"✅ Đã ghi vào Excel thành công! (tự thử lại nền)"
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

# ================== HÀM TELEGRAM BOT ==================
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📱 Lấy mã QR", callback_data="get_qr")],
        [InlineKeyboardButton("💰 Kiểm tra tiền", callback_data="check_money")],
        [InlineKeyboardButton("📝 Ghi sổ chi", callback_data="add_expense")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        "🤖 Xin chào!\nNhập số <b>1</b> hoặc bấm menu dưới đây để chọn chức năng.\nNhập số <b>2</b> để đăng xuất.\n\n"
        "📝 Muốn ghi khoản <b>chi</b> vào Excel, bấm nút <b>Ghi sổ chi</b> bên dưới — cần nhập thêm mật khẩu riêng.",
        reply_markup=reply_markup, parse_mode="HTML"
    )

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if logged_in_users.get(user_id):
        await show_menu(update, context)
    else:
        await update.message.reply_text("🔐 <b>Menu được bảo vệ.</b>\nVui lòng nhập mật khẩu để tiếp tục:", parse_mode="HTML")

async def do_clear_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat = update.effective_chat
    user_id = update.effective_user.id
    try:
        member = await context.bot.get_chat_member(chat.id, user_id)
        is_admin = member.status in ("administrator", "creator")
    except Exception as e:
        is_admin = False
        print(f"Không kiểm tra được quyền admin: {e}")

    if not is_admin:
        await update.message.reply_text("❌ Chỉ admin của nhóm mới được dùng lệnh này.")
        return

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
        await update.message.reply_text("⚠️ Lệnh này chỉ dùng trong nhóm, không dùng trong chat riêng.")
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
        chat_id = update.effective_chat.id
        current_msg_id = update.message.message_id
        await clear_chat_history(context, chat_id, current_msg_id)
        sent = await context.bot.send_message(chat_id, "🔒 <b>Bạn đã đăng xuất thành công. Menu đã được khóa lại!</b>", parse_mode="HTML")
        async def _self_delete():
            await asyncio.sleep(4)
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=sent.message_id)
            except Exception as e:
                print(f"Không tự xóa được tin nhắn xác nhận đăng xuất: {e}")
        asyncio.create_task(_self_delete())
        return

    if user_id in awaiting_expense_password:
        if text == EXPENSE_PASSWORD:
            awaiting_expense_password.discard(user_id)
            expense_mode_users.add(user_id)
            await update.message.reply_text(
                "✅ <b>Mở khóa sổ chi thành công!</b>\n"
                "Giờ hãy gõ số tiền để ghi khoản chi, ví dụ: <code>2.550.000</code> hoặc <code>-2.550.000</code>",
                parse_mode="HTML"
            )
        else:
            await update.message.reply_text("❌ <b>Sai mật khẩu sổ chi.</b> Vui lòng thử lại:", parse_mode="HTML")
        return

    if user_id in expense_mode_users:
        amount = parse_amount_message(text)
        if amount is not None:
            status_msg = await update.message.reply_text(f"⏳ Đang ghi khoản chi <code>{amount:,.0f}</code> VNĐ vào Excel...", parse_mode="HTML")
            try:
                await asyncio.to_thread(append_transaction_and_upload, amount, False)
                await status_msg.edit_text(
                    f"💸 Chi: <code>{amount:,.0f}</code> VNĐ\n"
                    f"✅ Đã ghi vào Excel thành công!",
                    parse_mode="HTML"
                )
            except Exception as e:
                error_text = str(e)
                is_lock_error = ("423" in error_text) or ("khóa" in error_text.lower())
                if is_lock_error:
                    await status_msg.edit_text(
                        f"💸 Chi: <code>{amount:,.0f}</code> VNĐ\n"
                        f"⏳ File đang bị khóa, bot sẽ tự thử lại nền trong vài phút, không cần gửi lại...",
                        parse_mode="HTML"
                    )
                    background_retry_write(amount, False, update.effective_chat.id, status_msg.message_id)
                else:
                    await status_msg.edit_text(
                        f"💸 Chi: <code>{amount:,.0f}</code> VNĐ\n"
                        f"❌ Lỗi ghi file:\n<code>{error_text}</code>",
                        parse_mode="HTML"
                    )
            return

    await update.message.reply_text("💡 Nếu Muốn Tìm Menu Ấn Số 1\nHoặc bấm nút \"Ghi sổ chi\" trong menu để ghi khoản chi.")

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
            caption_single = f"{bank['name']}\nSTK: {bank['account_number']}\nTên: {bank['account_name']}"
            media_group.append(InputMediaPhoto(media=bank['qr_url'], caption=caption_single))
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
            await query.message.reply_text(
                "📝 Sổ chi đã mở khóa sẵn.\nGõ số tiền để ghi khoản chi, ví dụ: <code>2.550.000</code>",
                parse_mode="HTML"
            )
        else:
            awaiting_expense_password.add(user_id)
            await query.message.reply_text("🔐 <b>Vui lòng nhập mật khẩu riêng để mở khóa sổ chi:</b>", parse_mode="HTML")

# ================== KEEP-ALIVE PING ==================
def keep_alive():
    render_url = "https://botnh-1.onrender.com/ping"
    while True:
        time.sleep(600)
        try:
            requests.get(render_url, timeout=10)
        except Exception:
            pass

# ================== ERROR HANDLER ==================
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    print(f"⚠️ Exception khi xử lý update: {context.error}")
    import traceback
    traceback.print_exception(type(context.error), context.error, context.error.__traceback__)

# ================== MAIN ==================
def main():
    if not TOKEN:
        print("LỖI: Chưa cài đặt BOT_TOKEN!")
        sys.exit(1)

    threading.Thread(target=keep_alive, daemon=True).start()

    port = int(os.environ.get("PORT", 10000))
    threading.Thread(target=lambda: app_web.run(host='0.0.0.0', port=port, debug=False, use_reloader=False), daemon=True).start()

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

if __name__ == "__main__":
    main()
