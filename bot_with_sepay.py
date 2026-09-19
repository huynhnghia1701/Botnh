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
EXPENSE_PASSWORD = os.getenv("EXPENSE_PASSWORD", "0939")  # Mật khẩu RIÊNG để mở khóa ghi sổ chi — đổi qua biến môi trường EXPENSE_PASSWORD nếu muốn

ONEDRIVE_URL = "https://1drv.ms/x/c/813BCA548F1AB473/IQDYUEgvvFlYRqhwhjmw-EFIAY0oGKUkTxQbKia9HGESO6o?download=1"
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")          
ONEDRIVE_TOKEN_CACHE = os.getenv("ONEDRIVE_TOKEN_CACHE", "")  
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "")      
SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")  
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")  

GRAPH_AUTHORITY = "https://login.microsoftonline.com/consumers"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]

app_web = Flask(__name__)
logged_in_users = {}
awaiting_expense_password = set()  # user_id đang chờ nhập mật khẩu riêng để mở khóa ghi sổ chi
expense_mode_users = set()  # user_id đã mở khóa, được phép gõ số tiền để ghi sổ chi
excel_write_lock = threading.Lock()  # đảm bảo chỉ ghi 1 giao dịch vào Excel tại 1 thời điểm

async def clear_chat_history(context, chat_id, from_message_id, max_delete=2000, max_consecutive_fail=40):
    """
    Xóa TOÀN BỘ lịch sử chat với bot mà KHÔNG cần lưu danh sách tin nhắn ở đâu cả.
    Vì message_id trong 1 chat luôn tăng dần, chỉ cần lùi dần từ tin nhắn hiện tại
    (from_message_id) xuống 1 và thử xóa từng ID. Dừng lại khi gặp nhiều ID liên tiếp
    không xóa được (nghĩa là đã hết lịch sử / vượt quá 48h) để tránh gọi API vô ích.
    """
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

# Regex nhận diện tin nhắn là một số tiền, ví dụ: "-2.550.000", "2.550.000", "+150000", "50k"
AMOUNT_PATTERN = re.compile(r'^([+-]?)\s*([\d.,]+)\s*([kK]?)$')

BANKS = [
    {"name": "Vietcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "96886693059121", "qr_url": "https://api.vietqr.io/image/MSB-96886693059121-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "OCB (Phương Đông)", "account_name": "HUYNH NGOC NGHIA", "account_number": "OCB-SEPHN48935", "qr_url": "https://api.vietqr.io/image/OCB-SEPHN48935-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "LPBank", "account_name": "HUYNH NGOC NGHIA", "account_number": "0916856322", "qr_url": "https://api.vietqr.io/image/970449-0916856322-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "Techcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "3838396852", "qr_url": "https://api.vietqr.io/image/970407-3838396852-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"}
]

# ================== HÀM ĐỌC EXCEL ==================
def get_excel_data():
    try:
        # Đọc CÙNG 1 FILE với chỗ ghi (qua Graph API + ONEDRIVE_FILE_PATH),
        # KHÔNG dùng link chia sẻ ONEDRIVE_URL cũ nữa (có thể trỏ nhầm file/bản khác)
        content = download_excel_for_write()
        excel_file = io.BytesIO(content)
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = workbook.active

        # Dữ liệu nằm ở dòng 2-27 (26 dòng), khớp đúng file Excel thật.
        # Tự cộng bằng Python từ dữ liệu THẬT (không đọc A29/B29/A31) để luôn
        # đúng ngay lập tức, không phụ thuộc việc file có được mở lại bằng Excel
        # để tính lại công thức hay chưa.
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

# ================== HÀM GHI EXCEL (TỐI ƯU NHẸ) ==================
def get_graph_access_token():
    if not ONEDRIVE_TOKEN_CACHE: raise Exception("ONEDRIVE_TOKEN_CACHE đang trống!")
    cache = msal.SerializableTokenCache()
    cache.deserialize(ONEDRIVE_TOKEN_CACHE)
    app = msal.PublicClientApplication(GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if not accounts: raise Exception("Token cache không có tài khoản nào!")
    result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    if not result or "access_token" not in result: raise Exception("Token hết hạn!")
    return result["access_token"]

def get_encoded_file_path():
    path = ONEDRIVE_FILE_PATH.strip().strip('/')
    if not path: raise Exception("ONEDRIVE_FILE_PATH đang trống!")
    if '/' in path:
        folder, filename = path.rsplit('/', 1)
        return f"{quote(folder)}:/{quote(filename)}"
    return quote(path)

def download_excel_for_write():
    token = get_graph_access_token()
    file_path = get_encoded_file_path()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content"
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
        resp.raise_for_status()
        return resp.content
    except Exception as e:
        raise Exception(f"Lỗi tải file: {str(e)}")

def upload_excel(content_bytes, max_retries=4):
    token = get_graph_access_token()
    file_path = get_encoded_file_path()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.put(url, headers=headers, data=content_bytes, timeout=30)
            if resp.status_code == 423:
                # File đang bị khóa (đang mở trong Excel / đang đồng bộ trên OneDrive).
                # Đợi rồi thử lại thay vì báo lỗi ngay, vì thường tự hết khóa sau vài giây.
                last_error = "File đang bị khóa trên OneDrive (423 Locked) — có thể đang mở trong Excel."
                print(f"⚠️ 423 Locked, thử lại lần {attempt}/{max_retries}...")
                time.sleep(2 * attempt)  # backoff: 2s, 4s, 6s, 8s
                continue
            if resp.status_code != 200 and resp.status_code != 201:
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
        f"Vui lòng đóng hẳn file Excel trên mọi thiết bị (Close file, không chỉ đóng tab) rồi gửi lại số tiền."
    )

def append_transaction_and_upload(amount, is_income):
    # Khóa lại: nếu 2 giao dịch đến gần nhau cùng lúc (kể cả từ SePay lẫn từ chat),
    # giao dịch thứ 2 phải CHỜ giao dịch thứ 1 ghi + upload xong hẳn mới được bắt đầu.
    # Tránh trường hợp cả 2 cùng tải bản cũ -> cùng ghi -> cái ghi sau đè mất cái ghi trước.
    with excel_write_lock:
        try:
            content = download_excel_for_write()
            workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
            sheet = workbook.active
            col = 1 if is_income else 2
            # Dữ liệu nằm ở dòng 2-27 (26 dòng), khớp đúng file Excel thật
            target_row = None
            for i in range(2, 28):
                if sheet.cell(row=i, column=col).value is None:
                    target_row = i
                    break
            if target_row is None: raise Exception("Hết chỗ trống trong bảng (đủ 26 dòng)! Cần dọn bớt Excel.")

            # CHỈ ghi số tiền giao dịch mới vào ô trống.
            # KHÔNG đụng tới A29/B29/A31 -> giữ nguyên công thức SUM có sẵn,
            # Excel sẽ tự tính lại đúng khi bạn mở file lên xem.
            sheet.cell(row=target_row, column=col).value = amount

            buf = io.BytesIO()
            workbook.save(buf)
            buf.seek(0)
            upload_excel(buf.read())

            del workbook, sheet, content, buf
            gc.collect()
            return True
        except Exception as e:
            # Ném lại lỗi NGUYÊN VĂN (không bọc thêm "Lỗi ghi file:") để
            # tin nhắn Telegram hiển thị đúng chi tiết lỗi gốc, dễ debug hơn.
            raise

# ================== HÀM PHÂN TÍCH SỐ TIỀN TỪ TIN NHẮN CHAT ==================
def parse_amount_message(text):
    """
    Nhận diện tin nhắn dạng số tiền để ghi vào cột CHI (cột B) mà thôi.
    Dấu +/- (nếu có) chỉ là ký hiệu, không quyết định thu/chi — mọi số tiền
    gõ vào chat đều được ghi là khoản CHI. Ví dụ:
      "-2.550.000" -> chi 2.550.000
      "2.550.000"  -> chi 2.550.000
      "50k"        -> chi 50.000
    Trả về amount:int hoặc None nếu không khớp.
    """
    if not text:
        return None
    match = AMOUNT_PATTERN.match(text.strip())
    if not match:
        return None
    _sign, number_part, k_suffix = match.groups()

    # Bỏ dấu chấm/phẩy phân cách hàng nghìn
    cleaned = number_part.replace('.', '').replace(',', '')
    if not cleaned.isdigit():
        return None

    amount = int(cleaned)
    if k_suffix:  # hỗ trợ viết tắt kiểu "50k" = 50.000
        amount *= 1000

    if amount <= 0:
        return None

    return amount

# ================== WEBHOOK SEPAY ==================
@app_web.route('/ping', methods=['GET'])
def ping():
    return "OK", 200

@app_web.route('/sepay-webhook', methods=['POST'])
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

        loai_text = f"💰 Nhận tiền (in)" if is_income else f"💸 Chi tiền (out)"
        
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
                    f"⏳ File đang bị khóa, bot sẽ tự thử lại nền trong vài phút, không cần làm gì thêm..."
                    if is_lock_error else
                    f"❌ Lỗi ghi file:\n<code>{loi_chi_tiet}</code>"
                )
            )
            if is_lock_error and msg_id and TELEGRAM_CHAT_ID:
                background_retry_write(so_tien, is_income, TELEGRAM_CHAT_ID, msg_id)
    except Exception as e:
        print(f"❌ LỖI XỬ LÝ GIAO DỊCH SEPAY: {str(e)}")

def send_telegram_notification(text):
    if not TOKEN or not TELEGRAM_CHAT_ID: return None
    try:
        resp = requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=15)
        return resp.json().get("result", {}).get("message_id")
    except:
        print("Không gửi được tin nhắn Telegram.")
        return None

def edit_telegram_message(chat_id, message_id, text):
    if not TOKEN: return
    try:
        requests.post(
            f"https://api.telegram.org/bot{TOKEN}/editMessageText",
            json={"chat_id": chat_id, "message_id": message_id, "text": text, "parse_mode": "HTML"},
            timeout=15
        )
    except Exception as e:
        print(f"Không sửa được tin nhắn Telegram: {e}")

def background_retry_write(amount, is_income, chat_id, message_id, attempt=1, max_attempts=5, delay_seconds=90):
    """
    Dùng khi ghi Excel thất bại do file bị khóa (423) ngay cả sau các lần thử nhanh.
    Thử lại NGẦM thêm vài lần, cách nhau delay_seconds, vì khóa phía OneDrive/Excel
    (đồng bộ, xem trước, phiên co-authoring còn sót) có thể mất vài phút mới tự hết,
    lâu hơn khoảng retry nhanh trong upload_excel(). Người dùng không cần gửi lại số tiền,
    tin nhắn gốc sẽ tự được cập nhật khi có kết quả.
    """
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
    if logged_in_users.get(user_id): await show_menu(update, context)
    else:
        await update.message.reply_text("🔐 <b>Menu được bảo vệ.</b>\nVui lòng nhập mật khẩu để tiếp tục:", parse_mode="HTML")

async def do_clear_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Xóa toàn bộ lịch sử tin nhắn gần đây trong NHÓM (của mọi người, không chỉ của bot).
    CHỈ admin của nhóm mới dùng được, để tránh thành viên thường phá nhóm. Bot phải
    được cấp quyền admin + "Xóa tin nhắn" (Delete Messages) trong nhóm thì mới xóa
    được tin của người khác; nếu chưa có quyền, bot chỉ xóa được tin của chính nó.
    """
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
    # Lùi dần từ CHÍNH tin nhắn "Đang xóa..." này (message_id cao nhất tại thời điểm gọi)
    # để xóa luôn cả nó, tin nhắn lệnh, và toàn bộ lịch sử phía trước.
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
    """Lệnh /xoahet: chỉ dùng trong nhóm."""
    if update.effective_chat.type not in ("group", "supergroup"):
        await update.message.reply_text("⚠️ Lệnh này chỉ dùng trong nhóm, không dùng trong chat riêng.")
        return
    await do_clear_group(update, context)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip() if update.message and update.message.text else ""

    # ---- Trong NHÓM: gõ "2" = xóa toàn bộ tin nhắn nhóm, KHÔNG chạy luồng menu/đăng nhập của bot qr ----
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
        # Xóa TOÀN BỘ lịch sử chat này (lùi dần từ tin nhắn "2" hiện tại), không cần lưu gì cả.
        await clear_chat_history(context, chat_id, current_msg_id)
        sent = await context.bot.send_message(chat_id, "🔒 <b>Bạn đã đăng xuất thành công. Menu đã được khóa lại!</b>", parse_mode="HTML")
        # Tự xóa luôn tin nhắn xác nhận này sau vài giây để cuộc trò chuyện sạch hoàn toàn.
        async def _self_delete():
            await asyncio.sleep(4)
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=sent.message_id)
            except Exception as e:
                print(f"Không tự xóa được tin nhắn xác nhận đăng xuất: {e}")
        asyncio.create_task(_self_delete())
        return

    # ---- Bước nhập mật khẩu RIÊNG để mở khóa ghi sổ chi ----
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

    # ---- Nhận diện tin nhắn là số tiền để ghi vào Excel (LUÔN ghi cột CHI) ----
    # CHỈ áp dụng khi user đã mở khóa qua nút "Ghi sổ chi" + mật khẩu riêng.
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
                    # Không bắt người dùng gửi lại: tự thử lại NGẦM thêm vài lần trong vài phút,
                    # rồi cập nhật lại đúng tin nhắn này khi có kết quả cuối cùng.
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
            # SỬA LỖI TREO: chạy hàm blocking trong thread pool nhưng
            # VẪN Ở TRONG event loop hiện tại (không tạo asyncio.run() ở thread khác
            # -> tránh deadlock khi gọi API Telegram từ 2 event loop khác nhau)
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

# ================== ERROR HANDLER ==================
async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """
    Bắt tất cả lỗi phát sinh trong quá trình xử lý update (kể cả lỗi mạng tạm thời
    khi polling, ví dụ network_retry_loop). Chỉ log ra, không làm crash bot.
    Trước đây không có handler này nên log Render hiện "No error handlers are registered".
    """
    print(f"⚠️ Exception khi xử lý update: {context.error}")
    import traceback
    traceback.print_exception(type(context.error), context.error, context.error.__traceback__)

# ================== MAIN ==================
def main():
    if not TOKEN:
        print("LỖI: Chưa cài đặt BOT_TOKEN!")
        sys.exit(1)

    port = int(os.environ.get("PORT", 10000))
    threading.Thread(target=lambda: app_web.run(host='0.0.0.0', port=port, debug=False, use_reloader=False), daemon=True).start()

    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/deleteWebhook")
        print("Đã xóa webhook cũ!")
    except:
        pass

    # concurrent_updates=True: cho phép xử lý nhiều tin nhắn/nút bấm CÙNG LÚC,
    # để 1 request chậm (vd tải Excel) không làm "treo" toàn bộ bot với người khác
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
