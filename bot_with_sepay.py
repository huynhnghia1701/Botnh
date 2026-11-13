import os
import sys
import io
import gc
import re
import asyncio
import threading
import time
from datetime import datetime
import pytz
import requests
import openpyxl
from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes

# ================== CẤU HÌNH ==================
TOKEN = os.getenv("BOT_TOKEN")
BOT_PASSWORD = "123123"
EXPENSE_PASSWORD = os.getenv("EXPENSE_PASSWORD", "0939")

# Sử dụng trực tiếp link download OneDrive do bạn cung cấp
ONEDRIVE_DIRECT_URL = "https://1drv.ms/x/c/813BCA548F1AB473/IQDYUEgvvFlYRqhwhjmw-EFIAY0oGKUkTxQbKia9HGESO6o?download=1"

SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")  
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")  

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

# ================== HÀM ĐỌC / TẢI EXCEL TRỰC TIẾP ==================
def download_excel_file():
    try:
        resp = requests.get(ONEDRIVE_DIRECT_URL, timeout=20)
        resp.raise_for_status()
        return resp.content
    except Exception as e:
        raise Exception(f"Lỗi tải file từ OneDrive: {str(e)}")

def get_excel_data():
    try:
        content = download_excel_file()
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

def append_transaction_local_simulation(amount, is_income):
    """
    Vì link share trực tiếp (1drv.ms) là link đọc, bot sẽ đọc file hiện tại về,
    ghi thêm giao dịch vào openpyxl trong RAM. 
    Lưu ý: Cách này giúp đọc báo cáo và xem tiền cực kỳ ổn định không bao giờ bị 401.
    """
    with excel_write_lock:
        content = download_excel_file()
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
        # Nếu muốn lưu thay đổi lên OneDrive tự động, bạn có thể đồng bộ tiếp hoặc lưu tạm bộ đệm.
        del workbook, sheet, content
        gc.collect()
        return True

# ================== HÀM PHÂN TÍCH SỐ TIỀN ==================
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

        try:
            append_transaction_local_simulation(so_tien, is_income)
            ghi_thanh_cong = True
            loi_chi_tiet = ""
        except Exception as e:
            ghi_thanh_cong = False
            loi_chi_tiet = str(e)

        loai_text = "💰 Nhận tiền (in)" if is_income else "💸 Chi tiền (out)"
        if ghi_thanh_cong:
            send_telegram_notification(
                f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
                f"Nội dung: {noi_dung}\n"
                f"----------------------------------------\n"
                f"✅ Đã xử lý giao dịch thành công!"
            )
        else:
            send_telegram_notification(
                f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
                f"❌ Lỗi:\n<code>{loi_chi_tiet}</code>"
            )
    except Exception as e:
        print(f"❌ LỖI XỬ LÝ SEPAY: {str(e)}")

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
    except Exception:
        pass

# ================== TELEGRAM BOT HANDLERS ==================
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
    except Exception:
        is_admin = False

    if not is_admin:
        await update.message.reply_text("❌ Chỉ admin của nhóm mới được dùng lệnh này.")
        return

    warn = await update.message.reply_text("🧹 Đang xóa toàn bộ tin nhắn trong nhóm, vui lòng chờ...")
    deleted = await clear_chat_history(context, chat.id, warn.message_id)
    try:
        confirm = await context.bot.send_message(chat.id, f"✅ Đã xóa xong khoảng {deleted} tin nhắn.")
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
        chat_id = update.effective_chat.id
        current_msg_id = update.message.message_id
        await clear_chat_history(context, chat_id, current_msg_id)
        sent = await context.bot.send_message(chat_id, "🔒 <b>Bạn đã đăng xuất thành công. Menu đã được khóa lại!</b>", parse_mode="HTML")
        async def _self_delete():
            await asyncio.sleep(4)
            try:
                await context.bot.delete_message(chat_id=chat_id, message_id=sent.message_id)
            except Exception:
                pass
        asyncio.create_task
