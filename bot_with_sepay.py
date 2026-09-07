import os
import sys
import io
import gc
import threading
import time
from datetime import datetime
from urllib.parse import quote
import pytz
import requests
import openpyxl
from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes

# ================== CẤU HÌNH ==================
TOKEN = os.getenv("BOT_TOKEN")
BOT_PASSWORD = "123123"

# Link OneDrive dạng tải về trực tiếp (dùng để ĐỌC)
ONEDRIVE_URL = "https://1drv.ms/x/c/813BCA548F1AB473/IQDYUEgvvFlYRqhwhjmw-EFIAY0oGKUkTxQbKia9HGESO6o?download=1"

SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")  
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")  

app_web = Flask(__name__)
logged_in_users = {}

BANKS = [
    {"name": "Vietcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "96886693059121", "qr_url": "https://api.vietqr.io/image/MSB-96886693059121-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "OCB (Phương Đông)", "account_name": "HUYNH NGOC NGHIA", "account_number": "OCB-SEPHN48935", "qr_url": "https://api.vietqr.io/image/OCB-SEPHN48935-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "LPBank", "account_name": "HUYNH NGOC NGHIA", "account_number": "0916856322", "qr_url": "https://api.vietqr.io/image/970449-0916856322-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "Techcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "3838396852", "qr_url": "https://api.vietqr.io/image/970407-3838396852-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"}
]

# ================== HÀM ĐỌC EXCEL (TỐI ƯU NHẸ) ==================
def get_excel_data():
    try:
        response = requests.get(ONEDRIVE_URL, timeout=10)
        if response.status_code != 200:
            return f"❌ Lỗi tải file Excel từ OneDrive (Mã lỗi: {response.status_code})"
        
        excel_file = io.BytesIO(response.content)
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = workbook.active

        # Tính tổng bằng Python (Không phụ thuộc công thức Excel)
        thu_total = sum([sheet.cell(row=i, column=1).value or 0 for i in range(2, 26) if isinstance(sheet.cell(row=i, column=1).value, (int, float))])
        chi_total = sum([sheet.cell(row=i, column=2).value or 0 for i in range(2, 26) if isinstance(sheet.cell(row=i, column=2).value, (int, float))])
        remain = thu_total - chi_total

        thu_list = [sheet.cell(row=i, column=1).value for i in range(2, 26) if sheet.cell(row=i, column=1).value is not None]
        chi_list = [sheet.cell(row=i, column=2).value for i in range(2, 26) if sheet.cell(row=i, column=2).value is not None]

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
        
        # Giải phóng RAM ngay lập tức
        del workbook, sheet, excel_file
        gc.collect()
        return msg
    except Exception as e:
        return f"❌ Lỗi xử lý file Excel: {str(e)}"

# ================== WEBHOOK SEPAY (CHỈ NHẬN VÀ BÁO CÁO) ==================
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

        loai_text = f"💰 Nhận tiền (in)" if is_income else f"💸 Chi tiền (out)"
        
        send_telegram_notification(
            f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
            f"Nội dung: {noi_dung}\n"
            f"----------------------------------------\n"
            f"✅ Đã nhận giao dịch!"
        )
    except Exception as e:
        print(f"❌ LỖI XỬ LÝ GIAO DỊCH SEPAY: {str(e)}")

def send_telegram_notification(text):
    if not TOKEN or not TELEGRAM_CHAT_ID: return
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=15)
    except: print("Không gửi được tin nhắn Telegram.")

# ================== HÀM TELEGRAM BOT ==================
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[InlineKeyboardButton("📱 Lấy mã QR", callback_data="get_qr")], [InlineKeyboardButton("💰 Kiểm tra tiền", callback_data="check_money")]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("🤖 Xin chào!\nNhập số <b>1</b> hoặc bấm menu dưới đây để chọn chức năng.\nNhập số <b>2</b> để đăng xuất.", reply_markup=reply_markup, parse_mode="HTML")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if logged_in_users.get(user_id): await show_menu(update, context)
    else: await update.message.reply_text("🔐 <b>Menu được bảo vệ.</b>\nVui lòng nhập mật khẩu để tiếp tục:", parse_mode="HTML")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip() if update.message and update.message.text else ""
    if not logged_in_users.get(user_id):
        if text == BOT_PASSWORD:
            logged_in_users[user_id] = True
            await update.message.reply_text("✅ <b>Đăng nhập thành công!</b>", parse_mode="HTML")
            await show_menu(update, context)
        else:
            await update.message.reply_text("❌ <b>Sai mật khẩu.</b> Vui lòng thử lại:", parse_mode="HTML")
        return
    if text == "1": await show_menu(update, context)
    elif text == "2":
        logged_in_users[user_id] = False
        await update.message.reply_text("🔒 <b>Bạn đã đăng xuất thành công. Menu đã được khóa lại!</b>", parse_mode="HTML")
    else: await update.message.reply_text("💡 Nếu Muốn Tìm Menu Ấn Số 1")

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
        
        def fetch_and_reply():
            try:
                report = get_excel_data()
                import asyncio
                asyncio.run(query.message.reply_text(report, parse_mode="HTML"))
            except Exception as e:
                print(f"Lỗi gửi báo cáo: {e}")
        
        threading.Thread(target=fetch_and_reply, daemon=True).start()

# ================== MAIN (MỘT TIẾN TRÌNH DUY NHẤT) ==================
def main():
    if not TOKEN:
        print("LỖI: Chưa cài đặt BOT_TOKEN!")
        sys.exit(1)

    # Chạy Flask trong 1 luồng (Thread) riêng - Nhẹ hơn nhiều so với Process
    port = int(os.environ.get("PORT", 10000))
    threading.Thread(target=lambda: app_web.run(host='0.0.0.0', port=port, debug=False, use_reloader=False), daemon=True).start()

    # Xóa Webhook cũ chống đứng máy
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/deleteWebhook")
        print("Đã xóa webhook cũ!")
    except:
        pass

    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    application.add_handler(CallbackQueryHandler(handle_button))

    print("Bot đang chạy...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()