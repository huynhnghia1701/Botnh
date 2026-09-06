import os
import sys
import io
import json
import threading
from datetime import datetime
import pytz
import requests
import msal
import openpyxl
from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes

# ================== CẤU HÌNH MÔI TRƯỜNG ==================
TOKEN = os.getenv("BOT_TOKEN")

# Link OneDrive dạng tải về trực tiếp (dùng để ĐỌC, giữ nguyên như cũ)
ONEDRIVE_URL = "https://1drv.ms/x/c/813BCA548F1AB473/IQDYUEgvvFlYRqhwhjmw-EFIASURiHvdRqvzgy28bqT6g0s?download=1"

# ====== CẤU HÌNH MỚI - dùng để GHI vào OneDrive qua Graph API ======
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")          
ONEDRIVE_TOKEN_CACHE = os.getenv("ONEDRIVE_TOKEN_CACHE", "")  
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "")      

SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")  
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")  

GRAPH_AUTHORITY = "https://login.microsoftonline.com/consumers"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]

app_web = Flask(__name__)

# ================== CẤU HÌNH NGÂN HÀNG ==================
BANKS = [
    {
        "name": "Vietcombank",
        "account_name": "HUYNH NGOC NGHIA",
        "account_number": "96886693059121",
        "qr_url": "https://api.vietqr.io/image/MSB-96886693059121-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"
    },
    {
        "name": "OCB (Phương Đông)",
        "account_name": "HUYNH NGOC NGHIA",
        "account_number": "OCB-SEPHN48935",
        "qr_url": "https://api.vietqr.io/image/OCB-SEPHN48935-compact2.png?accountName=HUYNH%20NGOC%20NGHIA",
    },
    {
        "name": "LPBank",
        "account_name": "HUYNH NGOC NGHIA",
        "account_number": "0916856322",
        "qr_url": "https://api.vietqr.io/image/970449-0916856322-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"
    },
    {
        "name": "Techcombank",
        "account_name": "HUYNH NGOC NGHIA",
        "account_number": "3838396852",
        "qr_url": "https://api.vietqr.io/image/970407-3838396852-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"
    }
]

# ================== HÀM ĐỌC EXCEL ==================
def get_excel_data():
    try:
        response = requests.get(ONEDRIVE_URL, timeout=10)
        if response.status_code != 200:
            return f"❌ Lỗi tải file Excel từ OneDrive (Mã lỗi: {response.status_code})"

        excel_file = io.BytesIO(response.content)
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = workbook.active

        # Lấy dữ liệu (đã có sẵn ô tổng và ô còn lại)
        thu_total = sheet['A29'].value or 0
        chi_total = sheet['B29'].value or 0
        remain = sheet['A31'].value or 0

        thu_list = [sheet.cell(row=i, column=1).value for i in range(2, 26) if sheet.cell(row=i, column=1).value is not None]
        chi_list = [sheet.cell(row=i, column=2).value for i in range(2, 26) if sheet.cell(row=i, column=2).value is not None]

        msg = "📊 <b>BÁO CÁO SỔ THU CHI (ONEDRIVE)</b>\n"
        msg += "----------------------------------------\n"
        msg += "📥 <b>DANH SÁCH THU:</b>\n"
        for val in thu_list:
            msg += f"  • {val:,.0f} VNĐ\n" if isinstance(val, (int, float)) else f"  • {val}\n"

        msg += "\n📤 <b>DANH SÁCH CHI:</b>\n"
        for val in chi_list:
            msg += f"  • {val:,.0f} VNĐ\n" if isinstance(val, (int, float)) else f"  • {val}\n"

        msg += "----------------------------------------\n"
        msg += f"🟢 <b>Tổng Tiền Nhận Vào:</b> <code>{thu_total:,.0f}</code> VNĐ\n" if isinstance(thu_total, (int, float)) else f"🟢 <b>Tổng Tiền Nhận Vào:</b> <code>{thu_total}</code> VNĐ\n"
        msg += f"🔴 <b>Tổng Tiền Đã Chi:</b> <code>{chi_total:,.0f}</code> VNĐ\n" if isinstance(chi_total, (int, float)) else f"🔴 <b>Tổng Tiền Đã Chi:</b> <code>{chi_total}</code> VNĐ\n"
        msg += f"💰 <b>SỐ TIỀN CÒN LẠI:</b> <code>{remain:,.0f}</code> VNĐ\n" if isinstance(remain, (int, float)) else f"💰 <b>SỐ TIỀN CÒN LẠI:</b> <code>{remain}</code> VNĐ\n"

        return msg
    except requests.exceptions.Timeout:
        return "❌ Lỗi: Yêu cầu tải file Excel hết thời gian chờ (Timeout)."
    except Exception as e:
        return f"❌ Lỗi xử lý file Excel: {str(e)}"

# ================== HÀM GHI EXCEL QUA GRAPH API ==================
def get_graph_access_token():
    """Lấy access token bằng token cache đã xin quyền từ trước."""
    if not ONEDRIVE_TOKEN_CACHE:
        raise Exception("Chưa cấu hình ONEDRIVE_TOKEN_CACHE")

    cache = msal.SerializableTokenCache()
    cache.deserialize(ONEDRIVE_TOKEN_CACHE)

    app = msal.PublicClientApplication(GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if not accounts:
        raise Exception("Token cache không có tài khoản nào, cần chạy lại setup_onedrive_token.py")

    result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise Exception(f"Không lấy được access token, cần chạy lại setup_onedrive_token.py")

    return result["access_token"]

def download_excel_for_write():
    """Tải file Excel qua Graph API (để có thể ghi lại đúng file này)."""
    token = get_graph_access_token()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{ONEDRIVE_FILE_PATH}:/content"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
    resp.raise_for_status()
    return resp.content

def upload_excel(content_bytes):
    """Ghi đè nội dung file Excel lên OneDrive qua Graph API."""
    token = get_graph_access_token()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{ONEDRIVE_FILE_PATH}:/content"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
    resp = requests.put(url, headers=headers, data=content_bytes, timeout=60) # Tăng timeout lên 60s
    resp.raise_for_status()
    return resp.json()

def append_transaction_and_upload(amount, is_income):
    """
    Thêm 1 giao dịch mới vào cột Thu (A) hoặc Chi (B), dòng 2-25,
    Dùng công thức Excel để tự cập nhật tổng, sau đó upload lại.
    """
    content = download_excel_for_write()
    workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=False) # data_only=False để giữ công thức
    sheet = workbook.active

    col = 1 if is_income else 2  # A=1 (Thu), B=2 (Chi)

    # Tìm dòng trống đầu tiên trong khoảng 2-25
    target_row = None
    for i in range(2, 26):
        if sheet.cell(row=i, column=col).value is None:
            target_row = i
            break
    if target_row is None:
        raise Exception("Hết chỗ trống trong bảng (đã đủ 24 dòng), cần dọn bớt dữ liệu cũ trong Excel")

    # Ghi số tiền vào ô trống
    sheet.cell(row=target_row, column=col).value = amount

    # Cập nhật lại công thức Excel (để chắc chắn nó tính đúng nếu bị xóa công thức)
    sheet['A29'].value = '=SUM(A2:A25)'
    sheet['B29'].value = '=SUM(B2:B25)'
    sheet['A31'].value = '=A29-B29'

    buf = io.BytesIO()
    workbook.save(buf)
    buf.seek(0)
    upload_excel(buf.read())

def send_telegram_notification(text):
    if not TOKEN or not TELEGRAM_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=15)

# ================== WEBHOOK SEPAY ==================
@app_web.route('/sepay-webhook', methods=['POST'])
def sepay_webhook():
    auth_header = request.headers.get("Authorization", "")
    if not SEPAY_API_KEY or SEPAY_API_KEY not in auth_header:
        return jsonify({"success": False, "message": "Unauthorized"}), 401

    data = request.get_json(force=True, silent=True) or {}
    
    # Trả về 200 ngay lập tức để SePay không retry
    # Sau đó xử lý ở thread riêng
    threading.Thread(target=process_transaction, args=(data,)).start()
    
    return jsonify({"success": True}), 200

def process_transaction(data):
    """Hàm chạy ngầm để ghi Excel, không làm nghẽn webhook."""
    try:
        noi_dung = data.get("content", "")
        so_tien = data.get("transferAmount", 0)
        loai_gd = data.get("transferType")  # "in" = tiền vào, "out" = tiền ra
        is_income = (loai_gd == "in")

        append_transaction_and_upload(so_tien, is_income)

        loai_text = f"💰 Nhận tiền ({loai_gd})" if is_income else f"💸 Chi tiền ({loai_gd})"
        send_telegram_notification(
            f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
            f"Nội dung: {noi_dung}\n"
            f"----------------------------------------\n"
            f"✅ Đã ghi vào Excel thành công!"
        )
    except Exception as e:
        # Gửi lỗi về Telegram để bạn biết và sửa
        try:
            send_telegram_notification(f"❌ LỖI GHI EXCEL: \n<code>{str(e)}</code>")
        except:
            pass
        print(f"❌ LỖI XỬ LÝ GIAO DỊCH SEPAY: {str(e)}")

# ================== HÀM TELEGRAM BOT ==================
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📱 Lấy mã QR", callback_data="get_qr")],
        [InlineKeyboardButton("💰 Kiểm tra tiền", callback_data="check_money")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("🤖 Xin chào!\nNhập số <b>1</b> hoặc bấm menu dưới đây để chọn chức năng:", reply_markup=reply_markup, parse_mode="HTML")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_menu(update, context)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip() if update.message and update.message.text else ""
    if text == "1":
        await show_menu(update, context)
    else:
        await update.message.reply_text("💡 Nếu Muốn Tìm Menu Ấn Số 1")

async def handle_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "get_qr":
        tz_vn = pytz.timezone('Asia/Ho_Chi_Minh')
        current_time = datetime.now(tz_vn).strftime("%d/%m/%Y %H:%M:%S")

        media_group = []
        info_text = f"📅 <b>NGÀY TẠO MÃ:</b> <code>{current_time}</code>\n"
        info_text += f"👤 <b>HỌ TÊN:</b> <code>HUYNH NGOC NGHIA</code>\n\n"
        info_text += "📌 <b>DANH SÁCH SỐ TÀI KHOẢN:</b>\n\n"

        for i, bank in enumerate(BANKS, 1):
            caption_single = f"{bank['name']}\nSTK: {bank['account_number']}\nTên: {bank['account_name']}"
            media_group.append(InputMediaPhoto(media=bank['qr_url'], caption=caption_single))
            info_text += f"{i}. <b>{bank['name']}</b>\n   - STK: <code>{bank['account_number']}</code>\n   - Tên: <code>{bank['account_name']}</code>\n\n"

        await query.message.reply_media_group(media=media_group)
        await query.message.reply_text(info_text, parse_mode="HTML")

    elif query.data == "check_money":
        await query.message.reply_text("⏳ Đang tải dữ liệu từ OneDrive...")
        report = get_excel_data()
        await query.message.reply_text(report, parse_mode="HTML")

# ================== KHAI BÁO WEBHOOK ==================
async def setup_webhook(application):
    """Tự động cài đặt Webhook khi bot khởi động trên Render."""
    port = int(os.environ.get("PORT", 10000))
    webhook_url = f"https://{os.environ.get('RENDER_EXTERNAL_HOSTNAME', 'your-app.onrender.com')}/webhook" 
    # Lưu ý: Nếu bạn set telegram bot chỉ để trả lời lệnh /start và /1 thì không cần webhook cho Telegram, dùng Polling cũng được.
    # Tuy nhiên, trên Render, Polling dễ gây Conflict nếu bạn deploy 2 lần. 
    # Nếu muốn chạy Webhook:
    # await application.bot.set_webhook(url=webhook_url)
    pass 

def main():
    if not TOKEN:
        print("LỖI: Chưa cài đặt BOT_TOKEN!")
        sys.exit(1)

    # Khởi động Flask trước (để webhook hoạt động)
    port = int(os.environ.get("PORT", 10000))
    threading.Thread(target=lambda: app_web.run(host='0.0.0.0', port=port), daemon=True).start()

    # Khởi động Telegram Bot
    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    application.add_handler(CallbackQueryHandler(handle_button))

    print("Bot đang chạy...")
    
    # Chạy Polling để lấy lệnh Telegram (Nếu bạn muốn chạy Webhook cho Telegram thì bỏ dòng này và dùng application.run_webhook)
    # Lưu ý: Render sẽ hỗ trợ chạy polling. Nếu trước đó bạn từng chạy webhook, hãy vào link deleteWebhook để xóa.
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()