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

TOKEN = os.getenv("BOT_TOKEN")

# Link OneDrive dạng tải về trực tiếp (dùng để ĐỌC, giữ nguyên như cũ)
ONEDRIVE_URL = "https://1drv.ms/x/c/813BCA548F1AB473/IQDYUEgvvFlYRqhwhjmw-EFIASURiHvdRqvzgy28bqT6g0s?download=1"

# ====== CẤU HÌNH MỚI - dùng để GHI vào OneDrive qua Graph API ======
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")          # Application (client) ID từ Azure
ONEDRIVE_TOKEN_CACHE = os.getenv("ONEDRIVE_TOKEN_CACHE", "")  # chuỗi lấy từ setup_onedrive_token.py
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "")      # đường dẫn file trong OneDrive, vd: "SoThuChi/SoThuChi.xlsx"

SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")  # chuỗi bí mật bạn tự đặt, khớp với cấu hình webhook trên SePay
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")  # id chat/group nhận thông báo giao dịch

GRAPH_AUTHORITY = "https://login.microsoftonline.com/consumers"
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"]
# ====================================================================

app_web = Flask(__name__)


@app_web.route('/')
def home():
    return "Bot Telegram đang chạy!"


def run_flask():
    port = int(os.environ.get("PORT", 10000))
    app_web.run(host='0.0.0.0', port=port)


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


def get_excel_data():
    try:
        response = requests.get(ONEDRIVE_URL, timeout=10)
        if response.status_code != 200:
            return f"❌ Lỗi tải file Excel từ OneDrive (Mã lỗi: {response.status_code})"

        excel_file = io.BytesIO(response.content)
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = workbook.active

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


# ================== PHẦN MỚI: GHI VÀO ONEDRIVE QUA GRAPH API ==================

def get_graph_access_token():
    """Lấy access token bằng token cache đã xin quyền từ trước (setup_onedrive_token.py)."""
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
        raise Exception(f"Không lấy được access token, cần chạy lại setup_onedrive_token.py: {result}")

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
    resp = requests.put(url, headers=headers, data=content_bytes, timeout=30)
    resp.raise_for_status()
    return resp.json()


def append_transaction_and_upload(amount, is_income):
    """
    Thêm 1 giao dịch mới vào cột Thu (A) hoặc Chi (B), dòng 2-25,
    rồi cập nhật lại tổng (A29/B29) và số dư còn lại (A31), sau đó upload lại.
    """
    content = download_excel_for_write()
    workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
    sheet = workbook.active

    col = 1 if is_income else 2  # A=1 (Thu), B=2 (Chi)

    # tìm dòng trống đầu tiên trong khoảng 2-25
    target_row = None
    for i in range(2, 26):
        if sheet.cell(row=i, column=col).value is None:
            target_row = i
            break
    if target_row is None:
        raise Exception("Hết chỗ trống trong bảng (đã đủ 24 dòng), cần dọn bớt dữ liệu cũ trong Excel")

    sheet.cell(row=target_row, column=col).value = amount

    # tính lại tổng thu / chi trực tiếp bằng Python (không dùng công thức Excel)
    thu_total = sum(sheet.cell(row=i, column=1).value or 0 for i in range(2, 26) if isinstance(sheet.cell(row=i, column=1).value, (int, float)))
    chi_total = sum(sheet.cell(row=i, column=2).value or 0 for i in range(2, 26) if isinstance(sheet.cell(row=i, column=2).value, (int, float)))

    sheet['A29'].value = thu_total
    sheet['B29'].value = chi_total
    sheet['A31'].value = thu_total - chi_total

    buf = io.BytesIO()
    workbook.save(buf)
    buf.seek(0)
    upload_excel(buf.read())

    return thu_total, chi_total, thu_total - chi_total


def send_telegram_notification(text):
    if not TOKEN or not TELEGRAM_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=15)


@app_web.route('/sepay-webhook', methods=['POST'])
def sepay_webhook():
    auth_header = request.headers.get("Authorization", "")
    if not SEPAY_API_KEY or SEPAY_API_KEY not in auth_header:
        return jsonify({"success": False, "message": "Unauthorized"}), 401

    data = request.get_json(force=True, silent=True) or {}

    noi_dung = data.get("content", "")
    so_tien = data.get("transferAmount", 0)
    loai_gd = data.get("transferType")  # "in" = tiền vào, "out" = tiền ra
    is_income = (loai_gd == "in")

    try:
        thu_total, chi_total, remain = append_transaction_and_upload(so_tien, is_income)
    except Exception as e:
        # trả lỗi để SePay tự động retry
        return jsonify({"success": False, "message": str(e)}), 500

    loai_text = f"💰 Nhận tiền (transferType={loai_gd})" if is_income else f"💸 Chi tiền (transferType={loai_gd})"
    send_telegram_notification(
        f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
        f"Nội dung: {noi_dung}\n"
        f"----------------------------------------\n"
        f"🟢 Tổng thu: <code>{thu_total:,.0f}</code> VNĐ\n"
        f"🔴 Tổng chi: <code>{chi_total:,.0f}</code> VNĐ\n"
        f"💰 Còn lại: <code>{remain:,.0f}</code> VNĐ"
    )

    return jsonify({"success": True}), 200

# ================================================================================


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


def main():
    if not TOKEN:
        print("LỖI: Chưa cài đặt BOT_TOKEN!")
        sys.exit(1)

    threading.Thread(target=run_flask, daemon=True).start()

    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    application.add_handler(CallbackQueryHandler(handle_button))

    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
