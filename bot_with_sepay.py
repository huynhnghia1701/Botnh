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
from PIL import Image
import pytesseract
from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputMediaPhoto
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes

# ================== CẤU HÌNH ==================
TOKEN = os.getenv("BOT_TOKEN")
BOT_PASSWORD = "123123"

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
excel_write_lock = threading.Lock()  # Đảm bảo chỉ 1 thread ghi Excel tại 1 thời điểm

BANKS = [
    {"name": "Vietcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "96886693059121", "qr_url": "https://api.vietqr.io/image/MSB-96886693059121-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "OCB (Phương Đông)", "account_name": "HUYNH NGOC NGHIA", "account_number": "OCB-SEPHN48935", "qr_url": "https://api.vietqr.io/image/OCB-SEPHN48935-compact2.png?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "LPBank", "account_name": "HUYNH NGOC NGHIA", "account_number": "0916856322", "qr_url": "https://api.vietqr.io/image/970449-0916856322-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"},
    {"name": "Techcombank", "account_name": "HUYNH NGOC NGHIA", "account_number": "3838396852", "qr_url": "https://api.vietqr.io/image/970407-3838396852-compact2.jpg?accountName=HUYNH%20NGOC%20NGHIA"}
]

# ================== HÀM ĐỌC EXCEL (TỐI ƯU RAM) ==================
def get_excel_data():
    excel_file = None
    workbook = None
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
        
        return msg
    except requests.exceptions.Timeout:
        return "❌ Lỗi: Tải file Excel quá chậm (Timeout)."
    except Exception as e:
        return f"❌ Lỗi xử lý file Excel: {str(e)}"
    finally:
        # Giải phóng bộ nhớ RAM triệt để
        if workbook: workbook.close()
        if excel_file: excel_file.close()
        del content, excel_file, workbook
        gc.collect()

# ================== HÀM GHI EXCEL (GRAPH API) ==================
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

def upload_excel(content_bytes):
    token = get_graph_access_token()
    file_path = get_encoded_file_path()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/content"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
    try:
        resp = requests.put(url, headers=headers, data=content_bytes, timeout=30)
        if resp.status_code not in (200, 201):
            print(f"❌ Lỗi Graph API chi tiết: {resp.text}")
            resp.raise_for_status()
        return resp.json()
    except Exception as e:
        raise Exception(f"Lỗi upload: {str(e)}")

def append_transaction_and_upload(amount, is_income):
    with excel_write_lock:
        workbook = None
        buf = None
        try:
            content = download_excel_for_write()
            workbook = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
            sheet = workbook.active
            col = 1 if is_income else 2
            
            target_row = None
            for i in range(2, 28):
                if sheet.cell(row=i, column=col).value is None:
                    target_row = i
                    break
            if target_row is None: raise Exception("Hết chỗ trống trong bảng (đủ 26 dòng)! Cần dọn bớt Excel.")

            sheet.cell(row=target_row, column=col).value = amount

            buf = io.BytesIO()
            workbook.save(buf)
            buf.seek(0)
            upload_excel(buf.read())

            return True, target_row
        finally:
            if workbook: workbook.close()
            if buf: buf.close()
            del content, buf, workbook
            gc.collect()

# ================== HÀM XỬ LÝ OCR (CỰC NHẸ RAM) ==================
def process_receipt_image(image_bytes):
    img = None
    try:
        # Sử dụng PIL + Pytesseract siêu nhẹ
        img = Image.open(io.BytesIO(image_bytes))
        full_text = pytesseract.image_to_string(img)

        # Tìm các chuỗi số tiền
        found_numbers = re.findall(r'\b\d{1,3}(?:[.,]\d{3})+\b|\b\d{4,9}\b', full_text)

        if not found_numbers:
            return False, "❌ Không nhận diện được chuỗi số tiền hợp lệ trên ảnh bill."

        parsed_amounts = []
        for num in found_numbers:
            clean_num = int(re.sub(r'[.,]', '', num))
            if 1000 <= clean_num <= 500000000:
                parsed_amounts.append(clean_num)

        if not parsed_amounts:
            return False, "❌ Các giá trị số đọc được không phải số tiền hợp lệ."

        extracted_amount = max(parsed_amounts)

        # Ghi khoản CHI vào OneDrive
        _, row_idx = append_transaction_and_upload(extracted_amount, is_income=False)
        return True, f"💸 **ĐÃ THÊM KHOẢN CHI THÀNH CÔNG!**\n----------------------------------------\n🔹 Số tiền: <code>{extracted_amount:,.0f}</code> VNĐ\n📍 Vị trí: Cột CHI (Dòng {row_idx})\n☁️ Đã đồng bộ lên OneDrive."

    except Exception as e:
        return False, f"❌ Lỗi ghi file OneDrive: {str(e)}"
    finally:
        if img: img.close()
        del image_bytes, img
        gc.collect()

# ================== WEBHOOK SEPAY ==================
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

        loai_text = f"💰 Nhận tiền (in)" if is_income else f"💸 Chi tiền (out)"
        
        if ghi_thanh_cong:
            send_telegram_notification(
                f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
                f"Nội dung: {noi_dung}\n"
                f"----------------------------------------\n"
                f"✅ Đã ghi vào Excel thành công!"
            )
        else:
            send_telegram_notification(
                f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
                f"Nội dung: {noi_dung}\n"
                f"❌ Lỗi ghi file:\n<code>{loi_chi_tiet}</code>"
            )
    except Exception as e:
        print(f"❌ LỖI XỬ LÝ GIAO DỊCH SEPAY: {str(e)}")
    finally:
        gc.collect()

def send_telegram_notification(text):
    if not TOKEN or not TELEGRAM_CHAT_ID: return
    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/sendMessage", json={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=15)
    except: pass

# ================== HÀM TELEGRAM BOT ==================
async def show_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[InlineKeyboardButton("📱 Lấy mã QR", callback_data="get_qr")], [InlineKeyboardButton("💰 Kiểm tra tiền", callback_data="check_money")]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("🤖 Xin chào!\nNhập số <b>1</b> hoặc bấm menu dưới đây để chọn chức năng.\nNhập số <b>2</b> để đăng xuất.\n📷 <i>Gửi ảnh bill chuyển khoản để tự động ghi khoản Chi!</i>", reply_markup=reply_markup, parse_mode="HTML")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if logged_in_users.get(user_id): await show_menu(update, context)
    else: await update.message.reply_text("🔐 <b>Menu được bảo vệ.</b>\nVui lòng nhập mật khẩu để tiếp tục:", parse_mode="HTML")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not logged_in_users.get(user_id):
        await update.message.reply_text("🔒 Vui lòng đăng nhập trước khi gửi ảnh bill!")
        return

    msg = await update.message.reply_text("🔍 Bot đang quét dữ liệu số tiền từ bill và cập nhật lên OneDrive...")
    
    try:
        photo_file = await update.message.photo[-1].get_file()
        image_bytes = await photo_file.download_as_bytearray()

        success, response_msg = await asyncio.to_thread(process_receipt_image, bytes(image_bytes))
        await msg.edit_text(response_msg, parse_mode="HTML")
    except Exception as e:
        await msg.edit_text(f"❌ Có lỗi trong quá trình xử lý ảnh: {str(e)}")
    finally:
        gc.collect()

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
    else: await update.message.reply_text("💡 Nếu Muốn Tìm Menu Ấn Số 1 hoặc gửi ảnh bill để ghi khoản Chi.")

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
    except: pass

    application = Application.builder().token(TOKEN).concurrent_updates(True).build()
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    application.add_handler(CallbackQueryHandler(handle_button))

    print("Bot đang chạy...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
