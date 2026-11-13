import os
import sys
import io
import gc
import re
import json
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
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")            # BẮT BUỘC
GRAPH_AUTHORITY = os.getenv("GRAPH_AUTHORITY", "https://login.microsoftonline.com/consumers")
GRAPH_SCOPES = ["Files.ReadWrite", "User.Read", "offline_access"]
GRAPH_REFRESH_TOKEN = os.getenv("GRAPH_REFRESH_TOKEN", "")    # token cache (serialized) HOẶC refresh_token thô
ONEDRIVE_FILE_PATH = os.getenv("ONEDRIVE_FILE_PATH", "")     # ví dụ: "Documents/SoThuChi.xlsx"

# --- SePay / Telegram ---
SEPAY_API_KEY = os.getenv("SEPAY_API_KEY", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")

# --- Excel ---
EXCEL_DATA_START_ROW = int(os.getenv("EXCEL_DATA_START_ROW", "2"))
EXCEL_DATA_END_ROW = int(os.getenv("EXCEL_DATA_END_ROW", "27"))  # inclusive

app_web = Flask(__name__)
logged_in_users = {}
awaiting_expense_password = set()
expense_mode_users = set()
excel_write_lock = threading.Lock()

# ================== MICROSOFT GRAPH AUTH ==================
def _validate_graph_config():
    if not GRAPH_CLIENT_ID:
        raise Exception("GRAPH_CLIENT_ID đang trống! Cần set env trên Render.")
    if not GRAPH_REFRESH_TOKEN:
        raise Exception("GRAPH_REFRESH_TOKEN đang trống! Cần set env trên Render.")
    if not ONEDRIVE_FILE_PATH:
        raise Exception("ONEDRIVE_FILE_PATH đang trống! Cần set env trên Render.")


def get_graph_access_token():
    """
    Lấy access token theo thứ tự ưu tiên:
      1) Thử acquire_token_silent từ SerializableTokenCache (nếu env là token cache serialized)
      2) Nếu thất bại, dùng refresh_token thô để lấy access token mới
    """
    _validate_graph_config()

    # --- Cách 1: env là token cache serialized (khuyến nghị) ---
    try:
        cache = msal.SerializableTokenCache()
        cache.deserialize(GRAPH_REFRESH_TOKEN)
        app = msal.PublicClientApplication(
            GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY, token_cache=cache
        )
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]
    except Exception as e:
        print(f"⚠️ acquire_token_silent thất bại: {e}")

    # --- Cách 2: env là refresh_token thô ---
    try:
        app = msal.PublicClientApplication(GRAPH_CLIENT_ID, authority=GRAPH_AUTHORITY)
        result = app.acquire_token_by_refresh_token(
            GRAPH_REFRESH_TOKEN, scopes=GRAPH_SCOPES
        )
        if result and "access_token" in result:
            new_refresh = result.get("refresh_token")
            if new_refresh:
                print(
                    "🔑 Refresh token mới (copy vào env GRAPH_REFRESH_TOKEN nếu muốn):",
                    new_refresh[:20] + "..."
                )
            return result["access_token"]
    except Exception as e:
        raise Exception(f"Không lấy được access token: {e}")

    raise Exception(
        "Không lấy được access token. Kiểm tra GRAPH_CLIENT_ID / GRAPH_REFRESH_TOKEN. "
        "(refresh_token có thể đã hết hạn hoặc bị thu hồi — cần login lại bằng script local.)"
    )


def _encoded_file_path():
    """Encode path OneDrive để đưa vào URL Graph API."""
    path = ONEDRIVE_FILE_PATH.strip().strip("/")
    if not path:
        raise Exception("ONEDRIVE_FILE_PATH đang trống!")
    if "/" in path:
        folder, filename = path.rsplit("/", 1)
        return f"{quote(folder)}/{quote(filename)}"
    return quote(path)


# ================== ĐỌC EXCEL (API) ==================
def download_excel_for_write():
    """Tải file Excel từ OneDrive qua Graph API -> trả về bytes."""
    token = get_graph_access_token()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{_encoded_file_path()}:/content"
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    if resp.status_code != 200:
        raise Exception(
            f"Graph API tải file lỗi {resp.status_code}: {resp.text[:300]}"
        )
    return resp.content


def get_excel_data():
    """
    Đọc dữ liệu Excel từ OneDrive qua Graph API, trả về text báo cáo cho Telegram.
    """
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


# ================== GHI EXCEL (API) ==================
def upload_excel(content_bytes, max_retries=4):
    """Upload file bytes lên OneDrive qua Graph API, retry khi bị 423 Locked."""
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
        f"{last_error} Đã thử {max_retries} lần. "
        f"Đóng hẳn file Excel trên mọi thiết bị rồi gửi lại."
    )

def append_transaction_and_upload(amount, is_income):
    """Ghi 1 giao dịch vào cột thu/chi ở dòng trống đầu tiên, rồi upload."""
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
            raise Exception(
                f"Hết chỗ trống trong bảng (dòng {EXCEL_DATA_START_ROW}-{EXCEL_DATA_END_ROW})."
            )

        sheet.cell(row=target_row, column=col).value = amount

        buf = io.BytesIO()
        workbook.save(buf)
        buf.seek(0)
        upload_excel(buf.read())

        del workbook, sheet, content, buf
        gc.collect()
        # ================== API TEST (dùng để debug trên Render) ==================
API_TEST_TOKEN = os.getenv("API_TEST_TOKEN", "")  # đặt env để bảo vệ endpoint

def _check_api_auth():
    if not API_TEST_TOKEN:
        return True  # nếu chưa đặt thì bỏ qua (chỉ dùng khi test local)
    return request.headers.get("X-Api-Token") == API_TEST_TOKEN


@app_web.route("/ping", methods=["GET"])
def ping():
    return "OK", 200


@app_web.route("/api/excel/read", methods=["GET"])
def api_excel_read():
    """GET: đọc toàn bộ dữ liệu thu/chi trong Excel."""
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
    """
    POST body JSON:
      { "amount": 2550000, "type": "expense" }   # hoặc "income"
    """
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


@app_web.route("/api/graph/token", methods=["GET"])
def api_graph_token():
    """GET: kiểm tra Graph API có lấy được token không (không trả token thật)."""
    if not _check_api_auth():
        return jsonify({"success": False, "message": "Unauthorized"}), 401
    try:
        token = get_graph_access_token()
        return jsonify({"success": True, "token_prefix": token[:20] + "..."})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
        def main():
    if not TOKEN:
        print("LỖI: Chưa cài BOT_TOKEN!")
        sys.exit(1)

    try:
        _validate_graph_config()
        print("✅ Env Graph API đầy đủ.")
    except Exception as e:
        print(f"❌ {e}")

    port = int(os.environ.get("PORT", 10000))
    from waitress import serve
    threading.Thread(
        target=lambda: serve(app_web, host="0.0.0.0", port=port),
        daemon=True,
    ).start()

    try:
        requests.post(f"https://api.telegram.org/bot{TOKEN}/deleteWebhook")
    except Exception:
        pass

    application = (
        Application.builder().token(TOKEN).concurrent_updates(True).build()
    )
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("xoahet", clearall_command))
    application.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))
    application.add_handler(CallbackQueryHandler(handle_button))
    application.add_error_handler(error_handler)

    print("Bot đang chạy...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)