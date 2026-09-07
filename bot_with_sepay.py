def process_transaction(data):
    try:
        noi_dung = data.get("content", "")
        so_tien = data.get("transferAmount", 0)
        loai_gd = data.get("transferType")
        is_income = (loai_gd == "in")
        append_transaction_and_upload(so_tien, is_income)
        
        content = download_excel_for_write()
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        sh = wb.active
        
        thu_total = sum([sh.cell(row=i, column=1).value or 0 for i in range(2, 26) if isinstance(sh.cell(row=i, column=1).value, (int, float))])
        chi_total = sum([sh.cell(row=i, column=2).value or 0 for i in range(2, 26) if isinstance(sh.cell(row=i, column=2).value, (int, float))])
        remain = thu_total - chi_total

        loai_text = f"💰 Nhận tiền (in)" if is_income else f"💸 Chi tiền (out)"
        
        send_telegram_notification(
            f"{loai_text}: <code>{so_tien:,.0f}</code> VNĐ\n"
            f"Nội dung: {noi_dung}\n"
            f"----------------------------------------\n"
            #f"🟢 Tổng thu: <code>{thu_total:,.0f}</code> VNĐ\n"
            #f"🔴 Tổng chi: <code>{chi_total:,.0f}</code> VNĐ\n"
            #f"💰 Còn lại: <code>{remain:,.0f}</code> VNĐ\n"
            f"✅ Đã ghi vào Excel thành công!"
        )
        
        del wb, sh, content
        gc.collect()
        
    except Exception as e:
        try:
            send_telegram_notification(f"❌ LỖI GHI EXCEL: \n<code>{str(e)}</code>")
        except:
            pass
        print(f"❌ LỖI XỬ LÝ GIAO DỊCH SEPAY: {str(e)}")