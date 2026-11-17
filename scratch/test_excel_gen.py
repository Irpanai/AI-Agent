import os
import sys
import io
from datetime import datetime
from dotenv import load_dotenv

sys.path.insert(0, r"D:\AI Agent")
load_dotenv()

from app import sheets_manager
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def generate_test_excel():
    target_date = datetime.now().strftime("%Y-%m-%d")
    tgl_display = datetime.now().strftime("%d %B %Y")
    
    ws_mutasi_sheet = sheets_manager.get_worksheet("Riwayat_Mutasi")
    mutasi_records = ws_mutasi_sheet.get_all_values()
    
    ws_stok_sheet = sheets_manager.get_worksheet("Stok_Barang")
    stok_records = ws_stok_sheet.get_all_values()
    
    ws_master_sheet = sheets_manager.get_worksheet("Master_Barang")
    master_records = ws_master_sheet.get_all_values()
    nama_map = {r[0].strip().upper(): r[1].strip() for r in master_records[1:] if len(r) >= 2}

    wb = openpyxl.Workbook()
    
    # ------------------ SHEET 1: MUTASI ------------------
    ws1 = wb.active
    ws1.title = "Mutasi Gudang"
    ws1.views.sheetView[0].showGridLines = True
    
    # Title
    ws1.merge_cells("A1:J1")
    ws1["A1"] = f"LAPORAN OPERASIONAL & MUTASI GUDANG ({tgl_display})"
    ws1["A1"].font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    ws1["A1"].fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 40
    
    # Header Table
    headers = ["No", "Timestamp", "Tanggal", "Tipe", "Kode SKU", "Nama Barang", "Qty (pcs)", "Sisa Stok", "Pengirim", "Keterangan"]
    ws1.row_dimensions[3].height = 26
    h_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    h_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )
    
    for c_idx, h in enumerate(headers, 1):
        cell = ws1.cell(row=3, column=c_idx, value=h)
        cell.font = h_font
        cell.fill = h_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_pos = 4
    total_keluar = 0
    total_masuk = 0
    
    for r_idx, r in enumerate(mutasi_records[1:], 1):
        # r = [Timestamp, Tanggal, Tipe, Kode_SKU, Nama_Barang, Qty, Sisa_Stok, Pengirim, Keterangan]
        tipe = r[2] if len(r) > 2 else ""
        try:
            qty = int(r[5]) if len(r) > 5 else 0
        except ValueError:
            qty = 0
        if tipe.upper() == "KELUAR":
            total_keluar += qty
        elif tipe.upper() == "MASUK":
            total_masuk += qty
            
        row_vals = [
            r_idx,
            r[0] if len(r) > 0 else "",
            r[1] if len(r) > 1 else "",
            tipe,
            r[3] if len(r) > 3 else "",
            r[4] if len(r) > 4 else "",
            qty,
            r[6] if len(r) > 6 else "",
            r[7] if len(r) > 7 else "",
            r[8] if len(r) > 8 else "",
        ]
        
        row_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid") if r_idx % 2 == 0 else None
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws1.cell(row=row_pos, column=c_idx, value=val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill
            cell.alignment = Alignment(horizontal="center" if c_idx in (1, 2, 3, 4, 5, 7, 8) else "left", vertical="center")
        row_pos += 1

    # Auto fit column widths
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 11)

    # ------------------ SHEET 2: STOK BARANG ------------------
    ws2 = wb.create_sheet(title="Stok Gudang Terkini")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2.merge_cells("A1:E1")
    ws2["A1"] = f"MASTER SISA STOK GUDANG TERKINI ({tgl_display})"
    ws2["A1"].font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 40
    
    stok_headers = ["No", "Kode SKU", "Nama Barang Standar", "Sisa Stok (pcs)", "Status"]
    ws2.row_dimensions[3].height = 26
    for c_idx, h in enumerate(stok_headers, 1):
        cell = ws2.cell(row=3, column=c_idx, value=h)
        cell.font = h_font
        cell.fill = h_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    danger_fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
    warning_fill = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")
    ok_fill = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")
    
    r2_pos = 4
    for idx, r in enumerate(stok_records[1:], 1):
        sku = r[0].strip().upper() if len(r) > 0 else ""
        try:
            sisa = int(r[1]) if len(r) > 1 and r[1].strip() else 0
        except ValueError:
            sisa = 0
        nama = nama_map.get(sku, "-")
        
        if sisa == 0:
            status = "HABIS"
            status_fill = danger_fill
        elif sisa < 10:
            status = "KRITIS (<10)"
            status_fill = warning_fill
        else:
            status = "AMAN"
            status_fill = ok_fill
            
        r_vals = [idx, sku, nama, sisa, status]
        for c_idx, val in enumerate(r_vals, 1):
            cell = ws2.cell(row=r2_pos, column=c_idx, value=val)
            cell.font = Font(name="Segoe UI", size=10)
            cell.border = thin_border
            if c_idx == 5:
                cell.fill = status_fill
                cell.font = Font(name="Segoe UI", size=10, bold=True)
            cell.alignment = Alignment(horizontal="center" if c_idx in (1, 2, 4, 5) else "left", vertical="center")
        r2_pos += 1

    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    data = buf.getvalue()
    filename = f"Laporan_Gudang_{datetime.now().strftime('%d%b%Y')}.xlsx"
    print(f"Generated {filename} ({len(data)} bytes), Total Keluar: {total_keluar}, Total Masuk: {total_masuk}")
    return data, filename

data, filename = generate_test_excel()
