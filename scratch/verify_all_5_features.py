import os
import sys
import io
import openpyxl
from datetime import datetime
from dotenv import load_dotenv

sys.path.insert(0, r"D:\AI Agent")
load_dotenv()

from app import sheets_manager, gemini_normalizer, extract_qty_from_text, extract_no_resi_from_text

print("=========================================================")
print("🧪 TESTING 5 FITUR BARU SISTEM AI GUDANG")
print("=========================================================")

# -----------------------------------------------------------------------------
# TEST 1: ANTI-DUPLICATE RESI
# -----------------------------------------------------------------------------
print("\n[TEST 1] Anti-Duplicate Resi & Regex Detection")
test_text_1 = "Tolong kirim orderan No. Resi: 11004683469987 ya segera"
resi_extracted = extract_no_resi_from_text(test_text_1)
print(f"1a. Ekstraksi Resi dari Teks: '{resi_extracted}' (Expected: '11004683469987')")
assert resi_extracted == "11004683469987", "Gagal ekstrak no resi!"

# Cek fungsi check_duplicate_resi
print("1b. Pengecekan Duplicate Resi ke Tab Riwayat_Mutasi...")
dup = sheets_manager.check_duplicate_resi("11004683469987")
print(f"Hasil check_duplicate_resi: {dup}")
print("✅ Test 1 Berhasil!")

# -----------------------------------------------------------------------------
# TEST 2: KONVERSI SATUAN KEMASAN (LUSIN, KODI, BOX, DUS -> PCS)
# -----------------------------------------------------------------------------
print("\n[TEST 2] Konversi Satuan Kemasan (Box, Lusin, Dus -> Pcs)")
qty_lusin = extract_qty_from_text("kirim 2 lusin bio binahong")
print(f"2a. 2 lusin Bio Binahong: {qty_lusin} pcs (Expected: 24)")
assert qty_lusin == 24, "Gagal konversi lusin!"

qty_box = extract_qty_from_text("ambil 2 box g-max")
print(f"2b. 2 box G-MAX: {qty_box} pcs (Expected: 20)")
assert qty_box == 20, "Gagal konversi box!"

qty_dus = extract_qty_from_text("1 dus royal channa")
print(f"2c. 1 dus Royal Channa: {qty_dus} pcs (Expected: 24)")
assert qty_dus == 24, "Gagal konversi dus!"

qty_kodi = extract_qty_from_text("3 kodi produk herbal")
print(f"2d. 3 kodi produk herbal: {qty_kodi} pcs (Expected: 60)")
assert qty_kodi == 60, "Gagal konversi kodi!"
print("✅ Test 2 Berhasil!")

# -----------------------------------------------------------------------------
# TEST 3: MULTI-ITEM DALAM 1 RESI VIA GEMINI AI
# -----------------------------------------------------------------------------
print("\n[TEST 3] Multi-Item dalam 1 Resi (Orderan Campuran / Bundle)")
daftar_master = sheets_manager.get_master_barang_reference()
multi_text = "No. Resi: 11004683469987. Pesanan 1 Dephero dan 2 Habbafit."
ai_res = gemini_normalizer.normalize(multi_text, daftar_master)
print(f"3a. AI Result Resi: {ai_res.get('no_resi')}")
print(f"3b. Jumlah Item Terdeteksi: {len(ai_res.get('items', []))}")
for idx, it in enumerate(ai_res.get('items', []), 1):
    print(f"   Item {idx}: [{it.get('kode_sku')}] {it.get('nama_standar')} - Qty: {it.get('qty')} {it.get('satuan_asli')}")

assert len(ai_res.get('items', [])) >= 2, "Harusnya terdeteksi minimal 2 item!"
print("✅ Test 3 Berhasil!")

# -----------------------------------------------------------------------------
# TEST 4: EKSPOR LAPORAN EXCEL (.xlsx)
# -----------------------------------------------------------------------------
print("\n[TEST 4] Ekspor Laporan Excel (.xlsx) 2-Tab Berwarna")
excel_bytes, filename, total_keluar, total_masuk, total_baris = sheets_manager.export_excel_bytes()
print(f"4a. File: {filename}")
print(f"4b. Ukuran File: {len(excel_bytes)} bytes")
print(f"4c. Total Keluar: {total_keluar} pcs, Total Masuk: {total_masuk} pcs, Baris: {total_baris}")

# Verifikasi workbook openpyxl
wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
print(f"4d. Sheet names: {wb.sheetnames}")
assert "Mutasi Gudang" in wb.sheetnames, "Sheet Mutasi Gudang tidak ada!"
assert "Stok Gudang Terkini" in wb.sheetnames, "Sheet Stok Gudang Terkini tidak ada!"
print("✅ Test 4 Berhasil!")

# -----------------------------------------------------------------------------
# TEST 5: AUTO-BRIEFING PAGI & REKAP SORE
# -----------------------------------------------------------------------------
print("\n[TEST 5] Auto-Briefing Pagi & Rekap Sore")
briefing_text = sheets_manager.get_morning_briefing_text()
print("5a. Cuplikan Briefing Pagi:")
print(briefing_text[:180] + "...\n")

recap_text = sheets_manager.get_evening_recap_text()
print("5b. Cuplikan Rekap Sore:")
print(recap_text[:180] + "...\n")
assert "SELAMAT PAGI" in briefing_text, "Briefing pagi header salah!"
assert "TUTUP BUKU" in recap_text, "Rekap sore header salah!"
print("✅ Test 5 Berhasil!")

print("\n=========================================================")
print("🎉 SELURUH 5 FITUR BARU TELAH TERVERIFIKASI 100% SUKSES!")
print("=========================================================")
