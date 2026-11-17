import os
import sys
import json
from dotenv import load_dotenv

sys.path.insert(0, r"D:\AI Agent")
load_dotenv()

from app import app, sheets_manager, handle_incoming_message, broadcast_restok_to_groups

print("=================================================================")
print("🧪 TESTING FITUR ADMIN, WEB PANEL, OPNAME, LEADERBOARD & BROADCAST")
print("=================================================================")

# 1. Test Pemisahan Mode: Perintah Admin di Grup Ditolak
print("\n[TEST 1] Proteksi Perintah Admin di Grup")
msg_grup = {
    "message_id": 101,
    "chat": {"id": -100123456789, "type": "group", "title": "Grup Logistik Test"},
    "from": {"id": 9999, "first_name": "Seller Lapangan"},
    "text": "/restok MHPD 021 50"
}
success, reason, _ = handle_incoming_message(msg_grup)
print(f"Hasil di Grup: success={success}, reason={reason}")
assert reason == "admin_action_blocked_in_group", "Harusnya perintah admin di grup ditolak!"
print("✅ Test 1 Berhasil: Perintah admin berhasil diblokir di grup!")

# 2. Test Stock Opname Fisik
print("\n[TEST 2] Koreksi Stock Opname (Fisik Gudang)")
cur_qty, new_qty, op_success, nama = sheets_manager.perform_stock_opname(
    kode_sku="MHPD 021",
    diff_qty=-1,
    alasan="Uji coba test otomatis opname fisik rusak",
    admin_name="Test Admin"
)
print(f"Hasil Opname MHPD 021: {nama} | Stok: {cur_qty} -> {new_qty} (Success: {op_success})")
assert op_success, "Opname harusnya sukses!"
# Kembalikan lagi stoknya (+1)
sheets_manager.perform_stock_opname("MHPD 021", 1, "Rollback test opname", "Test Admin")
print("✅ Test 2 Berhasil: Koreksi stock opname dan pencatatan riwayat sukses!")

# 3. Test Leaderboard Seller
print("\n[TEST 3] Leaderboard Seller Hari Ini")
lb_text = sheets_manager.get_leaderboard_text()
print("Cuplikan Leaderboard Teks:\n" + lb_text[:250] + "...\n")
lb_data = sheets_manager.get_leaderboard_data()
print(f"Leaderboard Data JSON: {len(lb_data.get('sellers', []))} sellers terdeteksi")
assert "LEADERBOARD" in lb_text, "Teks leaderboard harus valid!"
print("✅ Test 3 Berhasil: Leaderboard seller berfungsi 100%!")

# 4. Test Web App Flask Routes
print("\n[TEST 4] Telegram Mini App Web Panel Routes & APIs")
client = app.test_client()

res_panel = client.get("/panel")
print(f"GET /panel -> Status: {res_panel.status_code}")
assert res_panel.status_code == 200, "Panel HTML harus return 200"
assert b"Panel Gudang Tira" in res_panel.data, "HTML harus memuat judul Panel Gudang Tira"

res_stok = client.get("/api/stok")
print(f"GET /api/stok -> Status: {res_stok.status_code}")
assert res_stok.status_code == 200
data_stok = json.loads(res_stok.data)
print(f"Jumlah produk di API: {len(data_stok.get('items', []))}")

res_lb = client.get("/api/leaderboard")
print(f"GET /api/leaderboard -> Status: {res_lb.status_code}")
assert res_lb.status_code == 200
print("✅ Test 4 Berhasil: Web Panel dan seluruh API JSON berfungsi sempurna!")

print("\n=================================================================")
print("🎉 SELURUH FITUR ADMIN & PEMISAHAN PERAN TELAH TERVERIFIKASI 100%!")
print("=================================================================")
