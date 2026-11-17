"""
=============================================================================
Script Seeder: Mengisi Data Dummy ke Google Spreadsheet
=============================================================================
Tab yang diisi:
1. Master_Barang: 28 data SKU resmi & nama standar.
2. Stok_Barang  : 28 data SKU dengan kuantitas stok awal dummy.
3. Kamus_Alias  : Variasi nama pasar / marketplace / singkatan typo untuk cache.
=============================================================================
"""

import os
import sys
import random
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv()

SPREADSHEET_TARGET = os.getenv("SPREADSHEET_ID_OR_URL", "").strip()
CREDS_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json").strip()

if not SPREADSHEET_TARGET:
    print("❌ ERROR: SPREADSHEET_ID_OR_URL belum diisi di file .env!")
    print("Silakan buka file .env dan masukkan URL atau ID Google Spreadsheet Anda.")
    sys.exit(1)

if not os.path.exists(CREDS_FILE):
    print(f"❌ ERROR: File kredensial '{CREDS_FILE}' tidak ditemukan!")
    sys.exit(1)

# =============================================================================
# DATA MASTER BARANG (28 SKU)
# =============================================================================
MASTER_BARANG_DATA = [
    ["MHPD 001", "ROYAL CHANNA"],
    ["MHPD 002", "IN SLIM"],
    ["MHPD 003", "HABBAFIT"],
    ["MHPD 004", "BIO BINAHONG"],
    ["MHPD 005", "LACTACAPS"],
    ["MHPD 006", "VERBALITAS"],
    ["MHPD 007", "ENXIEHERB"],
    ["MHPD 008", "NONI MORINDA"],
    ["MHPD 009", "XT SPRAY 10 ML"],
    ["MHPD 010", "ARVOS"],
    ["MHPD 011", "COSMOCARE"],
    ["MHPD 012", "HIFZ MEN OIL"],
    ["MHPD 013", "BREAST G"],
    ["MHPD 014", "VERBA LUVE"],
    ["MHPD 015", "VERBAGEL RED"],
    ["MHPD 016", "VERBAGEL BLUE"],
    ["MHPD 017", "VIOLLA GEL"],
    ["MHPD 018", "VIOLLA SWEET CREAM"],
    ["MHPD 019", "VI-SOAP"],
    ["MHPD 020", "IMEN WASH"],
    ["MHPD 021", "DEPHERO"],
    ["MHPD 022", "G-MAX"],
    ["MHPD 023", "V-MAX"],
    ["MHPD 024", "ARJUNA"],
    ["MHPD 025", "GARAGA"],
    ["MHPD 026", "HAPPY MOMI"],
    ["MHPD 027", "G-MAX SACHET"],
    ["MHPD 028", "V-MAX SACHET"],
]

# =============================================================================
# DATA DUMMY STOK AWAL
# =============================================================================
# Menghasilkan stok realistis (antara 25 sampai 150 pcs per item)
DUMMY_STOCK_VALUES = [
    100, 75, 120, 85, 90, 60, 45, 50, 110, 40,
    65, 80, 95, 70, 85, 85, 60, 55, 130, 90,
    70, 140, 130, 85, 60, 75, 200, 200
]

STOK_BARANG_DATA = [
    [item[0], DUMMY_STOCK_VALUES[i]] for i, item in enumerate(MASTER_BARANG_DATA)
]

# =============================================================================
# DATA KAMUS ALIAS (Gaya Chat Seller Marketplace, Typo, dan Singkatan)
# =============================================================================
KAMUS_ALIAS_DATA = [
    # MHPD 001 - ROYAL CHANNA
    ["royal channa", "MHPD 001"],
    ["royal chana", "MHPD 001"],
    ["royl channa", "MHPD 001"],
    ["kapsul ikan gabus royal channa", "MHPD 001"],
    ["royal cana", "MHPD 001"],

    # MHPD 002 - IN SLIM
    ["in slim", "MHPD 002"],
    ["inslim", "MHPD 002"],
    ["inslim pelangsing", "MHPD 002"],
    ["in slim diet", "MHPD 002"],
    ["kapsul inslim", "MHPD 002"],

    # MHPD 003 - HABBAFIT
    ["habbafit", "MHPD 003"],
    ["habafit", "MHPD 003"],
    ["habbafit jinten hitam", "MHPD 003"],
    ["habasit", "MHPD 003"],

    # MHPD 004 - BIO BINAHONG
    ["bio binahong", "MHPD 004"],
    ["biobinahong", "MHPD 004"],
    ["kapsul bio binahong", "MHPD 004"],
    ["binahong kapsul", "MHPD 004"],

    # MHPD 005 - LACTACAPS
    ["lactacaps", "MHPD 005"],
    ["laktacaps", "MHPD 005"],
    ["lactacaps pelancar asi", "MHPD 005"],
    ["laktakap", "MHPD 005"],

    # MHPD 006 - VERBALITAS
    ["verbalitas", "MHPD 006"],
    ["verbalitas kapsul", "MHPD 006"],
    ["vitalitas verbalitas", "MHPD 006"],

    # MHPD 007 - ENXIEHERB
    ["enxieherb", "MHPD 007"],
    ["enxie herb", "MHPD 007"],
    ["enxi herb", "MHPD 007"],
    ["enxie herba", "MHPD 007"],

    # MHPD 008 - NONI MORINDA
    ["noni morinda", "MHPD 008"],
    ["morinda noni", "MHPD 008"],
    ["kapsul noni morinda", "MHPD 008"],
    ["ekstrak noni morinda", "MHPD 008"],

    # MHPD 009 - XT SPRAY 10 ML
    ["xt spray", "MHPD 009"],
    ["xt spray 10 ml", "MHPD 009"],
    ["xt spray 10ml", "MHPD 009"],
    ["spray xt 10 ml", "MHPD 009"],

    # MHPD 010 - ARVOS
    ["arvos", "MHPD 010"],
    ["kapsul arvos", "MHPD 010"],
    ["arvos herbal", "MHPD 010"],

    # MHPD 011 - COSMOCARE
    ["cosmocare", "MHPD 011"],
    ["cosmo care", "MHPD 011"],
    ["cream cosmocare", "MHPD 011"],

    # MHPD 012 - HIFZ MEN OIL
    ["hifz men oil", "MHPD 012"],
    ["hifz oil", "MHPD 012"],
    ["hifs men oil", "MHPD 012"],
    ["minyak hifz", "MHPD 012"],

    # MHPD 013 - BREAST G
    ["breast g", "MHPD 013"],
    ["breast-g", "MHPD 013"],
    ["krim breast g", "MHPD 013"],
    ["serum breast g", "MHPD 013"],

    # MHPD 014 - VERBA LUVE
    ["verba luve", "MHPD 014"],
    ["verbaluve", "MHPD 014"],
    ["verba love", "MHPD 014"],

    # MHPD 015 - VERBAGEL RED
    ["verbagel red", "MHPD 015"],
    ["verba gel red", "MHPD 015"],
    ["verbagel merah", "MHPD 015"],
    ["verba gel merah", "MHPD 015"],

    # MHPD 016 - VERBAGEL BLUE
    ["verbagel blue", "MHPD 016"],
    ["verba gel blue", "MHPD 016"],
    ["verbagel biru", "MHPD 016"],
    ["verba gel biru", "MHPD 016"],

    # MHPD 017 - VIOLLA GEL
    ["violla gel", "MHPD 017"],
    ["viola gel", "MHPD 017"],
    ["gel violla", "MHPD 017"],

    # MHPD 018 - VIOLLA SWEET CREAM
    ["violla sweet cream", "MHPD 018"],
    ["violla sweetcream", "MHPD 018"],
    ["viola sweet cream", "MHPD 018"],
    ["krim violla sweet", "MHPD 018"],

    # MHPD 019 - VI-SOAP
    ["vi-soap", "MHPD 019"],
    ["vi soap", "MHPD 019"],
    ["visoap", "MHPD 019"],
    ["sabun vi soap", "MHPD 019"],

    # MHPD 020 - IMEN WASH
    ["imen wash", "MHPD 020"],
    ["imenwash", "MHPD 020"],
    ["i-men wash", "MHPD 020"],
    ["sabun imen", "MHPD 020"],

    # MHPD 021 - DEPHERO
    ["dephero", "MHPD 021"],
    ["kapsul dephero", "MHPD 021"],
    ["depero", "MHPD 021"],

    # MHPD 022 - G-MAX
    ["g-max", "MHPD 022"],
    ["gmax", "MHPD 022"],
    ["g max box", "MHPD 022"],
    ["kopi g-max", "MHPD 022"],

    # MHPD 023 - V-MAX
    ["v-max", "MHPD 023"],
    ["vmax", "MHPD 023"],
    ["v max kapsul", "MHPD 023"],
    ["vmax herbal", "MHPD 023"],

    # MHPD 024 - ARJUNA
    ["arjuna", "MHPD 024"],
    ["kopi arjuna", "MHPD 024"],
    ["kapsul arjuna", "MHPD 024"],

    # MHPD 025 - GARAGA
    ["garaga", "MHPD 025"],
    ["kapsul garaga", "MHPD 025"],
    ["garaga herbal", "MHPD 025"],

    # MHPD 026 - HAPPY MOMI
    ["happy momi", "MHPD 026"],
    ["happy mommy", "MHPD 026"],
    ["happymomi", "MHPD 026"],
    ["kapsul happy momi", "MHPD 026"],

    # MHPD 027 - G-MAX SACHET
    ["g-max sachet", "MHPD 027"],
    ["gmax sachet", "MHPD 027"],
    ["g max saset", "MHPD 027"],
    ["gmax saset", "MHPD 027"],
    ["sachet g-max", "MHPD 027"],

    # MHPD 028 - V-MAX SACHET
    ["v-max sachet", "MHPD 028"],
    ["vmax sachet", "MHPD 028"],
    ["v max saset", "MHPD 028"],
    ["vmax saset", "MHPD 028"],
    ["sachet v-max", "MHPD 028"],
]


def setup_and_seed():
    print("==========================================================")
    print("🔄 Menghubungkan ke Google Spreadsheet...")
    print("==========================================================")

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
    client = gspread.authorize(creds)

    if "http" in SPREADSHEET_TARGET:
        sh = client.open_by_url(SPREADSHEET_TARGET)
    else:
        sh = client.open_by_key(SPREADSHEET_TARGET)

    print(f"✅ Berhasil terhubung ke: '{sh.title}'\n")

    # Helper untuk membuat atau mengambil worksheet
    def get_or_create_sheet(title: str, rows: int = 100, cols: int = 10):
        try:
            ws = sh.worksheet(title)
            print(f"ℹ️  Tab '{title}' sudah ada, memperbarui data...")
            ws.clear()
            return ws
        except gspread.exceptions.WorksheetNotFound:
            print(f"➕ Membuat Tab baru: '{title}'...")
            return sh.add_worksheet(title=title, rows=rows, cols=cols)

    # 1. Tab Master_Barang
    ws_master = get_or_create_sheet("Master_Barang", rows=50, cols=5)
    master_rows = [["Kode_SKU", "Nama_Standar"]] + MASTER_BARANG_DATA
    ws_master.update(range_name="A1", values=master_rows)
    print(f"   ✅ Tab 'Master_Barang': Berhasil memasukkan {len(MASTER_BARANG_DATA)} baris.")

    # 2. Tab Stok_Barang
    ws_stok = get_or_create_sheet("Stok_Barang", rows=50, cols=5)
    stok_rows = [["Kode_SKU", "Qty_Saat_Ini"]] + STOK_BARANG_DATA
    ws_stok.update(range_name="A1", values=stok_rows)
    print(f"   ✅ Tab 'Stok_Barang': Berhasil memasukkan {len(STOK_BARANG_DATA)} baris.")

    # 3. Tab Kamus_Alias
    ws_kamus = get_or_create_sheet("Kamus_Alias", rows=200, cols=5)
    kamus_rows = [["Nama_Berantakan", "Kode_SKU"]] + KAMUS_ALIAS_DATA
    ws_kamus.update(range_name="A1", values=kamus_rows)
    print(f"   ✅ Tab 'Kamus_Alias': Berhasil memasukkan {len(KAMUS_ALIAS_DATA)} alias variasi marketplace.")

    # 4. Tab Riwayat_Mutasi (Log Keluar Masuk)
    ws_mutasi = get_or_create_sheet("Riwayat_Mutasi", rows=1000, cols=10)
    mutasi_headers = [["Timestamp", "Tanggal", "Tipe", "Kode_SKU", "Nama_Barang", "Qty", "Sisa_Stok", "Pengirim", "Keterangan"]]
    ws_mutasi.update(range_name="A1", values=mutasi_headers)
    print(f"   ✅ Tab 'Riwayat_Mutasi': Berhasil disiapkan untuk pencatatan log transaksi.")

    print("\n==========================================================")
    print("🎉 SELESAI! Seluruh tab & data dummy berhasil disiapkan.")
    print("==========================================================")


if __name__ == "__main__":
    setup_and_seed()
