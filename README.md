# 📦 AI Agen Gudang (Tira): Telegram Bot & Warehouse Management System

Backend service cerdas berbasis **Python (Flask)**, **Google Sheets (`gspread`)**, dan **Google Gemini AI** (`@mh_gudang_bot` - Nama: **Tira**) untuk otomatisasi pencatatan resi, normalisasi SKU, pemotongan stok keluar, restok admin, stock opname, leaderboard seller, dan panel kontrol interaktif 100% di dalam Telegram.

---

## 🏛️ Arsitektur Pemisahan Mode (Grup vs Chat Pribadi / Japri)

```text
👥 GRUP TELEGRAM (Tim Lapangan & Seller)
├── 📄 Kirim Resi PDF / Foto Kurir (Potong Stok Otomatis)
├── 💬 Chat Orderan: 'Kirim 1 Dephero dan 2 Habbafit'
├── 📊 Monitoring Data: /stok, /laporan, /export Excel
├── 🏆 Leaderboard: /leaderboard (Cek Seller Teraktif Hari Ini)
└── ⛔ Perintah Admin DITOLAK (Diarahkan ke Chat Pribadi)

🔐 CHAT PRIBADI / JAPRI (Khusus Admin & Owner - 100% Native Telegram)
├── 📦 Katalog Sisa Stok: Tampilan live stok berhalaman di Telegram (Prev/Next)
├── 📥 Restok Cepat: Pilih produk & kuantitas (+10, +50, +100) langsung klik tombol
├── 🔄 Stock Opname Cepat: Pilih penyesuaian (-1 pecah, -2 rusak) via tombol
├── ➕ Tambah SKU Baru: Format copyable instan di chat
├── 🏆 Leaderboard Seller: Ranking live seller pengirim terbanyak
└── 📊 Unduh Excel: File .xlsx langsung terkirim ke chat
```

---

## 🌟 Fitur Utama (100% Langsung di Telegram)

### 1. 🎛️ Panel Kendali 100% Native Telegram (Tanpa Buka Web/Browser)
* **Katalog Sisa Stok Berhalaman:** Menampilkan 7 produk per halaman lengkap dengan indikator badge (🟢 AMAN, 🟡 KRITIS, 🔴 HABIS) dan tombol navigasi `[⬅️ Prev]` / `[Next ➡️]`.
* **Restok Cepat 1-Klik:** 
  1. Klik tombol produk (misal `[MHPD 021]`).
  2. Klik jumlah kuantitas (`[+10]`, `[+50]`, `[+100]`).
  3. Stok langsung terupdate di Google Sheets & notifikasi restok otomatis terkirim ke seluruh grup logistik!
* **Stock Opname Cepat:** 
  1. Klik tombol produk.
  2. Pilih tombol koreksi (misal `[💥 -2 pcs (Pecah/Rusak)]`).
  3. Stok langsung disesuaikan dan tercatat di tab `Riwayat_Mutasi`.
* **Ekspor Laporan Excel:** Dokumen `.xlsx` 2-tab berwarna langsung dikirim sebagai lampiran ke chat.

---

## 💬 Panduan Perintah

| Perintah | Ruang Akses | Fungsi |
| :--- | :--- | :--- |
| `/start` atau `menu` | **Japri Admin** | Buka Dashboard Panel Tombol Interaktif Native Telegram |
| Kirim Resi PDF / Foto | **Grup** | Potong stok otomatis (Anti Duplikat & Multi-Item) |
| `/stok` | Grup / Japri | Melihat sisa stok seluruh produk gudang |
| `/laporan` | Grup / Japri | Rekap ringkas mutasi hari ini |
| `/export` atau *"kirim excel"* | Grup / Japri | Mengunduh file dokumen Excel `.xlsx` |
| `/leaderboard` | Grup / Japri | Ranking pengiriman seller teraktif hari ini |
| `/restok [SKU] [Qty]` | **Khusus Japri** | Tambah stok masuk manual (atau via tombol 1-klik) |
| `/opname [SKU] [+/-Qty] [Alasan]` | **Khusus Japri** | Koreksi selisih fisik / barang rusak |
| `/tambah_sku [SKU] \| [Nama] \| [Stok]` | **Khusus Japri** | Pendaftaran master produk resmi baru |

---

## 🚀 Cara Menjalankan
```powershell
python run_polling.py
```
*(Bot berjalan dan merespons otomatis seluruh tombol dan pesan di Telegram)*.
