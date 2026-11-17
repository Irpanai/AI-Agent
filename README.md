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

### A. Menjalankan di Komputer Lokal (Mode Polling Cepat)
```powershell
python run_polling.py
```
*(Bot berjalan dan merespons otomatis seluruh tombol dan pesan di Telegram tanpa perlu setup domain/webhook)*.

---

### B. Deployment ke Vercel (Mode Serverless Webhook 24/7)

Proyek ini telah dikonfigurasi penuh untuk deployment **Vercel Serverless Function** menggunakan file [`vercel.json`](file:///d:/AI%20Agent/vercel.json) dan entrypoint [`api/index.py`](file:///d:/AI%20Agent/api/index.py).

#### Langkah-Langkah Deploy:
1. **Push ke GitHub:**
   Pastikan kode telah di-commit dan di-push ke repository GitHub Anda (file rahasia `.env` dan `credentials.json` sudah otomatis dilindungi oleh `.gitignore`).
2. **Import ke Vercel:**
   - Masuk ke dashboard [vercel.com](https://vercel.com/) -> klik **Add New Project** -> pilih repositori GitHub Anda.
   - Framework Preset: **Other** (Vercel akan otomatis mendeteksi konfigurasi `vercel.json` dan folder `api/`).
3. **Isi Environment Variables di Vercel Dashboard:**
   Pada bagian **Environment Variables**, tambahkan:
   - `TELEGRAM_TOKEN` : Token bot Telegram dari @BotFather.
   - `GEMINI_API_KEY` : API key Gemini AI.
   - `SPREADSHEET_ID_OR_URL` : URL atau ID Google Spreadsheet gudang.
   - `GOOGLE_SERVICE_ACCOUNT_JSON` : **(Penting)** Buka file `credentials.json` di komputer Anda, copy **seluruh isi teks JSON-nya**, lalu paste di sini sebagai nilainya.
   - `SUPERADMIN_TELEGRAM_IDS` : ID Telegram Owner (contoh: `8979934843`).
   - `ADMIN_TELEGRAM_IDS` : ID Telegram Admin (contoh: `7957273991`).
   - `ADMIN_PIN` : PIN akses admin (contoh: `1234`).
   - `WEB_APP_URL` : `https://<domain-anda>.vercel.app/panel`
4. **Klik Deploy:** Tunggu proses build selesai hingga Vercel memberikan domain HTTPS publik (misal: `https://mh-gudang.vercel.app`).
5. **Daftarkan Webhook Telegram (1-Klik):**
   Setelah deployment selesai, buka browser Anda dan kunjungi URL berikut:
   ```text
   https://<domain-anda>.vercel.app/set_webhook
   ```
   Telegram akan langsung membalas: `{"status": "ok", "message": "✅ Webhook Telegram berhasil didaftarkan!"}`.
6. **Otomatisasi Cron Jobs Vercel:**
   Jadwal **Auto-Briefing Pagi (08:00 WIB)** dan **Rekap Tutup Buku Sore (17:00 WIB)** otomatis di-trigger oleh Vercel Crons tanpa memakan kuota server standby.
