"""
=============================================================================
Runner Uji Coba Langsung (Ultra Fast Multi-Threaded Polling Mode)
=============================================================================
Fitur Performa:
1. ⚡ MULTI-THREADED WORKER: Pesan & Tombol diproses secara paralel tanpa saling blokir.
2. 🔄 REUSABLE CONNECTION POOL: requests.Session() untuk latensi respons ultra cepat (< 0.2s).
3. 🔒 SINGLE INSTANCE LOCK: Mencegah 2 instance berjalan bersamaan yang menyebabkan bentrok Telegram.
4. 🎛️ PEMISAHAN PERAN: Mode Grup vs Japri Admin Native Telegram.
=============================================================================
"""

import io
import os
import sys
import time
import socket
import logging
import threading
import requests
from concurrent.futures import ThreadPoolExecutor
from app import (
    app,
    PORT,
    TELEGRAM_TOKEN,
    handle_incoming_message,
    handle_incoming_callback_query,
    start_background_scheduler
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("WarehousePolling")

# ThreadPool untuk eksekusi non-blocking pesan dan klik tombol
executor = ThreadPoolExecutor(max_workers=12, thread_name_prefix="TgWorker")


def start_flask_server():
    """Menjalankan server web Flask di thread terpisah agar Web Panel /panel aktif."""
    try:
        logger.info(f"Menjalankan Flask Web Server untuk Web Panel di port {PORT}...")
        app.run(host="0.0.0.0", port=PORT, debug=False, use_reloader=False)
    except Exception as e:
        logger.warning(f"Flask Web Server info: {e}")


def _dispatch_callback(callback_query: dict):
    """Worker thread untuk menangani klik tombol secara paralel dan instan."""
    try:
        handle_incoming_callback_query(callback_query)
    except Exception as e:
        logger.error(f"Error pada worker callback: {e}", exc_info=True)


def _dispatch_message(message: dict):
    """Worker thread untuk memproses pesan teks / dokumen resi secara paralel."""
    try:
        handle_incoming_message(message)
    except Exception as e:
        logger.error(f"Error pada worker pesan: {e}", exc_info=True)


def start_polling():
    if not TELEGRAM_TOKEN:
        print("❌ Error: TELEGRAM_TOKEN belum diisi di .env")
        return

    # 1. Single Instance Lock (Cegah bot ganda bertabrakan)
    lock_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        lock_socket.bind(("127.0.0.1", 49888))
    except OSError:
        print("\n⚠️ PERINGATAN: Program run_polling.py sudah aktif di sistem ini!")
        print("🛑 Menolak menjalankan instance ganda agar polling Telegram tidak bentrok/lemot.\n")
        return

    # 2. Jalankan background scheduler auto-briefing (08:00 & 17:00 WIB)
    start_background_scheduler()

    # 3. Pre-warming cache Google Sheets di latar belakang
    def _prewarm():
        try:
            from app import sheets_manager
            logger.info("⚡ Melakukan pre-warming cache Google Sheets ke RAM...")
            sheets_manager._get_kamus_aliases()
            sheets_manager._get_master_map()
            sheets_manager.get_all_products_list()
            logger.info("✅ Pre-warming cache selesai! Seluruh data siap diakses instan.")
        except Exception as exc:
            logger.warning(f"Pre-warm cache warning: {exc}")

    threading.Thread(target=_prewarm, daemon=True).start()

    # 4. Jalankan Flask Web Panel Server di background thread
    t_web = threading.Thread(target=start_flask_server, daemon=True)
    t_web.start()

    # 4. Hapus webhook terlebih dahulu agar polling bisa berjalan lancar tanpa bentrok
    session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(pool_connections=15, pool_maxsize=15, max_retries=1)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    try:
        session.get(f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/deleteWebhook", timeout=10)
    except Exception as e:
        logger.warning(f"Gagal delete webhook: {e}")

    print("=====================================================================")
    print("🚀 BOT GUDANG AKTIF (HIGH SPEED ENGINE) - TIRA (@mh_gudang_bot)")
    print("=====================================================================")
    print("Fitur Terintegrasi:")
    print("1. ⚡ ULTRA FAST CONCURRENCY: Background ThreadPool (No Blocking UI)")
    print("2. 🎛️ PEMISAHAN PERAN       : Grup (Pantau & Resi) | Japri (Panel Admin)")
    print("3. 👑 SUPERADMIN & ADMIN     : Kelola Staf, Ganti PIN & Reset Sesi Instan")
    print("4. 🔄 CACHED GOOGLE SHEETS  : Akses Stok & Alias Memori RAM (Sub-detik)")
    print(f"5. 🌐 TELEGRAM MINI APP      : Web Panel Visual (http://localhost:{PORT}/panel)")
    print("6. 🛡️ ANTI-DUPLICATE RESI    : Perlindungan O(1) di RAM")
    print("7. 📦 MULTI-ITEM RESI        : Potong beberapa produk sekaligus")
    print("8. 📊 EKSPOR EXCEL (.xlsx)   : /export atau 'kirim excel'")
    print("Tekan Ctrl + C di terminal untuk menghentikan bot.")
    print("=====================================================================\n")

    offset = None
    while True:
        try:
            url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates"
            params = {"timeout": 20}
            if offset:
                params["offset"] = offset

            resp = session.get(url, params=params, timeout=30)
            if resp.status_code != 200:
                time.sleep(1)
                continue

            data = resp.json()
            for update in data.get("result", []):
                offset = update["update_id"] + 1

                # 1. Handle Klik Tombol Inline (Callback Query) - Non-Blocking
                callback_query = update.get("callback_query")
                if callback_query:
                    cb_data = callback_query.get("data")
                    from_u = callback_query.get("from", {}).get("first_name", "User")
                    logger.info(f"Menerima Callback Query: '{cb_data}' dari {from_u} -> Diproses di ThreadPool")
                    executor.submit(_dispatch_callback, callback_query)
                    continue

                # 2. Handle Pesan Teks / Media Dokumen - Non-Blocking
                message = update.get("message")
                if not message:
                    continue

                txt = message.get("text") or message.get("caption") or "(Media/Dokumen)"
                from_u = message.get("from", {}).get("first_name", "User")
                logger.info(f"Menerima Pesan: '{txt[:30]}' dari {from_u} -> Diproses di ThreadPool")
                executor.submit(_dispatch_message, message)

        except KeyboardInterrupt:
            print("\n🛑 Bot dihentikan oleh pengguna.")
            break
        except Exception as e:
            logger.error(f"Error pada loop polling: {e}")
            time.sleep(1)


if __name__ == "__main__":
    start_polling()
