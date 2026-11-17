import os
import sys
import requests
from dotenv import load_dotenv

sys.path.insert(0, r"D:\AI Agent")
load_dotenv()

from app import telegram_client

# Test editing message or sending keyboard
print("Testing send_message_with_keyboard...")
test_keyboard = {
    "inline_keyboard": [
        [
            {"text": "📦 Katalog Sisa Stok", "callback_data": "btn_stok_p1"},
            {"text": "📥 Restok Cepat", "callback_data": "btn_restok_menu"}
        ],
        [
            {"text": "🔄 Stock Opname", "callback_data": "btn_opname_menu"},
            {"text": "➕ Tambah Master SKU", "callback_data": "btn_tambah_sku"}
        ],
        [
            {"text": "🏆 Leaderboard Seller", "callback_data": "btn_leaderboard"},
            {"text": "📊 Unduh Laporan Excel", "callback_data": "btn_export"}
        ]
    ]
}

res = telegram_client.send_message_with_keyboard(
    8979934843,
    "🎛️ <b>PANEL ADMIN GUDANG TIRA (FULL TELEGRAM)</b> 👋\n━━━━━━━━━━━━━━━━━━━━━━\nSilakan pilih menu di bawah ini langsung di Telegram:",
    reply_markup=test_keyboard
)
print("Send result:", res)
