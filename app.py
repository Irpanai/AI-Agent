"""
=============================================================================
Sistem AI Agen Gudang - Normalisasi Stok E-Commerce (Telegram Webhook & Polling)
=============================================================================
Fitur Unggulan:
1. 🛡️ Anti-Duplicate Resi (Pencegahan Input Resi Ganda pada Hari yang Sama)
2. 📦 Multi-Item dalam 1 Resi (Orderan Campuran / Bundle Produk Sekaligus)
3. ⏰ Auto-Briefing Pagi (08:00 WIB) & Rekap Tutup Buku Sore (17:00 WIB)
4. 📊 Ekspor Laporan Excel (.xlsx) 2-Tab Berwarna Langsung ke Chat Telegram
5. 🔄 Konversi Otomatis Satuan Kemasan (Lusin x12, Kodi x20, Box x10, Dus x24 -> Pcs)
6. 🎛️ Pemisahan Mode (Grup Khusus Pantau & Resi, Japri Khusus Panel Tombol Admin)
7. 🌐 Telegram Mini App (Web Panel Visual Interaktif di dalam Telegram / Browser)
8. 🔄 Koreksi Stock Opname (Penyesuaian Selisih Fisik Gudang dengan Alasan)
9. 🏆 Leaderboard & Rekap Per-Seller (Ranking Pengiriman Resi Terbanyak Hari Ini)
10. 📢 Broadcast Notifikasi Restok Otomatis ke Grup saat Admin Melakukan Restok
=============================================================================
"""

import os
import re
import io
import base64
import json
import time
import shutil
import tempfile
import logging
import threading
from datetime import datetime, date, timedelta
from typing import Dict, Any, Optional, Tuple, List
import gspread

from flask import Flask, request, jsonify, render_template
from dotenv import load_dotenv
import requests
import pypdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from oauth2client.service_account import ServiceAccountCredentials
import google.generativeai as genai
from google.api_core import exceptions as google_exceptions

# =============================================================================
# 1. KONFIGURASI DAN LOGGING
# =============================================================================
load_dotenv()

# Deteksi lingkungan Serverless (Vercel / AWS Lambda)
IS_VERCEL = bool(os.getenv("VERCEL") or os.getenv("AWS_LAMBDA_FUNCTION_NAME"))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Di Vercel, filesystem bersifat read-only kecuali folder temp (/tmp)
DATA_DIR = tempfile.gettempdir() if IS_VERCEL else BASE_DIR

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("WarehouseBot")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
SPREADSHEET_ID_OR_URL = os.getenv("SPREADSHEET_ID_OR_URL", "").strip()
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json").strip()
PORT = int(os.getenv("PORT", 5000))
DEBUG = os.getenv("DEBUG", "False").lower() in ("true", "1", "yes")
WEB_APP_URL = os.getenv("WEB_APP_URL", f"http://localhost:{PORT}/panel").strip()

# Pengaturan Admin & Superadmin Gudang
SUPERADMIN_TELEGRAM_IDS = [
    int(x.strip()) for x in os.getenv("SUPERADMIN_TELEGRAM_IDS", "").split(",") if x.strip().isdigit()
]
ADMIN_TELEGRAM_IDS = [
    int(x.strip()) for x in os.getenv("ADMIN_TELEGRAM_IDS", "").split(",") if x.strip().isdigit()
]
# Default superadmin jika belum terdaftar secara eksplisit
if not SUPERADMIN_TELEGRAM_IDS and ADMIN_TELEGRAM_IDS:
    SUPERADMIN_TELEGRAM_IDS = [ADMIN_TELEGRAM_IDS[0]]

ADMIN_PIN = os.getenv("ADMIN_PIN", "1234").strip()
ADMIN_PHONE_NUMBERS = [
    re.sub(r"\D", "", x) for x in os.getenv("ADMIN_PHONE_NUMBERS", "").split(",") if x.strip()
]

ACTIVE_GROUPS_FILE = os.path.join(DATA_DIR, "active_groups.json")
AUTH_ADMINS_FILE = os.path.join(DATA_DIR, "authorized_admins.json")

# Inisialisasi awal file JSON di /tmp saat running di Vercel
if IS_VERCEL:
    for _fname in ["active_groups.json", "authorized_admins.json"]:
        _src = os.path.join(BASE_DIR, _fname)
        _dst = os.path.join(DATA_DIR, _fname)
        if os.path.exists(_src) and not os.path.exists(_dst):
            try:
                shutil.copyfile(_src, _dst)
            except Exception as _e:
                logger.warning(f"Gagal menyalin {_fname} ke /tmp: {_e}")


def load_authorized_admins() -> dict:
    """Memuat data Superadmin, Admin, nomor HP, dan PIN dinamis dari file JSON."""
    default_data = {
        "superadmin_ids": list(SUPERADMIN_TELEGRAM_IDS),
        "admins": {},
        "current_pin": ADMIN_PIN,
        "whitelist_ids": list(set(ADMIN_TELEGRAM_IDS + SUPERADMIN_TELEGRAM_IDS)),
        "whitelist_phones": list(ADMIN_PHONE_NUMBERS),
        "authenticated_sessions": {},
        "logged_out_ids": []
    }
    if not os.path.exists(AUTH_ADMINS_FILE):
        save_authorized_admins(default_data)
        return default_data

    try:
        with open(AUTH_ADMINS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

            # Pastikan superadmin_ids
            sa_ids = set(data.get("superadmin_ids", []))
            for said in SUPERADMIN_TELEGRAM_IDS:
                sa_ids.add(said)
            data["superadmin_ids"] = list(sa_ids)

            # Pastikan whitelist_ids
            wl_ids = set(data.get("whitelist_ids", []))
            for aid in ADMIN_TELEGRAM_IDS + list(sa_ids):
                wl_ids.add(aid)
            data["whitelist_ids"] = list(wl_ids)

            if "admins" not in data:
                data["admins"] = {}

            # Pastikan superadmin tercatat di data['admins']
            for said in sa_ids:
                s_key = str(said)
                if s_key not in data["admins"]:
                    data["admins"][s_key] = {
                        "name": "Superadmin (Owner)",
                        "role": "superadmin",
                        "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "added_by": "System"
                    }
                else:
                    data["admins"][s_key]["role"] = "superadmin"

            if "current_pin" not in data or not data["current_pin"]:
                data["current_pin"] = ADMIN_PIN

            if "whitelist_phones" not in data:
                data["whitelist_phones"] = list(ADMIN_PHONE_NUMBERS)

            if "authenticated_sessions" not in data:
                data["authenticated_sessions"] = {}
            if "logged_out_ids" not in data:
                data["logged_out_ids"] = []

            return data
    except Exception as e:
        logger.error(f"Gagal membaca file {AUTH_ADMINS_FILE}: {e}")
        return default_data


def save_authorized_admins(data: dict):
    """Menyimpan konfigurasi admin dan role ke file JSON."""
    try:
        with open(AUTH_ADMINS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        logger.error(f"Gagal menulis ke {AUTH_ADMINS_FILE}: {e}")


def get_current_pin() -> str:
    """Mengambil PIN login aktif saat ini (bisa diubah oleh Superadmin secara runtime)."""
    data = load_authorized_admins()
    return str(data.get("current_pin", ADMIN_PIN)).strip()


def set_current_pin(new_pin: str, changed_by: str = "Superadmin") -> bool:
    """Mengubah PIN masuk gudang secara runtime oleh Superadmin."""
    clean_pin = re.sub(r"\s+", "", str(new_pin)).strip()
    if not clean_pin or len(clean_pin) < 4:
        return False
    data = load_authorized_admins()
    data["current_pin"] = clean_pin
    save_authorized_admins(data)
    logger.info(f"PIN gudang berhasil diubah menjadi '{clean_pin}' oleh {changed_by}.")
    return True


def is_superadmin(user_id: int) -> bool:
    """Mengecek apakah user adalah Superadmin (Owner)."""
    if not user_id:
        return False
    if SUPERADMIN_TELEGRAM_IDS and user_id in SUPERADMIN_TELEGRAM_IDS:
        return True
    data = load_authorized_admins()
    if user_id in data.get("superadmin_ids", []):
        return True
    str_uid = str(user_id)
    if data.get("admins", {}).get(str_uid, {}).get("role") == "superadmin":
        return True
    return False


def get_user_role(user_id: int) -> str:
    """Mendapatkan role pengguna: 'superadmin', 'admin', atau 'guest'."""
    if is_superadmin(user_id):
        return "superadmin"
    if is_user_authorized(user_id):
        return "admin"
    return "guest"


def is_user_authorized(user_id: int) -> bool:
    """
    Memeriksa apakah user_id memiliki otorisasi admin / superadmin:
    1. Jika user sengaja mengetik /logout, harus login PIN dulu.
    2. Jika user_id terdaftar di whitelist_ids atau superadmin_ids.
    3. Atau jika user_id ada di authenticated_sessions (sudah login PIN).
    """
    if not user_id:
        return False
    data = load_authorized_admins()
    str_uid = str(user_id)

    if user_id in data.get("logged_out_ids", []):
        return str_uid in data.get("authenticated_sessions", {})

    if is_superadmin(user_id):
        return True

    if user_id in data.get("whitelist_ids", []):
        return True

    if str_uid in data.get("authenticated_sessions", {}):
        return True

    return False


def verify_and_login_pin(user_id: int, user_name: str, input_text: str) -> bool:
    """
    Mengecek apakah teks yang dikirim adalah PIN yang benar.
    Mendukung format: '1234', '/login 1234', '/pin 1234', 'pin=1234', 'pin: 1234'.
    """
    clean_text = input_text.strip().lower()
    for prefix in ["/login", "/pin", "pin=", "pin:"]:
        if clean_text.startswith(prefix):
            clean_text = clean_text[len(prefix):].strip()
            break

    active_pin = get_current_pin().lower()
    if clean_text == active_pin:
        data = load_authorized_admins()
        str_uid = str(user_id)
        if "authenticated_sessions" not in data:
            data["authenticated_sessions"] = {}
        data["authenticated_sessions"][str_uid] = {
            "name": user_name,
            "login_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        if str_uid not in data.get("admins", {}):
            data["admins"][str_uid] = {
                "name": user_name,
                "role": "admin",
                "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "added_by": "PIN Login"
            }
        if user_id not in data.get("whitelist_ids", []):
            data["whitelist_ids"].append(user_id)
        if user_id in data.get("logged_out_ids", []):
            data["logged_out_ids"].remove(user_id)
        save_authorized_admins(data)
        logger.info(f"Otorisasi Berhasil: {user_name} (ID: {user_id}) berhasil login dengan PIN.")
        return True
    return False


def verify_and_login_phone(user_id: int, user_name: str, phone_number: str) -> bool:
    """Memeriksa apakah nomor telepon yang dikirim cocok dengan whitelist nomor HP admin."""
    clean_phone = re.sub(r"\D", "", phone_number)
    data = load_authorized_admins()
    whitelist_phones = data.get("whitelist_phones", [])
    for wp in whitelist_phones:
        if clean_phone.endswith(wp) or wp.endswith(clean_phone):
            str_uid = str(user_id)
            if "authenticated_sessions" not in data:
                data["authenticated_sessions"] = {}
            data["authenticated_sessions"][str_uid] = {
                "name": user_name,
                "phone": clean_phone,
                "login_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            if str_uid not in data.get("admins", {}):
                data["admins"][str_uid] = {
                    "name": user_name,
                    "role": "admin",
                    "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "added_by": f"Phone ({clean_phone})"
                }
            if user_id not in data.get("whitelist_ids", []):
                data["whitelist_ids"].append(user_id)
            if user_id in data.get("logged_out_ids", []):
                data["logged_out_ids"].remove(user_id)
            save_authorized_admins(data)
            logger.info(f"Otorisasi Berhasil: {user_name} (ID: {user_id}, No: {clean_phone}) login via Phone.")
            return True
    return False


def logout_admin_session(user_id: int):
    """Mengunci akses admin (logout) sehingga memerlukan input PIN untuk masuk kembali."""
    data = load_authorized_admins()
    str_uid = str(user_id)
    if str_uid in data.get("authenticated_sessions", {}):
        del data["authenticated_sessions"][str_uid]
    if user_id not in data.get("logged_out_ids", []):
        data["logged_out_ids"].append(user_id)
    save_authorized_admins(data)


def add_admin_member(target_id: int, name: str = "", role: str = "admin", added_by: str = "Superadmin") -> bool:
    """Menambahkan staf admin baru dengan nama dan role."""
    data = load_authorized_admins()
    str_id = str(target_id)
    if target_id not in data["whitelist_ids"]:
        data["whitelist_ids"].append(target_id)
    data["admins"][str_id] = {
        "name": name.strip() or f"Admin-{target_id}",
        "role": role,
        "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "added_by": added_by
    }
    if target_id in data.get("logged_out_ids", []):
        data["logged_out_ids"].remove(target_id)
    save_authorized_admins(data)
    logger.info(f"Admin baru ditambahkan: {name} (ID: {target_id}) role: {role} oleh {added_by}")
    return True


def remove_admin_member(target_id: int) -> Tuple[bool, str]:
    """Menghapus hak akses admin. Superadmin tidak dapat dihapus."""
    if is_superadmin(target_id):
        return False, "Tidak dapat mencabut hak akses Superadmin (Owner)!"
    data = load_authorized_admins()
    str_id = str(target_id)
    removed_any = False
    if target_id in data.get("whitelist_ids", []):
        data["whitelist_ids"].remove(target_id)
        removed_any = True
    if str_id in data.get("admins", {}):
        del data["admins"][str_id]
        removed_any = True
    if str_id in data.get("authenticated_sessions", {}):
        del data["authenticated_sessions"][str_id]
        removed_any = True
    if target_id not in data.get("logged_out_ids", []):
        data["logged_out_ids"].append(target_id)
    save_authorized_admins(data)
    if removed_any:
        return True, f"Hak akses admin untuk ID {target_id} berhasil dicabut!"
    return False, f"ID {target_id} tidak ditemukan dalam daftar admin."


def reset_all_admin_sessions(except_superadmin: bool = True) -> int:
    """Mereset seluruh sesi login admin lain sehingga harus memasukkan PIN baru."""
    data = load_authorized_admins()
    count = 0
    sessions = dict(data.get("authenticated_sessions", {}))
    for str_uid in sessions:
        uid = int(str_uid)
        if except_superadmin and is_superadmin(uid):
            continue
        del data["authenticated_sessions"][str_uid]
        if uid not in data.get("logged_out_ids", []):
            data["logged_out_ids"].append(uid)
        count += 1
    save_authorized_admins(data)
    return count


def check_is_admin(user_id: int, text: str = "") -> bool:
    """Wrapper kompatibilitas untuk memeriksa otorisasi admin."""
    if is_user_authorized(user_id):
        return True
    if ADMIN_PIN and (ADMIN_PIN.lower() in text.lower() or f"pin={ADMIN_PIN}".lower() in text.lower()):
        return True
    return False


# Helper Manajemen Grup Aktif untuk Broadcast Otomatis
def save_active_group(chat_id: int, title: str = ""):
    """Menyimpan ID grup Telegram aktif agar bisa menerima briefing pagi, rekap sore, dan info restok."""
    try:
        groups = {}
        if os.path.exists(ACTIVE_GROUPS_FILE):
            with open(ACTIVE_GROUPS_FILE, "r", encoding="utf-8") as f:
                groups = json.load(f)

        chat_key = str(chat_id)
        if chat_key not in groups:
            groups[chat_key] = {
                "chat_id": chat_id,
                "title": title or "Grup Gudang",
                "registered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            with open(ACTIVE_GROUPS_FILE, "w", encoding="utf-8") as f:
                json.dump(groups, f, indent=2, ensure_ascii=False)
            logger.info(f"Grup baru terdaftar untuk broadcast: {title} ({chat_id})")
    except Exception as e:
        logger.warning(f"Gagal menyimpan active group: {e}")


def get_active_group_ids() -> List[int]:
    """Mengambil seluruh chat_id grup yang terdaftar."""
    try:
        if os.path.exists(ACTIVE_GROUPS_FILE):
            with open(ACTIVE_GROUPS_FILE, "r", encoding="utf-8") as f:
                groups = json.load(f)
                return [item["chat_id"] for item in groups.values()]
    except Exception as e:
        logger.warning(f"Gagal membaca active groups: {e}")
    return []


app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static")
)


# =============================================================================
# 2. HELPER: REGEX EKSTRAKSI KUANTITAS (QTY) & SATUAN KEMASAN
# =============================================================================
def extract_qty_from_text(text: str, kode_sku: str = "") -> int:
    """
    Mengekstrak kuantitas (QTY) dari teks pesan seller dengan cerdas dan
    mengonversi satuan kemasan grosir ke PCS secara otomatis:
    - 1 Lusin = 12 pcs
    - 1 Kodi = 20 pcs
    - 1 Dus / Karton = 24 pcs
    - 1 Box / Kotak G-MAX (MHPD 027) = 10 sachet/pcs
    - 1 Box / Kotak V-MAX (MHPD 028) = 10 sachet/pcs
    - 1 Box umum = 10 pcs
    """
    if not text:
        return 1

    clean_text = text.lower().strip()

    # 1. Cek satuan kemasan Lusin (x 12)
    match_lusin = re.search(r'(\d+)\s*lusin\b', clean_text)
    if match_lusin:
        return max(1, int(match_lusin.group(1)) * 12)

    # 2. Cek satuan kemasan Kodi (x 20)
    match_kodi = re.search(r'(\d+)\s*kodi\b', clean_text)
    if match_kodi:
        return max(1, int(match_kodi.group(1)) * 20)

    # 3. Cek satuan Dus / Karton (x 24)
    match_dus = re.search(r'(\d+)\s*(?:dus|karton)\b', clean_text)
    if match_dus:
        return max(1, int(match_dus.group(1)) * 24)

    # 4. Cek satuan Box / Kotak (10 pcs per box)
    match_box = re.search(r'(\d+)\s*(?:box|kotak)\b', clean_text)
    if match_box:
        box_count = int(match_box.group(1))
        return max(1, box_count * 10)

    # 5. Angka yang diikuti satuan kuantitas standar
    unit_pattern = (
        r'(?:^|\s|[x*])(\d+)\s*'
        r'(?:pcs|pc|buah|biji|item|pack|paket|btl|botol|sachet|sak|lembar|pasang)\b'
    )
    match_unit = re.search(unit_pattern, clean_text, re.IGNORECASE)
    if match_unit:
        return max(1, int(match_unit.group(1)))

    # 6. Angka setelah keyword jumlah / perkalian
    keyword_pattern = r'(?:qty|jumlah|total|ambil|x|\*)\s*[:=]?\s*(\d+)\b'
    match_keyword = re.search(keyword_pattern, clean_text, re.IGNORECASE)
    if match_keyword:
        return max(1, int(match_keyword.group(1)))

    # 7. Angka murni di akhir kalimat
    end_pattern = r'\b(\d+)\s*$'
    match_end = re.search(end_pattern, clean_text)
    if match_end:
        return max(1, int(match_end.group(1)))

    # 8. Angka yang TIDAK diikuti satuan ukuran fisik (ml, l, g, gr, kg, oz, cm, m)
    candidates = re.findall(
        r'\b(\d+)\b(?!\s*(?:ml|l|liter|g|gr|gram|kg|oz|mm|cm|m)\b)',
        clean_text,
        re.IGNORECASE
    )
    if candidates:
        return max(1, int(candidates[-1]))

    return 1


def extract_no_resi_from_text(text: str) -> Optional[str]:
    """
    Mengekstrak nomor resi kurir dari teks pesan atau hasil bacaan dokumen.
    Mendukung format Anteraja, Shopee Xpress (SPX), J&T Express, SiCepat, JNE, dsb.
    """
    if not text:
        return None

    pattern_label = r'(?i)(?:no\.?\s*resi|nomor\s*resi|resi|awb|tracking\s*no\.?|no\.?\s*pesanan)\s*[:#\s]*([A-Za-z0-9_-]{8,24})'
    m_label = re.search(pattern_label, text)
    if m_label:
        return m_label.group(1).strip()

    m_spx = re.search(r'\b(SPX[A-Za-z0-9]{8,18})\b', text, re.IGNORECASE)
    if m_spx:
        return m_spx.group(1).strip()

    m_jnt = re.search(r'\b(J[PX]\d{9,14})\b', text, re.IGNORECASE)
    if m_jnt:
        return m_jnt.group(1).strip()

    m_num = re.search(r'\b(1\d{11,15})\b', text)
    if m_num:
        return m_num.group(1).strip()

    m_sicepat = re.search(r'\b(00\d{10,14})\b', text)
    if m_sicepat:
        return m_sicepat.group(1).strip()

    return None


# =============================================================================
# 3. HELPER: TELEGRAM BOT API INTERACTION
# =============================================================================
class TelegramClient:
    """Kelas pembungkus untuk memanggil Telegram Bot API dengan Persistent Session & Connection Pooling."""

    def __init__(self, token: str):
        self.token = token
        self.api_url = f"https://api.telegram.org/bot{token}"
        self.session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=25,
            pool_maxsize=25,
            max_retries=1
        )
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

    def send_message(
        self,
        chat_id: int,
        text: str,
        reply_to_message_id: Optional[int] = None,
        parse_mode: Optional[str] = None
    ) -> bool:
        """Mengirim pesan teks ke chat atau grup Telegram."""
        if not self.token:
            logger.error("Telegram token belum dikonfigurasi.")
            return False

        endpoint = f"{self.api_url}/sendMessage"
        payload = {
            "chat_id": chat_id,
            "text": text,
        }
        if reply_to_message_id:
            payload["reply_to_message_id"] = reply_to_message_id
        if parse_mode:
            payload["parse_mode"] = parse_mode

        for attempt in range(1, 4):
            try:
                response = self.session.post(endpoint, json=payload, timeout=15)
                data = response.json()
                if not data.get("ok"):
                    logger.error(f"Telegram API gagal mengirim pesan: {data}")
                    return False
                return True
            except requests.RequestException as e:
                logger.warning(f"Telegram sendMessage attempt {attempt} error: {e}")
                time.sleep(1.0)

        return False

    def send_message_with_keyboard(
        self,
        chat_id: int,
        text: str,
        reply_markup: dict,
        parse_mode: Optional[str] = "HTML"
    ) -> bool:
        """Mengirim pesan dengan Inline Keyboard (Tombol Interaktif) ke chat Telegram."""
        if not self.token:
            return False

        endpoint = f"{self.api_url}/sendMessage"
        payload = {
            "chat_id": chat_id,
            "text": text,
            "reply_markup": reply_markup
        }
        if parse_mode:
            payload["parse_mode"] = parse_mode

        for attempt in range(1, 4):
            try:
                response = self.session.post(endpoint, json=payload, timeout=15)
                data = response.json()
                if not data.get("ok"):
                    logger.error(f"Telegram send_message_with_keyboard error: {data}")
                    return False
                return True
            except Exception as e:
                logger.warning(f"Telegram send_message_with_keyboard attempt {attempt} error: {e}")
                time.sleep(1.0)
        return False

    def answer_callback_query(self, callback_query_id: str, text: Optional[str] = None, show_alert: bool = False) -> bool:
        """Menjawab event klik tombol callback dari pengguna dengan respons instan."""
        if not self.token:
            return False
        endpoint = f"{self.api_url}/answerCallbackQuery"
        payload = {"callback_query_id": callback_query_id}
        if text:
            payload["text"] = text
        if show_alert:
            payload["show_alert"] = True
        try:
            self.session.post(endpoint, json=payload, timeout=6)
            return True
        except Exception as e:
            logger.warning(f"Error answerCallbackQuery: {e}")
            return False

    def edit_message_text(
        self,
        chat_id: int,
        message_id: int,
        text: str,
        reply_markup: Optional[dict] = None,
        parse_mode: Optional[str] = "HTML"
    ) -> bool:
        """Mengubah teks & tombol pada pesan yang sudah ada (Smooth In-Place Telegram UI)."""
        if not self.token:
            return False
        endpoint = f"{self.api_url}/editMessageText"
        payload = {
            "chat_id": chat_id,
            "message_id": message_id,
            "text": text,
        }
        if reply_markup:
            payload["reply_markup"] = reply_markup
        if parse_mode:
            payload["parse_mode"] = parse_mode

        for attempt in range(1, 4):
            try:
                resp = self.session.post(endpoint, json=payload, timeout=15)
                data = resp.json()
                if not data.get("ok"):
                    if "message is not modified" in str(data):
                        return True
                    logger.warning(f"Telegram editMessageText error: {data}")
                    return False
                return True
            except Exception as e:
                logger.warning(f"editMessageText attempt {attempt} error: {e}")
                time.sleep(0.8)
        return False

    def send_document(
        self,
        chat_id: int,
        file_bytes: bytes,
        filename: str,
        caption: Optional[str] = None,
        reply_to_message_id: Optional[int] = None,
        parse_mode: Optional[str] = "HTML"
    ) -> bool:
        """Mengirim dokumen/file (seperti .xlsx) langsung ke chat Telegram."""
        if not self.token:
            return False

        endpoint = f"{self.api_url}/sendDocument"
        data = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        if parse_mode:
            data["parse_mode"] = parse_mode
        if reply_to_message_id:
            data["reply_to_message_id"] = reply_to_message_id

        files = {
            "document": (
                filename,
                io.BytesIO(file_bytes),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        }

        for attempt in range(1, 4):
            try:
                response = self.session.post(endpoint, data=data, files=files, timeout=35)
                res_data = response.json()
                if not res_data.get("ok"):
                    logger.error(f"Telegram sendDocument gagal: {res_data}")
                    return False
                logger.info(f"Berhasil mengirim dokumen '{filename}' ke chat {chat_id}")
                return True
            except Exception as e:
                logger.warning(f"Telegram sendDocument attempt {attempt} error: {e}")
                time.sleep(1.5)

        return False

    def get_file_bytes(self, file_id: str) -> Optional[bytes]:
        """Mengunduh file dari server Telegram berdasarkan file_id."""
        if not self.token:
            return None
        for attempt in range(1, 4):
            try:
                get_file_url = f"{self.api_url}/getFile?file_id={file_id}"
                res = self.session.get(get_file_url, timeout=15).json()
                if not res.get("ok"):
                    logger.error(f"Telegram getFile gagal: {res}")
                    return None
                file_path = res["result"]["file_path"]
                download_url = f"https://api.telegram.org/file/bot{self.token}/{file_path}"
                r = self.session.get(download_url, timeout=30)
                if r.status_code == 200:
                    return r.content
                return None
            except Exception as e:
                logger.warning(f"Telegram get_file_bytes attempt {attempt} error: {e}")
                time.sleep(1.0)
        return None

    def set_webhook(self, webhook_url: str) -> Tuple[bool, dict]:
        """Mendaftarkan URL webhook ke Telegram."""
        if not self.token:
            return False, {"error": "Token Telegram belum dikonfigurasi"}
        endpoint = f"{self.api_url}/setWebhook"
        payload = {
            "url": webhook_url,
            "allowed_updates": ["message", "callback_query"],
            "drop_pending_updates": False
        }
        try:
            resp = self.session.post(endpoint, json=payload, timeout=15)
            data = resp.json()
            return bool(data.get("ok")), data
        except Exception as e:
            return False, {"error": str(e)}

    def get_webhook_info(self) -> dict:
        """Mengambil informasi status webhook Telegram saat ini."""
        if not self.token:
            return {"error": "Token Telegram belum dikonfigurasi"}
        endpoint = f"{self.api_url}/getWebhookInfo"
        try:
            resp = self.session.get(endpoint, timeout=15)
            return resp.json()
        except Exception as e:
            return {"error": str(e)}

    def delete_webhook(self) -> Tuple[bool, dict]:
        """Menghapus webhook Telegram (berguna jika ingin kembali ke mode polling lokal)."""
        if not self.token:
            return False, {"error": "Token Telegram belum dikonfigurasi"}
        endpoint = f"{self.api_url}/deleteWebhook"
        try:
            resp = self.session.post(endpoint, json={}, timeout=15)
            data = resp.json()
            return bool(data.get("ok")), data
        except Exception as e:
            return False, {"error": str(e)}


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    """Mengekstrak teks dari file PDF resi."""
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        extracted_texts = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                extracted_texts.append(t)
        return "\n".join(extracted_texts).strip()
    except Exception as e:
        logger.error(f"Gagal mengekstrak teks dari PDF: {e}")
        return ""


def extract_content_from_message(message: dict) -> Tuple[str, Optional[bytes], Optional[str]]:
    """Mengekstrak teks, PDF resi, atau foto dari message Telegram."""
    caption_or_text = message.get("text") or message.get("caption") or ""

    document = message.get("document")
    if document:
        file_id = document.get("file_id")
        file_name = document.get("file_name", "")
        mime_type = document.get("mime_type", "")
        logger.info(f"Menerima dokumen resi: '{file_name}' ({mime_type})")

        file_bytes = telegram_client.get_file_bytes(file_id)
        if file_bytes:
            if file_name.lower().endswith(".pdf") or "pdf" in mime_type.lower():
                extracted_pdf_text = extract_text_from_pdf_bytes(file_bytes)
                combined_text = f"Nama File Resi: {file_name}\nIsi Teks Resi:\n{extracted_pdf_text}\n{caption_or_text}".strip()
                logger.info(f"Hasil ekstraksi teks dari PDF ({len(extracted_pdf_text)} karakter).")
                if len(extracted_pdf_text) > 10:
                    return combined_text, None, None
                else:
                    return combined_text, file_bytes, "application/pdf"
            elif any(mime_type.lower().startswith(p) for p in ["image/", "photo/"]):
                return caption_or_text or f"Foto resi: {file_name}", file_bytes, mime_type

    photos = message.get("photo")
    if photos and isinstance(photos, list):
        largest_photo = photos[-1]
        file_id = largest_photo.get("file_id")
        logger.info("Menerima foto resi dari Telegram.")
        img_bytes = telegram_client.get_file_bytes(file_id)
        if img_bytes:
            return caption_or_text or "Foto resi kurir", img_bytes, "image/jpeg"

    return caption_or_text.strip(), None, None


telegram_client = TelegramClient(TELEGRAM_TOKEN)


# =============================================================================
# 4. HELPER: GOOGLE SHEETS MANAGER (gspread & oauth2client)
# =============================================================================
class GoogleSheetsManager:
    """Mengelola seluruh pembacaan dan pembaruan data di Google Sheets."""

    SCOPE = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    TAB_MASTER = "Master_Barang"
    TAB_STOK = "Stok_Barang"
    TAB_KAMUS = "Kamus_Alias"
    TAB_MUTASI = "Riwayat_Mutasi"

    def __init__(self, creds_path: str, spreadsheet_target: str):
        self.creds_path = creds_path
        self.spreadsheet_target = spreadsheet_target
        self.client: Optional[gspread.Client] = None
        self.spreadsheet: Optional[gspread.Spreadsheet] = None
        self._worksheets: Dict[str, gspread.Worksheet] = {}

        # In-Memory High Speed Caching
        self._kamus_cache: Optional[List[Tuple[str, str]]] = None
        self._kamus_cache_time: float = 0.0

        self._master_cache: Optional[Dict[str, str]] = None
        self._master_cache_time: float = 0.0

        self._stok_cache: Optional[Dict[str, int]] = None
        self._stok_cache_time: float = 0.0

        self._today_resi_cache: Dict[str, Dict[str, str]] = {}
        self._today_resi_cache_date: str = ""

    def _authenticate(self):
        """Autentikasi Service Account JSON via oauth2client & gspread (mendukung file fisik atau ENV var untuk Vercel)."""
        creds_json_str = (
            os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
            or os.getenv("GOOGLE_CREDENTIALS_JSON")
            or ""
        ).strip()

        if creds_json_str:
            try:
                key_dict = json.loads(creds_json_str)
                creds = ServiceAccountCredentials.from_json_keyfile_dict(key_dict, self.SCOPE)
            except Exception as e:
                logger.error(f"Gagal memproses GOOGLE_SERVICE_ACCOUNT_JSON env: {e}")
                raise ValueError(f"Format JSON pada GOOGLE_SERVICE_ACCOUNT_JSON tidak valid: {e}")
        else:
            target_path = self.creds_path
            if not os.path.isabs(target_path):
                target_path = os.path.join(BASE_DIR, target_path)

            if not os.path.exists(target_path):
                raise FileNotFoundError(
                    f"File kredensial '{self.creds_path}' tidak ditemukan! "
                    "Jika di-deploy ke Vercel, salin isi file kredensial ke Environment Variable 'GOOGLE_SERVICE_ACCOUNT_JSON'."
                )
            creds = ServiceAccountCredentials.from_json_keyfile_name(target_path, self.SCOPE)

        self.client = gspread.authorize(creds)

        if "http" in self.spreadsheet_target:
            self.spreadsheet = self.client.open_by_url(self.spreadsheet_target)
        else:
            self.spreadsheet = self.client.open_by_key(self.spreadsheet_target)
        self._worksheets.clear()
        logger.info(f"Berhasil terhubung ke Spreadsheet: '{self.spreadsheet.title}'")

    def _execute_with_retry(self, func, *args, max_retries: int = 3, **kwargs):
        """Mengeksekusi fungsi gspread dengan retry dan exponential backoff."""
        delay = 2.0
        for attempt in range(1, max_retries + 1):
            try:
                if not self.spreadsheet:
                    self._authenticate()
                return func(*args, **kwargs)
            except gspread.exceptions.APIError as e:
                error_code = getattr(e.response, "status_code", None)
                if error_code == 429 or "RESOURCE_EXHAUSTED" in str(e):
                    logger.warning(
                        f"Google Sheets Quota Exceeded (Attempt {attempt}/{max_retries}). Tidur {delay}s..."
                    )
                    time.sleep(delay)
                    delay *= 2
                    if attempt == max_retries:
                        raise e
                else:
                    logger.error(f"Gspread API Error: {e}")
                    raise e
            except Exception as e:
                logger.warning(f"Error gspread ({e}), mencoba re-autentikasi...")
                try:
                    self._authenticate()
                    return func(*args, **kwargs)
                except Exception as inner_e:
                    if attempt == max_retries:
                        raise inner_e
                    time.sleep(delay)

    def get_worksheet(self, tab_name: str) -> gspread.Worksheet:
        """Mengambil worksheet dengan caching agar tidak re-fetch metadata berulang-ulang."""
        if tab_name in self._worksheets and self._worksheets[tab_name] is not None:
            return self._worksheets[tab_name]

        def _get():
            ws = self.spreadsheet.worksheet(tab_name)
            self._worksheets[tab_name] = ws
            return ws

        return self._execute_with_retry(_get)

    def _get_kamus_aliases(self) -> List[Tuple[str, str]]:
        """Mengambil kamus alias dengan in-memory cache TTL 5 menit (Super Cepat)."""
        now = time.time()
        if self._kamus_cache is not None and (now - self._kamus_cache_time) < 300:
            return self._kamus_cache

        def _fetch():
            ws = self.get_worksheet(self.TAB_KAMUS)
            records = ws.get_all_values()
            aliases = []
            for row in records[1:]:
                if len(row) >= 2 and row[0].strip():
                    aliases.append((row[0].strip().lower(), row[1].strip()))
            return aliases

        try:
            aliases = self._execute_with_retry(_fetch)
            self._kamus_cache = aliases
            self._kamus_cache_time = now
            return aliases
        except Exception as e:
            logger.warning(f"Gagal refresh kamus alias: {e}")
            return self._kamus_cache or []

    def search_kamus_alias(self, pesan_seller: str) -> Optional[str]:
        """Pencarian alias lokal di RAM (0.0001s) tanpa menunggu round-trip Google Sheets."""
        clean_pesan = pesan_seller.lower().strip()
        aliases = self._get_kamus_aliases()
        if not aliases:
            return None

        for nama_alias, kode_sku in aliases:
            if nama_alias == clean_pesan:
                return kode_sku

        aliases_sorted = sorted(aliases, key=lambda x: len(x[0]), reverse=True)
        for nama_alias, kode_sku in aliases_sorted:
            if len(nama_alias) >= 3 and nama_alias in clean_pesan:
                return kode_sku

        return None

    def _get_master_map(self) -> Dict[str, str]:
        """Mengambil mapping SKU -> Nama Standar dengan in-memory cache TTL 5 menit."""
        now = time.time()
        if self._master_cache is not None and (now - self._master_cache_time) < 300:
            return self._master_cache

        def _fetch():
            ws = self.get_worksheet(self.TAB_MASTER)
            records = ws.get_all_values()
            m_map = {}
            for r in records[1:]:
                if len(r) >= 2 and r[0].strip():
                    m_map[r[0].strip().upper()] = r[1].strip()
            return m_map

        try:
            m_map = self._execute_with_retry(_fetch)
            self._master_cache = m_map
            self._master_cache_time = now
            return m_map
        except Exception as e:
            logger.warning(f"Gagal refresh master barang: {e}")
            return self._master_cache or {}

    def get_master_barang_reference(self) -> str:
        m_map = self._get_master_map()
        if not m_map:
            return "Belum ada data di Master_Barang."
        items_ref = [f"- SKU: {sku} | Nama: {nama}" for sku, nama in m_map.items()]
        return "\n".join(items_ref)

    def get_nama_standar_by_sku(self, kode_sku: str) -> str:
        clean_sku = kode_sku.strip().upper()
        m_map = self._get_master_map()
        return m_map.get(clean_sku, clean_sku)

    def get_stok_by_sku(self, kode_sku: str) -> int:
        """Mengambil sisa stok fisik saat ini untuk SKU tertentu secara cepat."""
        clean_sku = kode_sku.strip().upper()
        now = time.time()
        if self._stok_cache is not None and (now - self._stok_cache_time) < 20:
            return self._stok_cache.get(clean_sku, 0)

        items = self.get_all_products_list()
        for it in items:
            if it.get("sku") == clean_sku:
                return it.get("qty", 0)
        return 0

    def check_duplicate_resi(
        self,
        no_resi: str,
        target_date: Optional[str] = None
    ) -> Optional[Dict[str, str]]:
        if not no_resi or len(no_resi.strip()) < 5:
            return None

        clean_resi = no_resi.strip().lower()
        today_str = datetime.now().strftime("%Y-%m-%d")
        if not target_date:
            target_date = today_str

        # 1. Cek dari in-memory cache hari ini (Super Cepat 0.0001s)
        if target_date == today_str and self._today_resi_cache_date == today_str and self._today_resi_cache:
            if clean_resi in self._today_resi_cache:
                return self._today_resi_cache[clean_resi]
            return None

        def _check():
            try:
                ws = self.get_worksheet(self.TAB_MUTASI)
                records = ws.get_all_values()
            except gspread.exceptions.WorksheetNotFound:
                return None

            today_cache = {}
            matched_info = None

            for row in records[1:]:
                if len(row) >= 9:
                    row_tanggal = row[1].strip()
                    keterangan = row[8].strip().lower()

                    if (target_date in row_tanggal or row_tanggal.startswith(target_date)):
                        ts = row[0].strip()
                        jam = ts.split()[1][:5] if " " in ts else ts
                        pengirim = row[7].strip() or "Seller"

                        resi_found = extract_no_resi_from_text(row[8])
                        if resi_found:
                            today_cache[resi_found.lower()] = {
                                "timestamp": ts,
                                "jam": jam,
                                "pengirim": pengirim,
                                "no_resi": resi_found
                            }

                        if clean_resi in keterangan or any(clean_resi in str(c).lower() for c in row):
                            matched_info = {
                                "timestamp": ts,
                                "jam": jam,
                                "pengirim": pengirim,
                                "no_resi": no_resi.strip()
                            }

            if target_date == today_str:
                self._today_resi_cache = today_cache
                self._today_resi_cache_date = today_str

            return matched_info

        return self._execute_with_retry(_check)

    def update_or_create_stok(
        self,
        kode_sku: str,
        qty: int,
        action: str = "KELUAR"
    ) -> Tuple[int, int, bool, bool]:
        def _update():
            ws = self.get_worksheet(self.TAB_STOK)
            records = ws.get_all_values()

            target_row = None
            current_qty = 0
            clean_sku = kode_sku.strip().upper()

            for index, row in enumerate(records[1:], start=2):
                if len(row) >= 1 and row[0].strip().upper() == clean_sku:
                    target_row = index
                    try:
                        current_qty = int(row[1]) if len(row) >= 2 and row[1].strip() else 0
                    except ValueError:
                        current_qty = 0
                    break

            is_stok_kurang = False
            if action.upper() == "KELUAR":
                if current_qty < qty:
                    is_stok_kurang = True
                new_qty = max(0, current_qty - qty)
            else:
                new_qty = current_qty + qty

            if target_row:
                ws.update_cell(target_row, 2, new_qty)
                logger.info(f"Stok diupdate ({action}): SKU {clean_sku} dari {current_qty} -> {new_qty}")
            else:
                ws.append_row([clean_sku, new_qty])
                logger.info(f"SKU baru di Stok_Barang: SKU {clean_sku}, Stok: {new_qty}")

            if self._stok_cache is not None:
                self._stok_cache[clean_sku] = new_qty
                self._stok_cache_time = time.time()

            return current_qty, new_qty, is_stok_kurang, (target_row is None)

        return self._execute_with_retry(_update)

    def perform_stock_opname(
        self,
        kode_sku: str,
        diff_qty: int,
        alasan: str,
        admin_name: str = "Admin"
    ) -> Tuple[int, int, bool, str]:
        """
        Melakukan koreksi stock opname penyesuaian stok fisik (barang rusak, pecah, atau selisih).
        diff_qty bisa bertanda minus (berkurang) atau positif (bertambah).
        Return: (current_qty, new_qty, success, nama_barang)
        """
        def _opname():
            ws_stok = self.get_worksheet(self.TAB_STOK)
            records = ws_stok.get_all_values()

            target_row = None
            current_qty = 0
            clean_sku = kode_sku.strip().upper()

            for index, row in enumerate(records[1:], start=2):
                if len(row) >= 1 and row[0].strip().upper() == clean_sku:
                    target_row = index
                    try:
                        current_qty = int(row[1]) if len(row) >= 2 and row[1].strip() else 0
                    except ValueError:
                        current_qty = 0
                    break

            if not target_row:
                return current_qty, current_qty, False, clean_sku

            new_qty = max(0, current_qty + diff_qty)
            ws_stok.update_cell(target_row, 2, new_qty)

            if self._stok_cache is not None:
                self._stok_cache[clean_sku] = new_qty
                self._stok_cache_time = time.time()

            nama_barang = self.get_nama_standar_by_sku(clean_sku)
            tanda = "+" if diff_qty >= 0 else ""
            keterangan = f"Stock Opname ({tanda}{diff_qty}): {alasan.strip()}"

            self.record_mutasi(
                tipe="OPNAME",
                kode_sku=clean_sku,
                nama_barang=nama_barang,
                qty=abs(diff_qty),
                sisa_stok=new_qty,
                pengirim=admin_name,
                keterangan=keterangan
            )
            logger.info(f"Stock Opname Sukses: {clean_sku} {tanda}{diff_qty} pcs (Stok: {current_qty} -> {new_qty}). Alasan: {alasan}")
            return current_qty, new_qty, True, nama_barang

        return self._execute_with_retry(_opname)

    def record_mutasi(
        self,
        tipe: str,
        kode_sku: str,
        nama_barang: str,
        qty: int,
        sisa_stok: int,
        pengirim: str = "",
        keterangan: str = ""
    ) -> bool:
        def _record():
            try:
                ws = self.get_worksheet(self.TAB_MUTASI)
            except gspread.exceptions.WorksheetNotFound:
                ws = self.spreadsheet.add_worksheet(title=self.TAB_MUTASI, rows=1000, cols=10)
                self._worksheets[self.TAB_MUTASI] = ws
                ws.append_row([
                    "Timestamp", "Tanggal", "Tipe", "Kode_SKU", "Nama_Barang",
                    "Qty", "Sisa_Stok", "Pengirim", "Keterangan"
                ])

            now = datetime.now()
            row_data = [
                now.strftime("%d/%m/%Y %H:%M:%S"),
                now.strftime("%Y-%m-%d"),
                tipe.upper(),
                kode_sku.strip().upper(),
                nama_barang.strip(),
                qty,
                sisa_stok,
                pengirim.strip(),
                keterangan.strip()
            ]
            ws.append_row(row_data)

            # Update cache resi hari ini jika baris memuat resi
            clean_resi = extract_no_resi_from_text(keterangan)
            if clean_resi:
                self._today_resi_cache[clean_resi.lower()] = {
                    "timestamp": row_data[0],
                    "jam": now.strftime("%H:%M"),
                    "pengirim": pengirim.strip(),
                    "no_resi": clean_resi
                }
                self._today_resi_cache_date = now.strftime("%Y-%m-%d")

            logger.info(f"Mutasi dicatat: [{tipe}] SKU {kode_sku} Qty {qty} Sisa {sisa_stok}")
            return True

        return self._execute_with_retry(_record)

    def add_master_sku(self, kode_sku: str, nama_standar: str, stok_awal: int = 0) -> bool:
        def _add():
            clean_sku = kode_sku.strip().upper()
            clean_nama = nama_standar.strip()

            ws_master = self.get_worksheet(self.TAB_MASTER)
            records = ws_master.get_all_values()
            for row in records[1:]:
                if len(row) >= 1 and row[0].strip().upper() == clean_sku:
                    return False

            ws_master.append_row([clean_sku, clean_nama])

            ws_stok = self.get_worksheet(self.TAB_STOK)
            ws_stok.append_row([clean_sku, max(0, stok_awal)])

            self._master_cache = None
            self._stok_cache = None

            if stok_awal > 0:
                self.record_mutasi(
                    tipe="MASUK",
                    kode_sku=clean_sku,
                    nama_barang=clean_nama,
                    qty=stok_awal,
                    sisa_stok=stok_awal,
                    pengirim="Admin",
                    keterangan="Pendaftaran SKU Baru & Stok Awal"
                )
            return True

        return self._execute_with_retry(_add)

    def get_morning_briefing_text(self) -> str:
        def _briefing():
            ws_stok = self.spreadsheet.worksheet(self.TAB_STOK)
            stok_records = ws_stok.get_all_values()

            ws_master = self.spreadsheet.worksheet(self.TAB_MASTER)
            master_records = ws_master.get_all_values()
            nama_map = {r[0].strip().upper(): r[1].strip() for r in master_records[1:] if len(r) >= 2}

            low_stock_items = []
            for row in stok_records[1:]:
                if len(row) >= 2 and row[0].strip():
                    sku = row[0].strip().upper()
                    try:
                        sisa = int(row[1]) if row[1].strip() else 0
                        if sisa < 10:
                            nama = nama_map.get(sku, sku)
                            low_stock_items.append((sku, nama, sisa))
                    except ValueError:
                        pass

            now = datetime.now()
            tgl_display = now.strftime("%d %B %Y")

            msg_lines = [
                "☀️ <b>SELAMAT PAGI TIM GUDANG &amp; SELLER!</b> 👋",
                "━━━━━━━━━━━━━━━━━━━━━━",
                f"📅 <i>{tgl_display} | Jam 08:00 WIB</i>",
                "Stok awal siap diproses untuk aktivitas hari ini.\n"
            ]

            if low_stock_items:
                msg_lines.append(f"⚠️ <b>Catatan: Ada {len(low_stock_items)} barang yang stoknya kritis (&lt; 10 pcs):</b>")
                for sku, nama, sisa in low_stock_items:
                    msg_lines.append(f"• <b>{sku}</b> ({nama}: sisa <b>{sisa} pcs</b>)")
                msg_lines.append("\n<i>Mohon tim admin segera berkoordinasi untuk restok.</i>")
            else:
                msg_lines.append("✅ <b>Seluruh stok produk dalam kondisi AMAN (&gt;= 10 pcs).</b>")

            msg_lines.append("━━━━━━━━━━━━━━━━━━━━━━")
            msg_lines.append("🚀 <i>Selamat beraktivitas &amp; salam sukses pengiriman!</i>")
            return "\n".join(msg_lines)

        return self._execute_with_retry(_briefing)

    def get_evening_recap_text(self, target_date: Optional[str] = None) -> str:
        def _recap():
            nonlocal target_date
            if not target_date:
                target_date = datetime.now().strftime("%Y-%m-%d")

            try:
                ws_mutasi = self.spreadsheet.worksheet(self.TAB_MUTASI)
                mutasi_records = ws_mutasi.get_all_values()
            except gspread.exceptions.WorksheetNotFound:
                return "Belum ada transaksi di Riwayat_Mutasi."

            total_keluar = 0
            detail_keluar: Dict[str, Dict[str, Any]] = {}
            total_transaksi = 0

            for row in mutasi_records[1:]:
                if len(row) >= 6:
                    row_tanggal = row[1].strip()
                    if target_date in row_tanggal or row_tanggal.startswith(target_date):
                        tipe = row[2].strip().upper()
                        sku = row[3].strip().upper()
                        nama = row[4].strip()
                        try:
                            qty = int(row[5])
                        except ValueError:
                            qty = 0

                        if tipe == "KELUAR":
                            total_keluar += qty
                            total_transaksi += 1
                            if sku not in detail_keluar:
                                detail_keluar[sku] = {"nama": nama, "qty": 0}
                            detail_keluar[sku]["qty"] += qty

            now = datetime.now()
            tgl_display = now.strftime("%d %B %Y")

            msg_lines = [
                "🏁 <b>TUTUP BUKU OPERASIONAL GUDANG (17:00 WIB)</b>",
                "━━━━━━━━━━━━━━━━━━━━━━",
                f"📅 <i>Rekap Harian: {tgl_display}</i>\n",
                f"🔻 <b>TOTAL PAKET / BARANG KELUAR HARI INI: {total_keluar} pcs</b>",
                f"📋 <b>Total Transaksi Resi / Order: {total_transaksi} kali</b>\n"
            ]

            if detail_keluar:
                msg_lines.append("<b>Rincian Pengiriman Barang Hari Ini:</b>")
                sorted_items = sorted(detail_keluar.items(), key=lambda x: x[1]["qty"], reverse=True)
                for sku, data in sorted_items:
                    msg_lines.append(f"• <b>[{sku}]</b> {data['nama']}: <b>{data['qty']} pcs</b>")
            else:
                msg_lines.append("<i>(Tidak ada mutasi barang keluar yang tercatat hari ini)</i>")

            msg_lines.append("\n━━━━━━━━━━━━━━━━━━━━━━")
            msg_lines.append("🙏 <i>Terima kasih atas dedikasi dan kerja keras seluruh tim hari ini!</i>")
            return "\n".join(msg_lines)

        return self._execute_with_retry(_recap)

    def get_leaderboard_text(self, target_date: Optional[str] = None) -> str:
        """Menghasilkan teks ranking leaderboard seller pengirim resi terbanyak hari ini."""
        def _leaderboard():
            nonlocal target_date
            if not target_date:
                target_date = datetime.now().strftime("%Y-%m-%d")

            try:
                ws_mutasi = self.spreadsheet.worksheet(self.TAB_MUTASI)
                records = ws_mutasi.get_all_values()
            except gspread.exceptions.WorksheetNotFound:
                return "Belum ada catatan transaksi di Riwayat_Mutasi."

            seller_stats = {}
            for row in records[1:]:
                if len(row) >= 8:
                    row_tgl = row[1].strip()
                    if target_date in row_tgl or row_tgl.startswith(target_date):
                        tipe = row[2].strip().upper()
                        if tipe == "KELUAR":
                            pengirim = row[7].strip() or "Seller Lapangan"
                            try:
                                qty = int(row[5])
                            except ValueError:
                                qty = 1

                            ket = row[8].strip() if len(row) > 8 else ""
                            m_resi = re.search(r'Resi:\s*\[?([A-Za-z0-9_-]+)\]?', ket)
                            resi_key = m_resi.group(1) if m_resi else f"order_{row[0]}"

                            if pengirim not in seller_stats:
                                seller_stats[pengirim] = {"resi_set": set(), "total_pcs": 0}
                            seller_stats[pengirim]["resi_set"].add(resi_key)
                            seller_stats[pengirim]["total_pcs"] += qty

            if not seller_stats:
                tgl_str = datetime.now().strftime("%d %B %Y")
                return f"🏆 <b>LEADERBOARD SELLER GUDANG:</b>\n━━━━━━━━━━━━━━━━━━━━━━\n<i>Belum ada catatan order/resi keluar pada tanggal {tgl_str}.</i>"

            sorted_sellers = sorted(
                seller_stats.items(),
                key=lambda x: (len(x[1]["resi_set"]), x[1]["total_pcs"]),
                reverse=True
            )

            medals = ["🥇", "🥈", "🥉"]
            lines = [
                "🏆 <b>LEADERBOARD PENGIRIMAN SELLER HARI INI:</b>",
                "━━━━━━━━━━━━━━━━━━━━━━",
                f"📅 <i>Tanggal: {datetime.now().strftime('%d %B %Y')}</i>\n"
            ]

            total_semua_resi = sum(len(d["resi_set"]) for d in seller_stats.values())
            total_semua_pcs = sum(d["total_pcs"] for d in seller_stats.values())

            for idx, (seller, data) in enumerate(sorted_sellers, 1):
                badge = medals[idx-1] if idx <= 3 else f"<b>{idx}.</b>"
                resi_cnt = len(data["resi_set"])
                pcs_cnt = data["total_pcs"]
                lines.append(f"{badge} <b>{seller}</b>: <b>{resi_cnt} resi</b> <i>({pcs_cnt} pcs)</i>")

            lines.append("\n━━━━━━━━━━━━━━━━━━━━━━")
            lines.append(f"🔥 <b>Total Hari Ini:</b> {total_semua_resi} resi ({total_semua_pcs} pcs terkirim).")
            lines.append("<i>Terima kasih atas kerja keras seluruh seller & tim gudang! 🚀</i>")
            return "\n".join(lines)

        return self._execute_with_retry(_leaderboard)

    def get_leaderboard_data(self, target_date: Optional[str] = None) -> Dict[str, Any]:
        """Mengambil data leaderboard seller dalam format JSON untuk Web App."""
        def _get():
            nonlocal target_date
            if not target_date:
                target_date = datetime.now().strftime("%Y-%m-%d")

            try:
                ws_mutasi = self.spreadsheet.worksheet(self.TAB_MUTASI)
                records = ws_mutasi.get_all_values()
            except Exception:
                return {"status": "ok", "date": datetime.now().strftime("%d %B %Y"), "sellers": []}

            seller_stats = {}
            for row in records[1:]:
                if len(row) >= 8:
                    row_tgl = row[1].strip()
                    if target_date in row_tgl or row_tgl.startswith(target_date):
                        tipe = row[2].strip().upper()
                        if tipe == "KELUAR":
                            pengirim = row[7].strip() or "Seller Lapangan"
                            try:
                                qty = int(row[5])
                            except ValueError:
                                qty = 1

                            ket = row[8].strip() if len(row) > 8 else ""
                            m_resi = re.search(r'Resi:\s*\[?([A-Za-z0-9_-]+)\]?', ket)
                            resi_key = m_resi.group(1) if m_resi else f"order_{row[0]}"

                            if pengirim not in seller_stats:
                                seller_stats[pengirim] = {"resi_set": set(), "total_pcs": 0}
                            seller_stats[pengirim]["resi_set"].add(resi_key)
                            seller_stats[pengirim]["total_pcs"] += qty

            sorted_sellers = sorted(
                seller_stats.items(),
                key=lambda x: (len(x[1]["resi_set"]), x[1]["total_pcs"]),
                reverse=True
            )

            result_sellers = [
                {"name": seller, "resi": len(data["resi_set"]), "pcs": data["total_pcs"]}
                for seller, data in sorted_sellers
            ]

            return {
                "status": "ok",
                "date": datetime.now().strftime("%d %B %Y"),
                "sellers": result_sellers
            }

        return self._execute_with_retry(_get)

    def get_all_products_list(self, force_refresh: bool = False) -> List[Dict[str, Any]]:
        """Mengambil daftar seluruh SKU dan sisa stok untuk Web App & Tombol Native (Cached)."""
        now = time.time()
        if not force_refresh and self._stok_cache is not None and (now - self._stok_cache_time) < 20:
            master_map = self._get_master_map()
            return [
                {"sku": sku, "nama": master_map.get(sku, "-"), "qty": qty}
                for sku, qty in self._stok_cache.items()
            ]

        def _get():
            ws_stok = self.get_worksheet(self.TAB_STOK)
            stok_records = ws_stok.get_all_values()

            master_map = self._get_master_map()

            items = []
            new_stok_map = {}
            for r in stok_records[1:]:
                if len(r) >= 2 and r[0].strip():
                    sku = r[0].strip().upper()
                    try:
                        qty = int(r[1]) if r[1].strip() else 0
                    except ValueError:
                        qty = 0
                    new_stok_map[sku] = qty
                    items.append({
                        "sku": sku,
                        "nama": master_map.get(sku, "-"),
                        "qty": qty
                    })

            self._stok_cache = new_stok_map
            self._stok_cache_time = now
            return items

        return self._execute_with_retry(_get)

    def export_excel_bytes(self, target_date: Optional[str] = None) -> Tuple[bytes, str, int, int, int]:
        def _export():
            nonlocal target_date
            if not target_date:
                target_date = datetime.now().strftime("%Y-%m-%d")

            now = datetime.now()
            tgl_display = now.strftime("%d %B %Y")

            ws_mutasi_sheet = self.get_worksheet(self.TAB_MUTASI)
            mutasi_records = ws_mutasi_sheet.get_all_values()

            ws_stok_sheet = self.get_worksheet(self.TAB_STOK)
            stok_records = ws_stok_sheet.get_all_values()

            ws_master_sheet = self.get_worksheet(self.TAB_MASTER)
            master_records = ws_master_sheet.get_all_values()
            nama_map = {r[0].strip().upper(): r[1].strip() for r in master_records[1:] if len(r) >= 2}

            wb = openpyxl.Workbook()

            font_title = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
            fill_title = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
            font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            fill_header = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
            font_body = Font(name="Segoe UI", size=10)
            fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
            thin_border = Border(
                left=Side(style='thin', color='D0D7DE'),
                right=Side(style='thin', color='D0D7DE'),
                top=Side(style='thin', color='D0D7DE'),
                bottom=Side(style='thin', color='D0D7DE')
            )

            # TAB 1: MUTASI GUDANG
            ws1 = wb.active
            ws1.title = "Mutasi Gudang"
            ws1.views.sheetView[0].showGridLines = True

            ws1.merge_cells("A1:J1")
            ws1["A1"] = f"LAPORAN OPERASIONAL & MUTASI GUDANG ({tgl_display})"
            ws1["A1"].font = font_title
            ws1["A1"].fill = fill_title
            ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws1.row_dimensions[1].height = 38

            headers_mutasi = [
                "No", "Timestamp", "Tanggal", "Tipe", "Kode SKU",
                "Nama Barang", "Qty (pcs)", "Sisa Stok", "Pengirim", "Keterangan"
            ]
            ws1.row_dimensions[3].height = 25
            for c_idx, h in enumerate(headers_mutasi, 1):
                cell = ws1.cell(row=3, column=c_idx, value=h)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            row_pos = 4
            total_keluar = 0
            total_masuk = 0
            filtered_count = 0

            for r in mutasi_records[1:]:
                tipe = r[2] if len(r) > 2 else ""
                try:
                    qty = int(r[5]) if len(r) > 5 else 0
                except ValueError:
                    qty = 0

                if tipe.upper() == "KELUAR":
                    total_keluar += qty
                elif tipe.upper() in ["MASUK", "OPNAME"]:
                    total_masuk += qty

                filtered_count += 1
                row_vals = [
                    filtered_count,
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

                row_fill = fill_zebra if filtered_count % 2 == 0 else None
                for c_idx, val in enumerate(row_vals, 1):
                    cell = ws1.cell(row=row_pos, column=c_idx, value=val)
                    cell.font = font_body
                    cell.border = thin_border
                    if row_fill:
                        cell.fill = row_fill
                    cell.alignment = Alignment(
                        horizontal="center" if c_idx in (1, 2, 3, 4, 5, 7, 8) else "left",
                        vertical="center"
                    )
                row_pos += 1

            for col in ws1.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws1.column_dimensions[col_letter].width = max(max_len + 3, 11)

            # TAB 2: STOK BARANG TERKINI
            ws2 = wb.create_sheet(title="Stok Gudang Terkini")
            ws2.views.sheetView[0].showGridLines = True

            ws2.merge_cells("A1:E1")
            ws2["A1"] = f"MASTER SISA STOK GUDANG TERKINI ({tgl_display})"
            ws2["A1"].font = font_title
            ws2["A1"].fill = fill_title
            ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws2.row_dimensions[1].height = 38

            stok_headers = ["No", "Kode SKU", "Nama Produk Standar", "Sisa Stok (pcs)", "Status"]
            ws2.row_dimensions[3].height = 25
            for c_idx, h in enumerate(stok_headers, 1):
                cell = ws2.cell(row=3, column=c_idx, value=h)
                cell.font = font_header
                cell.fill = fill_header
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
                    status = "KRITIS (< 10)"
                    status_fill = warning_fill
                else:
                    status = "AMAN"
                    status_fill = ok_fill

                r_vals = [idx, sku, nama, sisa, status]
                for c_idx, val in enumerate(r_vals, 1):
                    cell = ws2.cell(row=r2_pos, column=c_idx, value=val)
                    cell.font = font_body
                    cell.border = thin_border
                    if c_idx == 5:
                        cell.fill = status_fill
                        cell.font = Font(name="Segoe UI", size=10, bold=True)
                    cell.alignment = Alignment(
                        horizontal="center" if c_idx in (1, 2, 4, 5) else "left",
                        vertical="center"
                    )
                r2_pos += 1

            for col in ws2.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            data_bytes = buf.getvalue()
            filename = f"Laporan_Gudang_{now.strftime('%d%b%Y')}.xlsx"

            return data_bytes, filename, total_keluar, total_masuk, filtered_count

        return self._execute_with_retry(_export)

    def get_laporan_mutasi(self, target_date: Optional[str] = None) -> str:
        def _laporan():
            nonlocal target_date
            if not target_date:
                target_date = datetime.now().strftime("%Y-%m-%d")

            try:
                ws_mutasi = self.get_worksheet(self.TAB_MUTASI)
                mutasi_records = ws_mutasi.get_all_values()
            except gspread.exceptions.WorksheetNotFound:
                return "Belum ada data di tab Riwayat_Mutasi."

            total_keluar = 0
            total_masuk = 0
            detail_keluar: Dict[str, Dict[str, Any]] = {}
            detail_masuk: Dict[str, Dict[str, Any]] = {}

            for row in mutasi_records[1:]:
                if len(row) >= 6:
                    row_tanggal = row[1].strip()
                    if target_date in row_tanggal or row_tanggal.startswith(target_date):
                        tipe = row[2].strip().upper()
                        sku = row[3].strip().upper()
                        nama = row[4].strip()
                        try:
                            qty = int(row[5])
                        except ValueError:
                            qty = 0

                        if tipe == "KELUAR":
                            total_keluar += qty
                            if sku not in detail_keluar:
                                detail_keluar[sku] = {"nama": nama, "qty": 0}
                            detail_keluar[sku]["qty"] += qty
                        elif tipe in ["MASUK", "OPNAME"]:
                            total_masuk += qty
                            if sku not in detail_masuk:
                                detail_masuk[sku] = {"nama": nama, "qty": 0}
                            detail_masuk[sku]["qty"] += qty

            ws_stok = self.get_worksheet(self.TAB_STOK)
            stok_records = ws_stok.get_all_values()
            low_stock_items = []
            for row in stok_records[1:]:
                if len(row) >= 2 and row[0].strip():
                    sku = row[0].strip().upper()
                    try:
                        sisa = int(row[1])
                        if sisa < 10:
                            low_stock_items.append((sku, sisa))
                    except ValueError:
                        pass

            tgl_tampil = target_date
            try:
                tgl_tampil = datetime.strptime(target_date, "%Y-%m-%d").strftime("%d %B %Y")
            except Exception:
                pass

            report_lines = [
                f"📊 <b>REKAP LAPORAN GUDANG ({tgl_tampil})</b>",
                "━━━━━━━━━━━━━━━━━━━━━━",
                f"🔻 <b>TOTAL BARANG KELUAR: {total_keluar} pcs</b>"
            ]

            if detail_keluar:
                for sku, data in detail_keluar.items():
                    report_lines.append(f"  • [{sku}] {data['nama']}: <b>{data['qty']} pcs</b>")
            else:
                report_lines.append("  <i>(Belum ada barang keluar pada tanggal ini)</i>")

            report_lines.append("")
            report_lines.append(f"🔺 <b>TOTAL BARANG MASUK (RESTOK / OPNAME): {total_masuk} pcs</b>")
            if detail_masuk:
                for sku, data in detail_masuk.items():
                    report_lines.append(f"  • [{sku}] {data['nama']}: <b>+{data['qty']} pcs</b>")
            else:
                report_lines.append("  <i>(Belum ada restok masuk pada tanggal ini)</i>")

            if low_stock_items:
                report_lines.append("")
                report_lines.append("⚠️ <b>PERINGATAN STOK MENIPIS (&lt; 10 pcs):</b>")
                for sku, sisa in low_stock_items[:5]:
                    report_lines.append(f"  • [{sku}]: sisa <b>{sisa} pcs</b>")

            report_lines.append("━━━━━━━━━━━━━━━━━━━━━━")
            report_lines.append("<i>💡 Ketik /export untuk download laporan Excel lengkap.</i>")

            return "\n".join(report_lines)

        return self._execute_with_retry(_laporan)

    def get_stok_overview(self, search_query: Optional[str] = None) -> str:
        items = self.get_all_products_list()
        if not items:
            return "Belum ada data stok barang."

        lines = ["📦 <b>STATUS SISA STOK GUDANG:</b>", "━━━━━━━━━━━━━━━━━━━━━━"]
        count = 0
        for item in items:
            sku = item["sku"]
            nama = item["nama"]
            qty = item["qty"]
            if search_query:
                sq = search_query.lower()
                if sq not in sku.lower() and sq not in nama.lower():
                    continue
            lines.append(f"• <b>[{sku}]</b> {nama}: <b>{qty} pcs</b>")
            count += 1

        if count == 0 and search_query:
            return f"Tidak ditemukan barang dengan kata kunci '{search_query}'."

        lines.append("━━━━━━━━━━━━━━━━━━━━━━")
        return "\n".join(lines)

    def append_kamus_alias(self, nama_asli_seller: str, kode_sku: str) -> bool:
        def _append():
            ws = self.get_worksheet(self.TAB_KAMUS)
            clean_nama = nama_asli_seller.strip()
            clean_sku = kode_sku.strip().upper()
            ws.append_row([clean_nama, clean_sku])
            self._kamus_cache = None
            logger.info(f"Data baru diappend ke Kamus_Alias: ['{clean_nama}', '{clean_sku}']")
            return True

        return self._execute_with_retry(_append)


sheets_manager = GoogleSheetsManager(GOOGLE_SERVICE_ACCOUNT_FILE, SPREADSHEET_ID_OR_URL)


# Helper Broadcast Notifikasi Restok Otomatis ke Seluruh Grup Aktif
def broadcast_restok_to_groups(kode_sku: str, nama_barang: str, qty_masuk: int, new_qty: int, admin_name: str = "Admin"):
    """Menyiarkan pesan restok secara otomatis ke seluruh grup Telegram aktif."""
    active_group_ids = get_active_group_ids()
    if not active_group_ids:
        logger.info("Tidak ada grup aktif untuk broadcast restok.")
        return

    msg = (
        f"📢 <b>INFO RESTOK GUDANG:</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"Produk <b>[{kode_sku}] {nama_barang}</b> baru saja ditambah <b>+{qty_masuk} pcs</b> oleh {admin_name}!\n"
        f"📊 <b>Total Stok Siap Kirim:</b> <b>{new_qty} pcs</b>\n\n"
        f"<i>Stok siap diproses kembali untuk pengiriman order &amp; resi seller!</i> 🚀\n"
        f"━━━━━━━━━━━━━━━━━━━━━━"
    )
    for gid in active_group_ids:
        try:
            telegram_client.send_message(gid, msg, parse_mode="HTML")
            time.sleep(0.5)
        except Exception as e:
            logger.warning(f"Gagal broadcast restok ke grup {gid}: {e}")


# =============================================================================
# 5. HELPER: GEMINI AI NORMALIZER (Multi-Item & Konversi Kemasan)
# =============================================================================
class GeminiNormalizer:
    def __init__(self, api_key: str):
        self.api_key = api_key
        if api_key:
            genai.configure(api_key=api_key)

    def normalize(
        self,
        pesan_seller: str,
        daftar_master_sku: str,
        media_bytes: Optional[bytes] = None,
        media_mime_type: Optional[str] = None
    ) -> Optional[Dict[str, Any]]:
        if not self.api_key:
            logger.error("Gemini API Key belum dikonfigurasi.")
            return None

        system_instruction = (
            "Kamu adalah sistem AI admin logistik gudang ahli di Indonesia. "
            "Tugasmu menormalisasi pesanan/resi barang masuk ke format JSON standar.\n\n"
            f"DAFTAR MASTER SKU RESMI:\n{daftar_master_sku}\n\n"
            "ATURAN KONVERSI SATUAN KEMASAN KE PCS:\n"
            "- 1 lusin = 12 pcs (misal: '2 lusin Bio Binahong' -> qty = 24 pcs)\n"
            "- 1 kodi = 20 pcs\n"
            "- 1 box / kotak G-MAX (MHPD 027) = 10 sachet/pcs (misal: '1 box G-MAX' -> qty = 10 pcs, '2 box' -> qty = 20 pcs)\n"
            "- 1 box / kotak V-MAX (MHPD 028) = 10 sachet/pcs\n"
            "- 1 box / kotak umum = 10 pcs\n"
            "- 1 dus / karton = 24 pcs\n"
            "- Satuan botol, kapsul, tube, sachet, pcs, buah, pack = dihitung per satuan (1 unit = 1 pcs)\n\n"
            "ATURAN EKSTRAKSI RESI & MULTI-ITEM:\n"
            "1. Ekstrak 'no_resi' dari teks/dokumen resi jika ada (misal nomor resi Anteraja, SPX, J&T, dsb). Jika tidak ada resi, isi string kosong \"\".\n"
            "2. Ekstrak SEMUA barang dalam pesanan ke dalam list 'items'. Jika pesanan berisi lebih dari 1 macam barang (orderan campuran/bundle), masukkan seluruh barang ke dalam array 'items'.\n"
            "3. Cocokkan setiap barang ke SKU terdekat di Master SKU. Jika barang tidak ada di Master SKU, set 'kode_sku': 'UNKNOWN' dan 'nama_standar': 'UNKNOWN'.\n"
            "4. Nilai 'qty' HARUS bertipe integer dalam satuan PCS setelah dikonversi kemasannya.\n"
            "5. Format output HARUS JSON valid dengan struktur persis seperti ini:\n"
            "{\n"
            "  \"no_resi\": \"11004683469987\",\n"
            "  \"items\": [\n"
            "    {\n"
            "      \"kode_sku\": \"MHPD 021\",\n"
            "      \"nama_standar\": \"DEPHERO KAPSUL\",\n"
            "      \"qty\": 1,\n"
            "      \"satuan_asli\": \"pcs\",\n"
            "      \"nama_asli_seller\": \"dephero 1 pcs\"\n"
            "    }\n"
            "  ]\n"
            "}\n"
            "Kembalikan HANYA JSON tanpa teks pengantar atau markdown."
        )

        user_prompt = f"Pesan/Resi seller: {pesan_seller}" if pesan_seller else "Ekstrak nomor resi, seluruh barang pesanan, dan kuantitas PCS dari resi ini."

        prompt_parts: List[Dict[str, Any]] = [{"text": user_prompt}]
        if media_bytes:
            b64_data = base64.b64encode(media_bytes).decode("utf-8")
            prompt_parts.append({
                "inline_data": {
                    "mime_type": media_mime_type or "application/pdf",
                    "data": b64_data
                }
            })

        candidate_models = ["gemini-flash-lite-latest", "gemini-3.1-flash-lite", "gemini-flash-latest"]
        raw_text = None

        for candidate in candidate_models:
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{candidate}:generateContent?key={self.api_key}"
                payload = {
                    "system_instruction": {
                        "parts": [{"text": system_instruction}]
                    },
                    "contents": [
                        {"parts": prompt_parts}
                    ],
                    "generationConfig": {
                        "responseMimeType": "application/json",
                        "temperature": 0.1
                    }
                }
                resp = requests.post(url, json=payload, timeout=25)
                if resp.status_code == 200:
                    data = resp.json()
                    candidates = data.get("candidates", [])
                    if candidates:
                        parts = candidates[0].get("content", {}).get("parts", [])
                        if parts:
                            raw_text = parts[0].get("text", "").strip()
                            logger.info(f"Gemini AI ({candidate}) Berhasil via REST API")
                            break
                else:
                    logger.warning(f"Gemini REST ({candidate}) HTTP {resp.status_code}: {resp.text[:100]}")
            except Exception as rest_err:
                logger.warning(f"Gemini REST error ({candidate}): {rest_err}")

        if not raw_text:
            for candidate in candidate_models:
                try:
                    model = genai.GenerativeModel(
                        model_name=candidate,
                        system_instruction=system_instruction,
                        generation_config={
                            "temperature": 0.1,
                            "response_mime_type": "application/json"
                        }
                    )
                    response = model.generate_content(user_prompt)
                    raw_text = response.text.strip()
                    logger.info(f"Gemini AI ({candidate}) Berhasil via SDK")
                    break
                except Exception as sdk_err:
                    logger.warning(f"Gemini SDK error ({candidate}): {sdk_err}")

        if not raw_text:
            logger.error("Semua kandidat model Gemini Flash tidak dapat merespons.")
            return None

        try:
            cleaned_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.MULTILINE)
            cleaned_text = re.sub(r"\s*```$", "", cleaned_text, flags=re.MULTILINE).strip()
            result = json.loads(cleaned_text)

            if isinstance(result, list) and len(result) > 0:
                result = {"no_resi": "", "items": result}

            if "items" not in result:
                if "kode_sku" in result:
                    result = {
                        "no_resi": result.get("no_resi", ""),
                        "items": [result]
                    }
                else:
                    result = {"no_resi": "", "items": []}

            valid_items = []
            for it in result.get("items", []):
                if not isinstance(it, dict):
                    continue
                try:
                    it["qty"] = max(1, int(it.get("qty", 1)))
                except (ValueError, TypeError):
                    it["qty"] = 1
                valid_items.append(it)

            result["items"] = valid_items
            return result

        except Exception as e:
            logger.error(f"Gagal mem-parsing JSON dari Gemini: {e}. Raw: {raw_text}")
            return None


gemini_normalizer = GeminiNormalizer(GEMINI_API_KEY)


# =============================================================================
# 6. SCHEDULER: BACKGROUND THREAD AUTO-BRIEFING PAGI (08:00) & REKAP SORE (17:00)
# =============================================================================
_scheduler_started = False
_scheduler_lock = threading.Lock()


def run_scheduler_loop():
    logger.info("Background Scheduler Auto-Briefing (08:00 WIB) & Rekap Sore (17:00 WIB) telah aktif.")
    last_morning_date = ""
    last_evening_date = ""

    while True:
        try:
            now = datetime.now()
            today_str = now.strftime("%Y-%m-%d")
            hour = now.hour
            minute = now.minute

            if hour == 8 and minute == 0 and last_morning_date != today_str:
                active_group_ids = get_active_group_ids()
                if active_group_ids:
                    logger.info(f"Memicu Auto-Briefing Pagi ke {len(active_group_ids)} grup...")
                    briefing_msg = sheets_manager.get_morning_briefing_text()
                    for gid in active_group_ids:
                        telegram_client.send_message(gid, briefing_msg, parse_mode="HTML")
                        time.sleep(1)
                last_morning_date = today_str

            if hour == 17 and minute == 0 and last_evening_date != today_str:
                active_group_ids = get_active_group_ids()
                if active_group_ids:
                    logger.info(f"Memicu Rekap Sore Tutup Buku ke {len(active_group_ids)} grup...")
                    recap_msg = sheets_manager.get_evening_recap_text()
                    for gid in active_group_ids:
                        telegram_client.send_message(gid, recap_msg, parse_mode="HTML")
                        time.sleep(1)
                last_evening_date = today_str

        except Exception as e:
            logger.error(f"Error pada Scheduler Loop: {e}")

        time.sleep(30)


def start_background_scheduler():
    global _scheduler_started
    with _scheduler_lock:
        if not _scheduler_started:
            t = threading.Thread(target=run_scheduler_loop, daemon=True)
            t.start()
            _scheduler_started = True


# =============================================================================
# 7. BUSINESS LOGIC HANDLER: MEMPROSES PESAN TELEGRAM (100% NATIVE TELEGRAM UI)
# =============================================================================
def get_admin_dashboard_markup(is_superadmin_user: bool = False) -> dict:
    """Membuat Inline Keyboard untuk Panel Admin / Superadmin."""
    keyboard = [
        [
            {"text": "📦 Katalog Sisa Stok", "callback_data": "btn_stok_p1"},
            {"text": "📥 Restok Cepat (1-Klik)", "callback_data": "btn_restok_menu"}
        ],
        [
            {"text": "🔄 Stock Opname Cepat", "callback_data": "btn_opname_menu"},
            {"text": "➕ Tambah Master SKU", "callback_data": "btn_tambah_sku"}
        ],
        [
            {"text": "🏆 Leaderboard Seller", "callback_data": "btn_leaderboard"},
            {"text": "📊 Unduh Laporan Excel", "callback_data": "btn_export"}
        ]
    ]
    if is_superadmin_user:
        keyboard.append([
            {"text": "👑 Kelola Admin & Keamanan", "callback_data": "btn_superadmin_menu"}
        ])
    return {"inline_keyboard": keyboard}


def get_superadmin_menu_markup() -> Tuple[str, dict]:
    """Menampilkan menu khusus Superadmin untuk kelola staf admin dan keamanan PIN."""
    current_pin = get_current_pin()
    data = load_authorized_admins()
    admins_dict = data.get("admins", {})
    total_admins = len(admins_dict) or len(data.get("whitelist_ids", []))

    text = (
        "👑 <b>PANEL SUPERADMIN GUDANG (OWNER)</b> 🛡️\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        f"🔑 <b>PIN Masuk Gudang Saat Ini:</b> <code>{current_pin}</code>\n"
        f"👥 <b>Total Staf Admin Terdaftar:</b> <b>{total_admins} orang</b>\n\n"
        "Sebagai Superadmin, Anda memiliki wewenang penuh untuk:\n"
        "• Mendaftarkan staf admin baru agar bisa akses japri\n"
        "• Mencabut hak akses admin staf\n"
        "• Mengganti PIN masuk bot kapan saja\n"
        "• Reset / kunci seluruh sesi login staf\n\n"
        "Silakan pilih menu pengaturan di bawah ini:"
    )
    buttons = [
        [
            {"text": "👥 Daftar Seluruh Admin", "callback_data": "btn_sa_list"},
            {"text": "➕ Tambah Admin Baru", "callback_data": "btn_sa_add"}
        ],
        [
            {"text": "❌ Cabut Akses Admin", "callback_data": "btn_sa_del_menu"},
            {"text": "🔑 Ganti PIN Gudang", "callback_data": "btn_sa_change_pin"}
        ],
        [
            {"text": "🔒 Kunci / Reset Semua Sesi", "callback_data": "btn_sa_reset_sessions"}
        ],
        [
            {"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}
        ]
    ]
    return text, {"inline_keyboard": buttons}


def get_sa_delete_admin_markup() -> Tuple[str, dict]:
    """Menampilkan pilihan admin yang bisa dicabut hak aksesnya dengan 1 klik."""
    data = load_authorized_admins()
    admins_dict = data.get("admins", {})
    buttons = []
    text_lines = [
        "❌ <b>CABUT HAK AKSES ADMIN GUDANG:</b>",
        "━━━━━━━━━━━━━━━━━━━━━━",
        "Pilih salah satu staf admin di bawah untuk mencabut aksesnya langsung:",
        ""
    ]
    has_removable = False
    for str_uid, info in admins_dict.items():
        try:
            uid = int(str_uid)
        except ValueError:
            continue
        if is_superadmin(uid):
            continue
        has_removable = True
        nama = info.get("name", f"Admin {uid}")
        buttons.append([
            {"text": f"❌ Hapus: {nama} ({uid})", "callback_data": f"cb_sa_del_{uid}"}
        ])

    if not has_removable:
        text_lines.append("<i>Belum ada staf admin tambahan yang dapat dihapus.</i>\n")
        text_lines.append("<i>Atau gunakan perintah manual: <code>/hapus_admin [ID_Telegram]</code></i>")
    else:
        text_lines.append("<i>💡 Atau gunakan perintah manual: <code>/hapus_admin [ID_Telegram]</code></i>")

    buttons.append([{"text": "🔙 Kembali ke Menu Superadmin", "callback_data": "btn_superadmin_menu"}])
    return "\n".join(text_lines), {"inline_keyboard": buttons}


def get_stok_page_content(page: int = 1) -> Tuple[str, dict]:
    """Menghasilkan teks katalog stok berhalaman (pagination) langsung di chat Telegram."""
    items = sheets_manager.get_all_products_list()
    page_size = 7
    total_pages = max(1, (len(items) + page_size - 1) // page_size)
    page = max(1, min(page, total_pages))

    start_idx = (page - 1) * page_size
    page_items = items[start_idx:start_idx + page_size]

    lines = [
        f"📦 <b>KATALOG SISA STOK GUDANG (Hal {page}/{total_pages})</b>",
        "━━━━━━━━━━━━━━━━━━━━━━"
    ]

    for it in page_items:
        sku = it["sku"]
        nama = it["nama"]
        qty = it["qty"]
        if qty == 0:
            badge = "🔴 <b>HABIS (0 pcs)</b>"
        elif qty < 10:
            badge = f"🟡 <b>KRITIS ({qty} pcs)</b>"
        else:
            badge = f"🟢 <b>AMAN ({qty} pcs)</b>"
        lines.append(f"• <b>[{sku}]</b> {nama}\n  └ Status: {badge}")

    lines.append("━━━━━━━━━━━━━━━━━━━━━━")
    lines.append(f"<i>Total Produk: {len(items)} SKU | Diperbarui real-time.</i>")

    nav_row = []
    if page > 1:
        nav_row.append({"text": "⬅️ Prev", "callback_data": f"btn_stok_p{page - 1}"})
    nav_row.append({"text": f"📄 {page}/{total_pages}", "callback_data": "btn_noop"})
    if page < total_pages:
        nav_row.append({"text": "Next ➡️", "callback_data": f"btn_stok_p{page + 1}"})

    keyboard = {
        "inline_keyboard": [
            nav_row,
            [{"text": "📥 Restok Produk", "callback_data": "btn_restok_menu"}],
            [{"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}]
        ]
    }
    return "\n".join(lines), keyboard


def get_restok_menu_markup(page: int = 1) -> Tuple[str, dict]:
    """Menampilkan pilihan produk untuk restok cepat langsung di Telegram dengan nama produk dan sisa stok."""
    items = sheets_manager.get_all_products_list()
    page_size = 8
    total_pages = max(1, (len(items) + page_size - 1) // page_size)
    page = max(1, min(page, total_pages))

    start_idx = (page - 1) * page_size
    page_items = items[start_idx:start_idx + page_size]

    text = (
        f"📥 <b>RESTOK CEPAT (1-KLIK DI TELEGRAM)</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"Pilih salah satu produk di bawah yang ingin ditambah stoknya (Hal {page}/{total_pages}):\n\n"
        f"<i>💡 Atau gunakan perintah manual:\n<code>/restok [Kode_SKU / Nama] [Qty]</code></i>"
    )

    buttons = []
    for i in range(0, len(page_items), 2):
        row = []
        for it in page_items[i:i+2]:
            nama_clean = (it["nama"] or "").strip().title()
            if not nama_clean or nama_clean == "-":
                nama_clean = it["sku"]
            nama_clean = (
                nama_clean.replace("Violla Sweet Cream", "Violla Sweet")
                          .replace("Xt Spray 10 Ml", "XT Spray 10ml")
            )
            if len(nama_clean) > 15:
                nama_clean = nama_clean[:14] + "…"
            button_label = f"📦 {nama_clean} ({it['qty']})"
            row.append({
                "text": button_label,
                "callback_data": f"cb_p_restok_{it['sku']}"
            })
        buttons.append(row)

    nav_row = []
    if page > 1:
        nav_row.append({"text": "⬅️ Prev", "callback_data": f"btn_restok_p{page - 1}"})
    nav_row.append({"text": f"📄 Hal {page}/{total_pages}", "callback_data": "btn_noop"})
    if page < total_pages:
        nav_row.append({"text": "Next ➡️", "callback_data": f"btn_restok_p{page + 1}"})
    buttons.append(nav_row)

    buttons.append([{"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}])
    return text, {"inline_keyboard": buttons}


def get_restok_qty_markup(sku: str) -> Tuple[str, dict]:
    """Menampilkan tombol kuantitas penambahan stok untuk produk tertentu."""
    nama = sheets_manager.get_nama_standar_by_sku(sku)
    cur_qty = sheets_manager.get_stok_by_sku(sku)
    text = (
        f"📥 <b>PILIH JUMLAH RESTOK PRODUK:</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📦 <b>Nama Produk:</b> {nama}\n"
        f"🏷️ <b>SKU:</b> <code>[{sku}]</code>\n"
        f"📊 <b>Stok Gudang Saat Ini:</b> <b>{cur_qty} pcs</b>\n\n"
        f"Klik salah satu jumlah penambahan stok di bawah:"
    )
    buttons = [
        [
            {"text": "➕ 10 pcs", "callback_data": f"do_restok_{sku}_10"},
            {"text": "➕ 25 pcs", "callback_data": f"do_restok_{sku}_25"},
            {"text": "➕ 50 pcs", "callback_data": f"do_restok_{sku}_50"}
        ],
        [
            {"text": "➕ 100 pcs", "callback_data": f"do_restok_{sku}_100"},
            {"text": "➕ 200 pcs", "callback_data": f"do_restok_{sku}_200"},
            {"text": "➕ 500 pcs", "callback_data": f"do_restok_{sku}_500"}
        ],
        [
            {"text": "🔙 Batal / Pilih Produk Lain", "callback_data": "btn_restok_menu"}
        ]
    ]
    return text, {"inline_keyboard": buttons}


def get_opname_menu_markup(page: int = 1) -> Tuple[str, dict]:
    """Menampilkan pilihan produk untuk stock opname fisik langsung di Telegram dengan nama produk dan sisa stok."""
    items = sheets_manager.get_all_products_list()
    page_size = 8
    total_pages = max(1, (len(items) + page_size - 1) // page_size)
    page = max(1, min(page, total_pages))

    start_idx = (page - 1) * page_size
    page_items = items[start_idx:start_idx + page_size]

    text = (
        f"🔄 <b>STOCK OPNAME (KOREKSI FISIK GUDANG)</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"Pilih salah satu produk yang ingin disesuaikan stok fisiknya (Hal {page}/{total_pages}):\n\n"
        f"<i>💡 Atau gunakan perintah manual:\n<code>/opname [Kode_SKU] [Selisih] [Alasan]</code></i>"
    )

    buttons = []
    for i in range(0, len(page_items), 2):
        row = []
        for it in page_items[i:i+2]:
            nama_clean = (it["nama"] or "").strip().title()
            if not nama_clean or nama_clean == "-":
                nama_clean = it["sku"]
            nama_clean = (
                nama_clean.replace("Violla Sweet Cream", "Violla Sweet")
                          .replace("Xt Spray 10 Ml", "XT Spray 10ml")
            )
            if len(nama_clean) > 15:
                nama_clean = nama_clean[:14] + "…"
            button_label = f"🏷️ {nama_clean} ({it['qty']})"
            row.append({
                "text": button_label,
                "callback_data": f"cb_p_opname_{it['sku']}"
            })
        buttons.append(row)

    nav_row = []
    if page > 1:
        nav_row.append({"text": "⬅️ Prev", "callback_data": f"btn_opname_p{page - 1}"})
    nav_row.append({"text": f"📄 Hal {page}/{total_pages}", "callback_data": "btn_noop"})
    if page < total_pages:
        nav_row.append({"text": "Next ➡️", "callback_data": f"btn_opname_p{page + 1}"})
    buttons.append(nav_row)

    buttons.append([{"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}])
    return text, {"inline_keyboard": buttons}


def get_opname_qty_markup(sku: str) -> Tuple[str, dict]:
    """Menampilkan pilihan selisih koreksi stock opname."""
    nama = sheets_manager.get_nama_standar_by_sku(sku)
    cur_qty = sheets_manager.get_stok_by_sku(sku)
    text = (
        f"🔄 <b>PILIH PENYESUAIAN STOCK OPNAME:</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📦 <b>Nama Produk:</b> {nama}\n"
        f"🏷️ <b>SKU:</b> <code>[{sku}]</code>\n"
        f"📊 <b>Stok Gudang Saat Ini:</b> <b>{cur_qty} pcs</b>\n\n"
        f"Pilih salah satu jenis koreksi fisik di bawah ini:"
    )
    buttons = [
        [
            {"text": "💥 -1 pcs (Pecah/Rusak)", "callback_data": f"do_opname_{sku}_-1_pecah"},
            {"text": "💥 -2 pcs (Pecah/Rusak)", "callback_data": f"do_opname_{sku}_-2_pecah"}
        ],
        [
            {"text": "📉 -5 pcs (Selisih Kurang)", "callback_data": f"do_opname_{sku}_-5_selisih"},
            {"text": "📈 +5 pcs (Fisik Lebih)", "callback_data": f"do_opname_{sku}_5_lebih"}
        ],
        [
            {"text": "🔙 Batal / Pilih Produk Lain", "callback_data": "btn_opname_menu"}
        ]
    ]
    return text, {"inline_keyboard": buttons}


def handle_incoming_callback_query(callback_query: dict) -> Tuple[bool, str]:
    """Menangani aksi klik tombol dari inline keyboard (100% Native Telegram)."""
    cb_id = callback_query.get("id")
    data = callback_query.get("data", "")
    message = callback_query.get("message", {})
    chat = message.get("chat", {})
    chat_id = chat.get("id")
    message_id = message.get("message_id")
    from_user = callback_query.get("from", {})
    user_id = from_user.get("id")
    user_name = (from_user.get("first_name", "") + " " + from_user.get("last_name", "")).strip() or "Admin"

    # Verifikasi Otorisasi: Tombol admin hanya bisa diklik oleh user yang sudah terotorisasi / login PIN
    if not is_user_authorized(user_id):
        telegram_client.answer_callback_query(
            cb_id,
            text="🔒 Akses Ditolak: Anda belum login! Silakan ketik PIN Admin di chat terlebih dahulu.",
            show_alert=True
        )
        return False, "callback_unauthorized"

    telegram_client.answer_callback_query(cb_id)

    # 1. Menu Utama
    if data == "btn_menu_utama":
        is_sa = is_superadmin(user_id)
        role_label = "Superadmin (Owner) 👑" if is_sa else "Admin Gudang 🚀"
        dashboard_text = (
            f"🎛️ <b>PANEL KENDALI GUDANG TIRA ({role_label})</b> 👋\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"Halo <b>{user_name}</b>! Seluruh kontrol dan pencatatan gudang dapat dilakukan langsung di Telegram tanpa perlu membuka web:\n\n"
            f"Silakan pilih menu aksi cepat di bawah ini:"
        )
        telegram_client.edit_message_text(
            chat_id, message_id, dashboard_text, reply_markup=get_admin_dashboard_markup(is_superadmin_user=is_sa)
        )
        return True, "cb_menu_utama"

    # 2. Katalog Sisa Stok (Pagination)
    elif data.startswith("btn_stok_p"):
        try:
            page_num = int(data.replace("btn_stok_p", ""))
        except ValueError:
            page_num = 1
        stok_text, stok_markup = get_stok_page_content(page_num)
        telegram_client.edit_message_text(chat_id, message_id, stok_text, reply_markup=stok_markup)
        return True, "cb_stok_page"

    # 3. Menu Pilihan Produk Restok (dengan Pagination)
    elif data == "btn_restok_menu" or data.startswith("btn_restok_p"):
        page_num = 1
        if data.startswith("btn_restok_p"):
            try:
                page_num = int(data.replace("btn_restok_p", ""))
            except ValueError:
                page_num = 1
        menu_text, menu_markup = get_restok_menu_markup(page=page_num)
        telegram_client.edit_message_text(chat_id, message_id, menu_text, reply_markup=menu_markup)
        return True, "cb_restok_menu"

    # 4. Pilih Jumlah Restok untuk SKU Tertentu
    elif data.startswith("cb_p_restok_"):
        sku = data.replace("cb_p_restok_", "").strip()
        qty_text, qty_markup = get_restok_qty_markup(sku)
        telegram_client.edit_message_text(chat_id, message_id, qty_text, reply_markup=qty_markup)
        return True, "cb_p_restok"

    # 5. Eksekusi Restok Cepat
    elif data.startswith("do_restok_"):
        parts = data.split("_")
        sku = parts[2]
        qty = int(parts[3])
        nama_resmi = sheets_manager.get_nama_standar_by_sku(sku)
        cur_qty, new_qty, _, _ = sheets_manager.update_or_create_stok(sku, qty, action="MASUK")
        sheets_manager.record_mutasi(
            tipe="MASUK",
            kode_sku=sku,
            nama_barang=nama_resmi,
            qty=qty,
            sisa_stok=new_qty,
            pengirim=user_name,
            keterangan="Restok 1-Klik Telegram"
        )
        # Broadcast ke seluruh grup aktif
        broadcast_restok_to_groups(sku, nama_resmi, qty, new_qty, admin_name=user_name)

        success_text = (
            f"✅ <b>RESTOK 1-KLIK BERHASIL!</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🏷️ <b>SKU:</b> <code>[{sku}]</code>\n"
            f"📦 <b>Produk:</b> {nama_resmi}\n"
            f"🔺 <b>Qty Masuk:</b> +{qty} pcs\n"
            f"📊 <b>Total Stok Sekarang:</b> <b>{new_qty} pcs</b>\n\n"
            f"📢 <i>Pengumuman restok otomatis telah disiarkan ke seluruh grup logistik!</i> 🚀"
        )
        success_markup = {
            "inline_keyboard": [
                [{"text": "📥 Restok Produk Lain", "callback_data": "btn_restok_menu"}],
                [{"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}]
            ]
        }
        telegram_client.edit_message_text(chat_id, message_id, success_text, reply_markup=success_markup)
        return True, "cb_restok_executed"

    # 6. Menu Pilihan Produk Stock Opname (dengan Pagination)
    elif data == "btn_opname_menu" or data.startswith("btn_opname_p"):
        page_num = 1
        if data.startswith("btn_opname_p"):
            try:
                page_num = int(data.replace("btn_opname_p", ""))
            except ValueError:
                page_num = 1
        menu_text, menu_markup = get_opname_menu_markup(page=page_num)
        telegram_client.edit_message_text(chat_id, message_id, menu_text, reply_markup=menu_markup)
        return True, "cb_opname_menu"

    # 7. Pilih Penyesuaian Opname untuk SKU Tertentu
    elif data.startswith("cb_p_opname_"):
        sku = data.replace("cb_p_opname_", "").strip()
        qty_text, qty_markup = get_opname_qty_markup(sku)
        telegram_client.edit_message_text(chat_id, message_id, qty_text, reply_markup=qty_markup)
        return True, "cb_p_opname"

    # 8. Eksekusi Stock Opname Cepat
    elif data.startswith("do_opname_"):
        parts = data.split("_")
        sku = parts[2]
        diff_qty = int(parts[3])
        alasan_code = parts[4] if len(parts) > 4 else "koreksi"
        alasan_map = {
            "pecah": "Barang Pecah/Rusak saat Bongkar Muat",
            "selisih": "Selisih Kurang saat Hitung Fisik Gudang",
            "lebih": "Fisik Lebih Banyak saat Hitung Fisik"
        }
        alasan_str = alasan_map.get(alasan_code, "Koreksi Fisik Mingguan")

        cur_qty, new_qty, success, nama_barang = sheets_manager.perform_stock_opname(
            sku, diff_qty, alasan_str, admin_name=user_name
        )

        tanda = "+" if diff_qty >= 0 else ""
        res_text = (
            f"✅ <b>KOREKSI STOCK OPNAME BERHASIL!</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🏷️ <b>SKU:</b> <code>[{sku}]</code>\n"
            f"📦 <b>Produk:</b> {nama_barang}\n"
            f"📉 <b>Penyesuaian:</b> <b>{tanda}{diff_qty} pcs</b>\n"
            f"📊 <b>Stok:</b> {cur_qty} pcs ➡️ <b>{new_qty} pcs</b>\n"
            f"📝 <b>Alasan:</b> <i>{alasan_str}</i>\n"
            f"👤 <b>Petugas:</b> {user_name}\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"<i>Tercatat resmi di Tab Riwayat_Mutasi sebagai OPNAME.</i>"
        )
        res_markup = {
            "inline_keyboard": [
                [{"text": "🔄 Opname Produk Lain", "callback_data": "btn_opname_menu"}],
                [{"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}]
            ]
        }
        telegram_client.edit_message_text(chat_id, message_id, res_text, reply_markup=res_markup)
        return True, "cb_opname_executed"

    # 9. Panduan Tambah Master SKU
    elif data == "btn_tambah_sku":
        msg = (
            "➕ <b>PANDUAN TAMBAH MASTER SKU BARU:</b>\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Ketik perintah di chat ini dengan format:\n"
            "<code>/tambah_sku [Kode_SKU] | [Nama_Standar] | [Stok_Awal]</code>\n\n"
            "Contoh (klik untuk salin):\n"
            "<code>/tambah_sku MHPD 029 | KOPI MANTAP | 100</code>"
        )
        markup = {"inline_keyboard": [[{"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}]]}
        telegram_client.edit_message_text(chat_id, message_id, msg, reply_markup=markup)
        return True, "cb_tambah_sku_sent"

    # 10. Leaderboard Seller Langsung di Chat
    elif data == "btn_leaderboard":
        lb_text = sheets_manager.get_leaderboard_text()
        markup = {"inline_keyboard": [[{"text": "🔙 Kembali ke Menu Utama", "callback_data": "btn_menu_utama"}]]}
        telegram_client.edit_message_text(chat_id, message_id, lb_text, reply_markup=markup)
        return True, "cb_leaderboard_sent"

    # 11. Unduh File Excel Langsung ke Chat
    elif data == "btn_export":
        excel_bytes, filename, total_keluar, total_masuk, total_baris = sheets_manager.export_excel_bytes()
        caption = (
            f"📊 <b>DOKUMEN LAPORAN GUDANG SIAP!</b>\n"
            f"📁 <code>{filename}</code> | Keluar: <b>{total_keluar} pcs</b> | Masuk: <b>{total_masuk} pcs</b>\n"
            f"<i>File Excel terlampir siap dibuka langsung di HP / Komputer.</i>"
        )
        telegram_client.send_document(chat_id, excel_bytes, filename, caption=caption, parse_mode="HTML")
        return True, "cb_export_sent"

    # 12. Menu Khusus Superadmin (Kelola Admin & Keamanan)
    elif data == "btn_superadmin_menu":
        if not is_superadmin(user_id):
            telegram_client.answer_callback_query(
                cb_id, text="⛔ Akses Ditolak: Fitur ini khusus untuk Superadmin (Owner).", show_alert=True
            )
            return False, "superadmin_forbidden"
        sa_text, sa_markup = get_superadmin_menu_markup()
        telegram_client.edit_message_text(chat_id, message_id, sa_text, reply_markup=sa_markup)
        return True, "cb_superadmin_menu"

    # 13. Superadmin: Daftar Seluruh Admin
    elif data == "btn_sa_list":
        if not is_superadmin(user_id):
            telegram_client.answer_callback_query(cb_id, text="⛔ Khusus Superadmin.", show_alert=True)
            return False, "superadmin_forbidden"
        auth_data = load_authorized_admins()
        current_pin = get_current_pin()
        admins = auth_data.get("admins", {})
        sa_list = []
        adm_list = []
        for str_uid, info in admins.items():
            try:
                line = f"• <b>{info.get('name', 'Admin')}</b> (<code>{str_uid}</code>) - Ditambahkan: {info.get('added_at', '-')[:10]}"
                if info.get("role") == "superadmin" or int(str_uid) in auth_data.get("superadmin_ids", []):
                    sa_list.append(line)
                else:
                    adm_list.append(line)
            except Exception:
                continue

        list_text = (
            "👥 <b>DAFTAR STAF & ROLE ADMIN GUDANG:</b>\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👑 <b>Superadmin (Owner):</b>\n" + ("\n".join(sa_list) or "• Belum diset") + "\n\n"
            f"🛠️ <b>Admin Gudang Terdaftar:</b>\n" + ("\n".join(adm_list) or "• Belum ada staf admin tambahan") + "\n\n"
            f"🔑 <b>PIN Masuk Aktif:</b> <code>{current_pin}</code>\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "<i>💡 Untuk menambah staf baru, klik tombol 'Tambah Admin' atau ketik <code>/tambah_admin [ID] [Nama]</code></i>"
        )
        back_markup = {
            "inline_keyboard": [
                [{"text": "➕ Tambah Admin Baru", "callback_data": "btn_sa_add"}],
                [{"text": "🔙 Kembali ke Menu Superadmin", "callback_data": "btn_superadmin_menu"}]
            ]
        }
        telegram_client.edit_message_text(chat_id, message_id, list_text, reply_markup=back_markup)
        return True, "cb_sa_list"

    # 14. Superadmin: Panduan Tambah Admin Baru
    elif data == "btn_sa_add":
        if not is_superadmin(user_id):
            telegram_client.answer_callback_query(cb_id, text="⛔ Khusus Superadmin.", show_alert=True)
            return False, "superadmin_forbidden"
        add_text = (
            "➕ <b>CARA MENAMBAH ADMIN BARU:</b>\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Superadmin dapat mendaftarkan ID Telegram staf agar langsung bisa mengakses bot secara japri tanpa perlu memasukkan PIN:\n\n"
            "Ketik perintah di chat ini:\n"
            "<code>/tambah_admin [ID_Telegram] [Nama_Staf]</code>\n\n"
            "Contoh:\n"
            "<code>/tambah_admin 7957273991 Dwi Cahya</code>\n\n"
            "<i>💡 Minta staf untuk membuka chat @mh_gudang_bot lalu ketik /start. Bot akan menampilkan ID Telegram staf tersebut di layar ponsel mereka.</i>"
        )
        back_markup = {
            "inline_keyboard": [
                [{"text": "👥 Lihat Daftar Admin", "callback_data": "btn_sa_list"}],
                [{"text": "🔙 Kembali ke Menu Superadmin", "callback_data": "btn_superadmin_menu"}]
            ]
        }
        telegram_client.edit_message_text(chat_id, message_id, add_text, reply_markup=back_markup)
        return True, "cb_sa_add"

    # 15. Superadmin: Menu Cabut Hak Akses Admin
    elif data == "btn_sa_del_menu":
        if not is_superadmin(user_id):
            telegram_client.answer_callback_query(cb_id, text="⛔ Khusus Superadmin.", show_alert=True)
            return False, "superadmin_forbidden"
        del_text, del_markup = get_sa_delete_admin_markup()
        telegram_client.edit_message_text(chat_id, message_id, del_text, reply_markup=del_markup)
        return True, "cb_sa_del_menu"

    # 16. Superadmin: Eksekusi Cabut Hak Akses
    elif data.startswith("cb_sa_del_"):
        if not is_superadmin(user_id):
            telegram_client.answer_callback_query(cb_id, text="⛔ Khusus Superadmin.", show_alert=True)
            return False, "superadmin_forbidden"
        target_uid = int(data.replace("cb_sa_del_", ""))
        ok, msg = remove_admin_member(target_uid)
        del_text, del_markup = get_sa_delete_admin_markup()
        res_text = f"✅ <b>HASIL:</b> {msg}\n\n" + del_text
        telegram_client.edit_message_text(chat_id, message_id, res_text, reply_markup=del_markup)
        return True, "cb_sa_del_executed"

    # 17. Superadmin: Panduan Ganti PIN Gudang
    elif data == "btn_sa_change_pin":
        if not is_superadmin(user_id):
            telegram_client.answer_callback_query(cb_id, text="⛔ Khusus Superadmin.", show_alert=True)
            return False, "superadmin_forbidden"
        current_pin = get_current_pin()
        pin_text = (
            "🔑 <b>CARA MENGGANTI PIN GUDANG:</b>\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            f"PIN aktif saat ini: <code>{current_pin}</code>\n\n"
            "Untuk mengatur PIN masuk baru, ketik perintah:\n"
            "<code>/ubah_pin [PIN_BARU]</code>\n\n"
            "Contoh:\n"
            "<code>/ubah_pin 5678</code>\n\n"
            "<i>💡 PIN baru harus berupa 4 hingga 8 digit angka. Setelah diubah, staf yang login dengan PIN lama akan otomatis terputus dan wajib memasukkan PIN baru ini.</i>"
        )
        back_markup = {
            "inline_keyboard": [
                [{"text": "🔙 Kembali ke Menu Superadmin", "callback_data": "btn_superadmin_menu"}]
            ]
        }
        telegram_client.edit_message_text(chat_id, message_id, pin_text, reply_markup=back_markup)
        return True, "cb_sa_change_pin"

    # 18. Superadmin: Kunci / Reset Semua Sesi Staf
    elif data == "btn_sa_reset_sessions":
        if not is_superadmin(user_id):
            telegram_client.answer_callback_query(cb_id, text="⛔ Khusus Superadmin.", show_alert=True)
            return False, "superadmin_forbidden"
        count = reset_all_admin_sessions(except_superadmin=True)
        current_pin = get_current_pin()
        reset_text = (
            f"🔒 <b>SELURUH SESI LOGIN STAF TELAH DIKUNCI!</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"Sebanyak <b>{count} sesi login</b> staf admin telah dikeluarkan (logged out).\n"
            f"Staf admin yang ingin membuka bot kembali wajib memasukkan PIN aktif: <code>{current_pin}</code>."
        )
        back_markup = {
            "inline_keyboard": [
                [{"text": "🔙 Kembali ke Menu Superadmin", "callback_data": "btn_superadmin_menu"}]
            ]
        }
        telegram_client.edit_message_text(chat_id, message_id, reset_text, reply_markup=back_markup)
        return True, "cb_sa_reset_sessions"

    elif data == "btn_noop":
        return True, "noop"

    return False, "unknown_callback"


def handle_incoming_message(message: dict) -> Tuple[bool, str, dict]:
    """
    Memproses setiap pesan masuk dari Telegram (Webhook maupun Polling):
    1. Pemisahan Mode: Grup (khusus resi & monitoring) vs Japri (Panel Admin & Tombol)
    2. Cek Anti-Duplicate Resi
    3. Cek Leaderboard Seller
    4. Cek Stock Opname
    5. Cek Restok Masuk & Broadcast Otomatis ke Grup
    6. Ekspor Excel & Laporan
    7. Normalisasi Barang Keluar Multi-Item & Konversi Kemasan
    """
    chat = message.get("chat", {})
    chat_id = chat.get("id")
    chat_type = chat.get("type", "")
    chat_title = chat.get("title", "")
    message_id = message.get("message_id")
    from_user = message.get("from", {})
    user_id = from_user.get("id", 0)
    user_name = (from_user.get("first_name", "") + " " + from_user.get("last_name", "")).strip() or from_user.get("username", "Seller")

    if chat_type not in ["group", "supergroup", "private"]:
        return False, "unsupported_chat_type", {}

    if chat_type in ["group", "supergroup"]:
        save_active_group(chat_id, chat_title)

    pesan_seller, media_bytes, media_mime_type = extract_content_from_message(message)
    if not pesan_seller and not media_bytes:
        return False, "no_content", {}

    clean_text = pesan_seller.strip().lower()

    # -------------------------------------------------------------------------
    # PEMISAHAN MODE 1: PROTEKSI PERINTAH ADMIN DI GRUP
    # -------------------------------------------------------------------------
    is_admin_cmd_attempt = clean_text.startswith((
        "/restok", "restok ", "/tambah_sku", "/tambahsku", "/opname", "opname ",
        "/laporan", "laporan", "/export", "/excel", "kirim excel", "download excel",
        "unduh excel", "laporan excel", "/leaderboard", "/top_seller", "leaderboard", "top seller", "rekap seller"
    ))

    if chat_type in ["group", "supergroup"] and is_admin_cmd_attempt:
        warning_grup = (
            "⛔ <b>Perintah Dibatasi:</b>\n"
            "Perintah ini merupakan fitur internal yang hanya dapat diakses oleh Admin yang memiliki otorisasi di chat pribadi.\n\n"
            "<i>Di grup ini, Anda dapat mengecek stok barang dengan <code>/stok</code> atau mengirim resi pesanan kurir.</i>"
        )
        telegram_client.send_message(chat_id, warning_grup, reply_to_message_id=message_id, parse_mode="HTML")
        return False, "admin_action_blocked_in_group", {}

    # -------------------------------------------------------------------------
    # PEMISAHAN MODE 2: PANEL KHUSUS ADMIN DI CHAT PRIBADI (JAPRI) DENGAN PIN & ID
    # -------------------------------------------------------------------------
    if chat_type == "private":
        # 1. Cek Perintah Logout / Kunci
        if clean_text in ["/logout", "logout", "/kunci", "kunci"]:
            logout_admin_session(user_id)
            lock_msg = (
                "🔒 <b>PANEL ADMIN TELAH DIKUNCI (LOGOUT)</b>\n"
                "━━━━━━━━━━━━━━━━━━━━━━\n"
                "Sesi kendali Anda telah diakhiri.\n\n"
                "🔑 Silakan masukkan <b>PIN Admin</b> kembali untuk membuka menu kendali:\n"
                "<i>(Contoh: <code>1234</code>)</i>"
            )
            telegram_client.send_message(chat_id, lock_msg, reply_to_message_id=message_id, parse_mode="HTML")
            return True, "admin_logged_out", {}

        # 2. Cek Manajemen Superadmin & Admin (/superadmin, /ubah_pin, /tambah_admin, /hapus_admin, /list_admin)
        if is_user_authorized(user_id):
            is_sa = is_superadmin(user_id)

            if clean_text in ["/superadmin", "superadmin", "/owner", "owner"]:
                if not is_sa:
                    telegram_client.send_message(
                        chat_id,
                        "⛔ <b>Akses Ditolak:</b> Menu ini khusus untuk <b>Superadmin (Owner)</b>.",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return False, "superadmin_forbidden", {}
                sa_text, sa_markup = get_superadmin_menu_markup()
                telegram_client.send_message_with_keyboard(
                    chat_id, sa_text, reply_markup=sa_markup, parse_mode="HTML"
                )
                return True, "superadmin_menu_opened", {}

            elif clean_text.startswith("/ubah_pin") or clean_text.startswith("/ganti_pin") or clean_text.startswith("/ubahpin"):
                if not is_sa:
                    telegram_client.send_message(
                        chat_id,
                        "⛔ <b>Akses Ditolak:</b> Hanya <b>Superadmin (Owner)</b> yang berhak mengubah PIN login gudang.",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return False, "change_pin_forbidden", {}

                parts = clean_text.split()
                if len(parts) >= 2 and parts[1].isdigit() and len(parts[1]) >= 4:
                    new_pin = parts[1].strip()
                    set_current_pin(new_pin, changed_by=user_name)
                    reset_count = reset_all_admin_sessions(except_superadmin=True)
                    reply_pin = (
                        "🔑 <b>PIN GUDANG BERHASIL DIUBAH!</b>\n"
                        "━━━━━━━━━━━━━━━━━━━━━━\n"
                        f"PIN baru aktif sekarang: <code>{new_pin}</code>\n"
                        f"👤 Diubah oleh: <b>{user_name} (Superadmin)</b>\n\n"
                        f"🔒 Sebanyak <b>{reset_count} sesi login</b> staf admin telah dikeluarkan (reset) demi keamanan. "
                        f"Staf yang masuk kembali wajib menggunakan PIN baru ini."
                    )
                    telegram_client.send_message(chat_id, reply_pin, reply_to_message_id=message_id, parse_mode="HTML")
                    return True, "pin_changed_success", {"new_pin": new_pin}
                else:
                    telegram_client.send_message(
                        chat_id,
                        "⚠️ Format salah! PIN harus berupa minimal 4 angka.\nContoh: <code>/ubah_pin 5678</code>",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return False, "invalid_pin_format", {}

            elif clean_text.startswith("/tambah_admin") or clean_text.startswith("/tambahadmin"):
                if not is_sa:
                    telegram_client.send_message(
                        chat_id,
                        "⛔ <b>Akses Ditolak:</b> Hanya <b>Superadmin (Owner)</b> yang berhak menambahkan admin baru.",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return False, "tambah_admin_forbidden", {}

                parts = pesan_seller.split(maxsplit=2)
                if len(parts) >= 2 and parts[1].isdigit():
                    new_aid = int(parts[1])
                    nama_staf = parts[2].strip() if len(parts) > 2 else f"Admin-{new_aid}"
                    add_admin_member(new_aid, nama_staf, role="admin", added_by=user_name)
                    succ_msg = (
                        "✅ <b>ADMIN BARU BERHASIL DITAMBAHKAN!</b>\n"
                        "━━━━━━━━━━━━━━━━━━━━━━\n"
                        f"🏷️ <b>Nama:</b> {nama_staf}\n"
                        f"🆔 <b>ID Telegram:</b> <code>{new_aid}</code>\n"
                        f"🛡️ <b>Role:</b> Admin Gudang\n\n"
                        f"<i>Staf ini kini dapat langsung membuka menu admin di chat pribadi bot tanpa perlu memasukkan PIN!</i>"
                    )
                    telegram_client.send_message(chat_id, succ_msg, reply_to_message_id=message_id, parse_mode="HTML")
                    return True, "admin_added", {"id": new_aid, "name": nama_staf}
                else:
                    telegram_client.send_message(
                        chat_id,
                        "⚠️ Format salah! Gunakan format:\n<code>/tambah_admin [ID_Telegram] [Nama_Staf]</code>\nContoh: <code>/tambah_admin 7957273991 Dwi Cahya</code>",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return False, "invalid_tambah_admin", {}

            elif clean_text.startswith("/hapus_admin") or clean_text.startswith("/hapusadmin"):
                if not is_sa:
                    telegram_client.send_message(
                        chat_id,
                        "⛔ <b>Akses Ditolak:</b> Hanya <b>Superadmin (Owner)</b> yang berhak mencabut hak akses admin.",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return False, "hapus_admin_forbidden", {}

                parts = pesan_seller.split()
                if len(parts) >= 2 and parts[1].isdigit():
                    del_aid = int(parts[1])
                    ok, msg = remove_admin_member(del_aid)
                    icon = "✅" if ok else "⚠️"
                    telegram_client.send_message(
                        chat_id,
                        f"{icon} {msg}",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return ok, "admin_removed", {}
                else:
                    telegram_client.send_message(
                        chat_id,
                        "⚠️ Format salah! Gunakan format:\n<code>/hapus_admin [ID_Telegram]</code>\nContoh: <code>/hapus_admin 7957273991</code>",
                        reply_to_message_id=message_id,
                        parse_mode="HTML"
                    )
                    return False, "invalid_hapus_admin", {}

            elif clean_text in ["/list_admin", "/daftar_admin", "list admin", "daftar admin"]:
                auth_data = load_authorized_admins()
                current_pin = get_current_pin()
                admins = auth_data.get("admins", {})
                sa_list = []
                adm_list = []
                for str_uid, info in admins.items():
                    try:
                        line = f"• <b>{info.get('name', 'Admin')}</b> (<code>{str_uid}</code>) - Sejak: {info.get('added_at', '-')[:10]}"
                        if info.get("role") == "superadmin" or int(str_uid) in auth_data.get("superadmin_ids", []):
                            sa_list.append(line)
                        else:
                            adm_list.append(line)
                    except Exception:
                        continue

                list_text = (
                    "👥 <b>DAFTAR ANGGOTA & ROLE ADMIN GUDANG:</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━━━\n"
                    f"👑 <b>Superadmin (Owner):</b>\n" + ("\n".join(sa_list) or "• Belum diset") + "\n\n"
                    f"🛠️ <b>Admin Gudang:</b>\n" + ("\n".join(adm_list) or "• Belum ada staf admin tambahan") + "\n\n"
                    f"🔑 <b>PIN Masuk Aktif:</b> <code>{current_pin}</code>\n"
                    "━━━━━━━━━━━━━━━━━━━━━━\n"
                    "<i>💡 Superadmin dapat menggunakan <code>/tambah_admin</code>, <code>/hapus_admin</code>, dan <code>/ubah_pin</code>.</i>"
                )
                telegram_client.send_message(chat_id, list_text, reply_to_message_id=message_id, parse_mode="HTML")
                return True, "list_admin_sent", {}

        # 3. Cek Input PIN
        is_pin_attempt = (
            clean_text == get_current_pin().lower() or
            clean_text == ADMIN_PIN.lower() or
            clean_text.isdigit() or
            any(clean_text.startswith(p) for p in ["/login", "/pin", "pin=", "pin:"])
        )

        if is_pin_attempt:
            if verify_and_login_pin(user_id, user_name, clean_text):
                is_sa = is_superadmin(user_id)
                role_label = "Superadmin (Owner) 👑" if is_sa else "Admin Gudang 🚀"
                login_success_text = (
                    f"🔓 <b>LOGIN ADMIN BERHASIL!</b>\n"
                    f"━━━━━━━━━━━━━━━━━━━━━━\n"
                    f"Selamat datang <b>{user_name}</b>!\n"
                    f"🆔 <b>ID Telegram:</b> <code>{user_id}</code>\n"
                    f"🛡️ <b>Role Akses:</b> <b>{role_label}</b>\n\n"
                    f"Panel Kendali Operasional Gudang kini telah terbuka:"
                )
                telegram_client.send_message_with_keyboard(
                    chat_id,
                    login_success_text,
                    reply_markup=get_admin_dashboard_markup(is_superadmin_user=is_sa),
                    parse_mode="HTML"
                )
                return True, "pin_login_success", {}
            else:
                if clean_text.isdigit() or any(clean_text.startswith(p) for p in ["/login", "/pin", "pin=", "pin:"]):
                    denied_msg = (
                        "❌ <b>PIN ADMIN SALAH!</b>\n"
                        "━━━━━━━━━━━━━━━━━━━━━━\n"
                        "Akses ke panel kendali gudang ditolak.\n\n"
                        "Silakan masukkan PIN Admin yang valid, atau hubungi Owner untuk mendaftarkan ID Telegram Anda:\n"
                        f"🆔 <b>ID Anda:</b> <code>{user_id}</code>"
                    )
                    telegram_client.send_message(chat_id, denied_msg, reply_to_message_id=message_id, parse_mode="HTML")
                    return False, "invalid_pin_rejected", {}

        # 4. Cek Kirim Kontak / Nomor Telepon
        contact = message.get("contact")
        if contact and contact.get("phone_number"):
            phone_num = contact.get("phone_number")
            if verify_and_login_phone(user_id, user_name, phone_num):
                is_sa = is_superadmin(user_id)
                role_label = "Superadmin (Owner) 👑" if is_sa else "Admin Gudang 🚀"
                login_phone_text = (
                    f"🔓 <b>LOGIN NOMOR TELEPON BERHASIL!</b>\n"
                    f"━━━━━━━━━━━━━━━━━━━━━━\n"
                    f"Selamat datang <b>{user_name}</b>!\n"
                    f"Nomor HP <code>{phone_num}</code> terverifikasi sebagai nomor admin resmi.\n"
                    f"🛡️ <b>Role:</b> {role_label}\n\n"
                    f"Panel Kendali Operasional Gudang kini telah aktif:"
                )
                telegram_client.send_message_with_keyboard(
                    chat_id,
                    login_phone_text,
                    reply_markup=get_admin_dashboard_markup(is_superadmin_user=is_sa),
                    parse_mode="HTML"
                )
                return True, "phone_login_success", {}

        # 5. Cek apakah User SUDAH diotorisasi
        if not is_user_authorized(user_id):
            login_prompt = (
                f"🔒 <b>AKSES TERBATAS — PANEL ADMIN GUDANG TIRA</b>\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"Halo <b>{user_name}</b>!\n"
                f"Chat bot pribadi ini diproteksi khusus untuk pengelolaan stok gudang internal.\n\n"
                f"🔑 <b>Silakan masukkan PIN Admin</b> untuk membuka menu kendali:\n"
                f"<i>(Cukup ketik PIN Anda langsung di chat ini, contoh: <code>1234</code>)</i>\n\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"🆔 <b>ID Telegram Anda:</b> <code>{user_id}</code>\n"
                f"<i>(ID ini dapat didaftarkan langsung oleh Owner ke sistem)</i>"
            )
            telegram_client.send_message(chat_id, login_prompt, reply_to_message_id=message_id, parse_mode="HTML")
            return False, "unauthorized_user_prompted_for_pin", {}

        # 6. User Terverifikasi: Tampilkan Dashboard jika kirim salam / /start / menu
        greeting_words = [
            "tes", "test", "tira", "halo", "hai", "/start", "/menu", "menu", "admin", "panel", "buka panel"
        ]
        is_japri_menu = clean_text in greeting_words

        if is_japri_menu and not media_bytes:
            is_sa = is_superadmin(user_id)
            role_label = "Superadmin (Owner) 👑" if is_sa else "Admin Gudang 🚀"
            dashboard_text = (
                f"🎛️ <b>PANEL KENDALI GUDANG TIRA ({role_label})</b>\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"Halo <b>{user_name}</b>! Seluruh operasional gudang dapat Anda kendalikan <b>100% langsung di Telegram</b> tanpa perlu membuka browser web.\n\n"
                f"Silakan pilih menu aksi cepat di bawah ini:"
            )
            telegram_client.send_message_with_keyboard(
                chat_id,
                dashboard_text,
                reply_markup=get_admin_dashboard_markup(is_superadmin_user=is_sa),
                parse_mode="HTML"
            )
            return True, "admin_dashboard_opened", {}

    # -------------------------------------------------------------------------
    # FITUR LEADERBOARD & REKAP SELLER (Khusus Admin di Chat Pribadi)
    # -------------------------------------------------------------------------
    if clean_text in ["/leaderboard", "/top_seller", "leaderboard", "top seller", "rekap seller"]:
        if chat_type != "private":
            telegram_client.send_message(
                chat_id,
                "⛔ <i>Leaderboard Seller merupakan fitur internal khusus Admin di chat pribadi.</i>",
                reply_to_message_id=message_id,
                parse_mode="HTML"
            )
            return False, "leaderboard_not_in_private", {}
        lb_text = sheets_manager.get_leaderboard_text()
        telegram_client.send_message(chat_id, lb_text, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "leaderboard_sent", {}

    # -------------------------------------------------------------------------
    # FITUR KOREKSI STOCK OPNAME (Khusus Admin di Japri)
    # -------------------------------------------------------------------------
    if clean_text.startswith("/opname") or clean_text.startswith("opname"):
        if chat_type != "private":
            telegram_client.send_message(
                chat_id,
                "⛔ <i>Perintah Stock Opname dibatasi untuk otorisasi admin internal di chat pribadi.</i>",
                reply_to_message_id=message_id,
                parse_mode="HTML"
            )
            return False, "opname_not_in_private", {}

        # Format: /opname [SKU] [selisih] [alasan]
        parts = pesan_seller.split(maxsplit=3)
        if len(parts) < 4:
            msg = (
                "⚠️ Format salah! Gunakan format:\n"
                "<code>/opname [Kode_SKU] [Selisih_Qty] [Alasan]</code>\n\n"
                "Contoh (Barang Pecah / Hilang):\n"
                "<code>/opname MHPD 021 -2 Pecah saat bongkar muat</code>\n\n"
                "Contoh (Fisik Lebih Banyak):\n"
                "<code>/opname MHPD 004 +5 Lebih hitung fisik gudang</code>"
            )
            telegram_client.send_message(chat_id, msg, reply_to_message_id=message_id, parse_mode="HTML")
            return False, "invalid_opname_format", {}

        sku_target = parts[1].strip().upper()
        try:
            diff_qty = int(parts[2].replace("+", "").strip())
        except ValueError:
            telegram_client.send_message(chat_id, "⚠️ Nilai selisih Qty harus berupa angka (contoh: -2 atau +5).", reply_to_message_id=message_id)
            return False, "invalid_opname_qty", {}

        alasan_opname = parts[3].strip()
        cur_qty, new_qty, success, nama_barang = sheets_manager.perform_stock_opname(
            sku_target, diff_qty, alasan_opname, admin_name=user_name
        )

        if not success:
            telegram_client.send_message(chat_id, f"⚠️ SKU <code>{sku_target}</code> tidak ditemukan di database stok.", reply_to_message_id=message_id, parse_mode="HTML")
            return False, "opname_sku_not_found", {}

        tanda = "+" if diff_qty >= 0 else ""
        res_opname = (
            f"🔄 <b>KOREKSI STOCK OPNAME BERHASIL (ADMIN):</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🏷️ <b>SKU:</b> <code>[{sku_target}]</code>\n"
            f"📦 <b>Produk:</b> {nama_barang}\n"
            f"📉 <b>Penyesuaian Fisik:</b> <b>{tanda}{diff_qty} pcs</b>\n"
            f"📊 <b>Stok Sebelumnya:</b> {cur_qty} pcs ➡️ <b>Stok Baru:</b> {new_qty} pcs\n"
            f"📝 <b>Alasan:</b> <i>{alasan_opname}</i>\n"
            f"👤 <b>Petugas:</b> {user_name}\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"<i>Data tersimpan resmi di Tab Riwayat_Mutasi &amp; Stok_Barang.</i>"
        )
        telegram_client.send_message(chat_id, res_opname, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "opname_success", {"sku": sku_target, "diff": diff_qty}

    # -------------------------------------------------------------------------
    # FITUR RESTOK MASUK ADMIN & BROADCAST OTOMATIS KE GRUP
    # -------------------------------------------------------------------------
    if clean_text.startswith("/restok") or clean_text.startswith("restok"):
        qty_masuk = extract_qty_from_text(pesan_seller)
        clean_target = re.sub(r"^/restok\b|^restok\b", "", pesan_seller, flags=re.IGNORECASE).strip()
        clean_target = re.sub(r"\b\d+\s*(?:pcs|pc|buah|btl|botol|pack|box|dus|lusin)?\b", "", clean_target, flags=re.IGNORECASE).strip()
        clean_target = re.sub(r"pin=\S+", "", clean_target, flags=re.IGNORECASE).strip()

        target_sku = sheets_manager.search_kamus_alias(clean_target) if clean_target else None
        if not target_sku and clean_target:
            daftar_master = sheets_manager.get_master_barang_reference()
            ai_match = gemini_normalizer.normalize(clean_target, daftar_master)
            if ai_match and ai_match.get("items"):
                first_item = ai_match["items"][0]
                if first_item.get("kode_sku") != "UNKNOWN":
                    target_sku = first_item.get("kode_sku")

        if not target_sku:
            telegram_client.send_message(
                chat_id,
                f"⚠️ Tidak dapat mengenali barang yang ingin direstok dari teks: '{clean_target}'.\n"
                "Gunakan format: <code>/restok [Kode_SKU / Nama] [Qty]</code>",
                reply_to_message_id=message_id,
                parse_mode="HTML"
            )
            return False, "restok_unknown_sku", {}

        nama_resmi = sheets_manager.get_nama_standar_by_sku(target_sku)
        current_qty, new_qty, _, _ = sheets_manager.update_or_create_stok(target_sku, qty_masuk, action="MASUK")
        sheets_manager.record_mutasi(
            tipe="MASUK",
            kode_sku=target_sku,
            nama_barang=nama_resmi,
            qty=qty_masuk,
            sisa_stok=new_qty,
            pengirim=user_name,
            keterangan="Restok Pabrik / Suplai Admin"
        )

        reply_msg = (
            f"📥 <b>RESTOK MASUK BERHASIL (ADMIN):</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🏷️ <b>SKU:</b> <code>[{target_sku}]</code>\n"
            f"📦 <b>Barang:</b> {nama_resmi}\n"
            f"🔺 <b>Qty Masuk:</b> +{qty_masuk} pcs\n"
            f"📊 <b>Total Stok Sekarang:</b> {new_qty} pcs\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"<i>Tercatat di Riwayat Mutasi oleh {user_name}.</i>"
        )
        telegram_client.send_message(chat_id, reply_msg, reply_to_message_id=message_id, parse_mode="HTML")

        # FITUR: BROADCAST NOTIFIKASI RESTOK OTOMATIS KE GRUP
        broadcast_restok_to_groups(target_sku, nama_resmi, qty_masuk, new_qty, admin_name=user_name)
        return True, "restok_success_and_broadcasted", {"sku": target_sku, "qty": qty_masuk}

    # -------------------------------------------------------------------------
    # FITUR TAMBAH SKU BARU KHUSUS ADMIN
    # -------------------------------------------------------------------------
    if clean_text.startswith("/tambah_sku") or clean_text.startswith("/tambahsku"):
        raw_cmd = pesan_seller.split(maxsplit=1)
        if len(raw_cmd) < 2 or "|" not in raw_cmd[1]:
            msg = (
                "⚠️ Format salah! Gunakan format:\n"
                "<code>/tambah_sku [Kode_SKU] | [Nama_Standar] | [Stok_Awal]</code>\n"
                "Contoh: <code>/tambah_sku MHPD 029 | KOPI SEHAT | 100</code>"
            )
            telegram_client.send_message(chat_id, msg, reply_to_message_id=message_id, parse_mode="HTML")
            return False, "invalid_format", {}

        params = [p.strip() for p in raw_cmd[1].split("|")]
        new_sku = params[0].upper()
        new_nama = params[1]
        stok_awal = int(params[2]) if len(params) >= 3 and params[2].isdigit() else 0

        success = sheets_manager.add_master_sku(new_sku, new_nama, stok_awal)
        if success:
            res_msg = (
                f"✅ <b>MASTER SKU BERHASIL DITAMBAHKAN (ADMIN):</b>\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"🏷️ <b>Kode SKU:</b> <code>{new_sku}</code>\n"
                f"📦 <b>Nama Standar:</b> {new_nama}\n"
                f"📊 <b>Stok Awal:</b> {stok_awal} pcs\n"
                f"<i>Data resmi telah tersimpan di Tab Master_Barang &amp; Stok_Barang.</i>"
            )
        else:
            res_msg = f"⚠️ Kode SKU <code>{new_sku}</code> sudah terdaftar sebelumnya di Master_Barang."

        telegram_client.send_message(chat_id, res_msg, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "sku_added", {}

    # -------------------------------------------------------------------------
    # FITUR EKSPOR LAPORAN EXCEL (.xlsx) KE CHAT (Khusus Admin di Chat Pribadi)
    # -------------------------------------------------------------------------
    if clean_text in ["/export", "/excel", "/laporan_excel", "kirim excel", "download excel", "unduh excel", "laporan excel"] or clean_text.startswith("/export"):
        if chat_type != "private":
            telegram_client.send_message(
                chat_id,
                "⛔ <i>Ekspor file Excel laporan gudang merupakan fitur internal khusus Admin di chat pribadi.</i>",
                reply_to_message_id=message_id,
                parse_mode="HTML"
            )
            return False, "export_not_in_private", {}

        telegram_client.send_message(
            chat_id,
            "⏳ <i>Sedang menyiapkan dokumen Excel laporan gudang terbaru dari Google Sheets...</i>",
            reply_to_message_id=message_id,
            parse_mode="HTML"
        )
        try:
            excel_bytes, filename, total_keluar, total_masuk, total_baris = sheets_manager.export_excel_bytes()
            caption = (
                f"📊 <b>DOKUMEN LAPORAN GUDANG SIAP!</b>\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"📅 <b>Periode:</b> {datetime.now().strftime('%d %B %Y')}\n"
                f"📁 <b>Nama File:</b> <code>{filename}</code>\n\n"
                f"🔻 Total Barang Keluar: <b>{total_keluar} pcs</b>\n"
                f"🔺 Total Restok / Masuk: <b>{total_masuk} pcs</b>\n"
                f"📑 Total Baris Mutasi: <b>{total_baris} transaksi</b>\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"<i>💡 Berisi 2 Tab: 'Mutasi Gudang' &amp; 'Stok Gudang Terkini' dengan formatting warna.</i>"
            )
            success = telegram_client.send_document(
                chat_id,
                excel_bytes,
                filename,
                caption=caption,
                reply_to_message_id=message_id,
                parse_mode="HTML"
            )
            return success, "excel_exported", {}
        except Exception as e:
            logger.error(f"Gagal generate laporan Excel: {e}")
            telegram_client.send_message(
                chat_id,
                f"⚠️ Gagal membuat file Excel: {e}",
                reply_to_message_id=message_id
            )
            return False, "excel_error", {}

    # -------------------------------------------------------------------------
    # FITUR BRIEFING PAGI & REKAP SORE MANUAL
    # -------------------------------------------------------------------------
    if clean_text in ["/briefing_pagi", "/briefing", "briefing pagi"]:
        briefing_text = sheets_manager.get_morning_briefing_text()
        telegram_client.send_message(chat_id, briefing_text, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "briefing_pagi_sent", {}

    if clean_text in ["/rekap_sore", "/rekap", "rekap sore", "tutup buku"]:
        recap_text = sheets_manager.get_evening_recap_text()
        telegram_client.send_message(chat_id, recap_text, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "rekap_sore_sent", {}

    # -------------------------------------------------------------------------
    # LAPORAN TEKS MUTASI (Khusus Admin di Chat Pribadi) & CEK SISA STOK
    # -------------------------------------------------------------------------
    if clean_text.startswith("/laporan") or any(k in clean_text for k in ["laporan hari ini", "rekap hari ini", "rekap mutasi", "laporan stok"]):
        if chat_type != "private":
            telegram_client.send_message(
                chat_id,
                "⛔ <i>Laporan mutasi gudang merupakan fitur internal khusus Admin di chat pribadi.</i>",
                reply_to_message_id=message_id,
                parse_mode="HTML"
            )
            return False, "laporan_not_in_private", {}

        parts = clean_text.split()
        target_date = None
        for p in parts[1:]:
            if re.match(r"^\d{4}-\d{2}-\d{2}$", p):
                target_date = p
                break
            elif p in ["kemarin", "yesterday"]:
                target_date = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
                break

        report_text = sheets_manager.get_laporan_mutasi(target_date)
        telegram_client.send_message(chat_id, report_text, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "laporan_sent", {}

    if clean_text.startswith("/stok") or any(k in clean_text for k in ["cek stok", "sisa stok", "lihat stok", "status stok"]):
        query = None
        if clean_text.startswith("/stok"):
            parts = clean_text.split(maxsplit=1)
            query = parts[1].strip() if len(parts) > 1 else None
        else:
            match_item = re.search(r"(?:cek stok|sisa stok|lihat stok)\s+(.+)", clean_text)
            if match_item:
                query = match_item.group(1).strip()

        stok_text = sheets_manager.get_stok_overview(query)
        telegram_client.send_message(chat_id, stok_text, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "stok_sent", {}

    # -------------------------------------------------------------------------
    # GREETING / PERKENALAN DIRI DI GRUP
    # -------------------------------------------------------------------------
    greeting_words = [
        "tes", "test", "testing", "tira", "halo tira", "hai tira", "hei tira",
        "halo", "hai", "pagi", "siang", "sore", "malam",
        "/start", "/help", "help", "panduan", "/panduan", "menu", "/menu",
        "@mh_gudang_bot"
    ]
    is_greeting_match = (clean_text in greeting_words) or any(
        clean_text.startswith(w) for w in ["halo tira", "hai tira", "tira ", "hei tira", "tes ", "test "]
    )
    is_order_intent = any(k in clean_text for k in [
        "siapin", "kirim", "order", "ambilkan", "minta", "pcs", "botol", "sachet",
        "box", "dus", "lusin", "kodi", "resi"
    ])

    if is_greeting_match and not is_order_intent and not media_bytes:
        help_msg = (
            f"🤖 <b>Halo {user_name}! Kenalkan, aku Tira</b> 👋\n"
            f"<i>AI Asisten Operasional Logistik &amp; Gudang Anda</i>\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Aku bertugas mencatat barang keluar dari resi kurir dan orderan seller secara otomatis.\n\n"
            "📋 <b>FITUR YANG BISA ANDA GUNAKAN DI GRUP INI:</b>\n\n"
            "📦 <b>1. Pencatatan Barang Keluar:</b>\n"
            "• Kirim file <b>PDF Resi</b> atau <b>Foto Resi</b> kurir\n"
            "• Ketik order: <i>'Kirim 1 Dephero dan 2 Habbafit'</i>\n"
            "• Ketik kemasan: <i>'Siapin 2 box G-MAX & 1 lusin Bio Binahong'</i>\n"
            "👉 <i>Stok otomatis dipotong &amp; dilindungi Anti-Duplicate Resi!</i>\n\n"
            "📊 <b>2. Pengecekan Sisa Stok:</b>\n"
            "• <code>/stok</code> (Cek seluruh sisa stok barang gudang)\n"
            "• <code>/stok [Nama / SKU]</code> (Cek sisa stok produk tertentu, misal: <code>/stok dephero</code>)\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "💡 <i>Terhubung langsung ke Google Sheets secara real-time.</i>"
        )
        telegram_client.send_message(chat_id, help_msg, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "greeting_and_docs_sent", {}

    # -------------------------------------------------------------------------
    # PROSES BARANG KELUAR: CEK RESI GANDA, MULTI-ITEM, KONVERSI KEMASAN
    # -------------------------------------------------------------------------
    logger.info(f"Memproses orderan / resi barang keluar: '{pesan_seller[:80]}' (media={bool(media_bytes)})")

    detected_resi = extract_no_resi_from_text(pesan_seller)

    if detected_resi:
        dup_info = sheets_manager.check_duplicate_resi(detected_resi)
        if dup_info:
            dup_msg = (
                f"⚠️ Resi [<code>{dup_info['no_resi']}</code>] sudah pernah diproses "
                f"hari ini pukul {dup_info['jam']} oleh {dup_info['pengirim']}!\n\n"
                f"Stok <b>TIDAK</b> dipotong dua kali."
            )
            telegram_client.send_message(chat_id, dup_msg, reply_to_message_id=message_id, parse_mode="HTML")
            logger.info(f"Resi ditolak (DUPLIKAT): {detected_resi}")
            return False, "duplicate_resi_prevented", {"no_resi": detected_resi}

    matched_sku = None
    if pesan_seller and not media_bytes and not detected_resi and not any(k in clean_text for k in [",", " dan ", "&", "+"]):
        try:
            matched_sku = sheets_manager.search_kamus_alias(pesan_seller)
        except Exception as e:
            logger.error(f"Gagal saat membaca Kamus_Alias: {e}")

    if matched_sku:
        qty_keluar = extract_qty_from_text(pesan_seller, matched_sku)
        nama_standar = sheets_manager.get_nama_standar_by_sku(matched_sku)
        current_qty, new_qty, is_stok_kurang, _ = sheets_manager.update_or_create_stok(
            matched_sku, qty_keluar, action="KELUAR"
        )

        sheets_manager.record_mutasi(
            tipe="KELUAR",
            kode_sku=matched_sku,
            nama_barang=nama_standar,
            qty=qty_keluar,
            sisa_stok=new_qty,
            pengirim=user_name,
            keterangan=f"Order: {pesan_seller[:80]}"
        )

        warning_stok = ""
        if is_stok_kurang:
            warning_stok = "\n⚠️ <b>PERINGATAN: Stok sebelumnya tidak mencukupi permintaan order ini!</b>"
        elif new_qty < 10:
            warning_stok = f"\n⚠️ <i>Perhatian: Sisa stok menipis (&lt; 10 pcs). Segera restok!</i>"

        telegram_msg = (
            f"📦 <b>BARANG KELUAR (VIA KAMUS ALIAS):</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🏷️ <b>SKU:</b> <code>[{matched_sku}]</code>\n"
            f"📦 <b>Barang:</b> {nama_standar}\n"
            f"🔻 <b>Qty Keluar:</b> -{qty_keluar} pcs\n"
            f"📊 <b>Sisa Stok:</b> {new_qty} pcs"
            f"{warning_stok}\n"
            f"<i>Tercatat di Riwayat Mutasi.</i>"
        )
        telegram_client.send_message(chat_id, telegram_msg, reply_to_message_id=message_id, parse_mode="HTML")
        return True, "kamus_outflow", {"sku": matched_sku, "qty": qty_keluar}

    # PROSES VIA AI GEMINI
    logger.info("Menghubungi Gemini AI (Multi-Item Normalizer)...")
    try:
        daftar_master_sku = sheets_manager.get_master_barang_reference()
    except Exception as e:
        logger.error(f"Gagal membaca Master_Barang: {e}")
        telegram_client.send_message(
            chat_id,
            "⚠️ Gagal mengambil referensi Master_Barang dari Google Sheets.",
            reply_to_message_id=message_id
        )
        return False, "master_error", {}

    ai_result = gemini_normalizer.normalize(
        pesan_seller,
        daftar_master_sku,
        media_bytes=media_bytes,
        media_mime_type=media_mime_type
    )

    if not ai_result or not ai_result.get("items"):
        telegram_client.send_message(
            chat_id,
            "⚠️ Maaf, AI belum berhasil membaca nama produk dari pesan / resi tersebut. "
            "Pastikan teks atau label produk terbaca jelas.",
            reply_to_message_id=message_id
        )
        return False, "ai_failed", {}

    final_no_resi = detected_resi or ai_result.get("no_resi", "").strip()

    if final_no_resi:
        dup_info = sheets_manager.check_duplicate_resi(final_no_resi)
        if dup_info:
            dup_msg = (
                f"⚠️ Resi [<code>{dup_info['no_resi']}</code>] sudah pernah diproses "
                f"hari ini pukul {dup_info['jam']} oleh {dup_info['pengirim']}!\n\n"
                f"Stok <b>TIDAK</b> dipotong dua kali."
            )
            telegram_client.send_message(chat_id, dup_msg, reply_to_message_id=message_id, parse_mode="HTML")
            logger.info(f"Resi ditolak (DUPLIKAT via AI): {final_no_resi}")
            return False, "duplicate_resi_prevented", {"no_resi": final_no_resi}

    items_to_process = ai_result.get("items", [])
    processed_items = []
    unknown_items = []

    for item in items_to_process:
        kode_sku = item.get("kode_sku", "").strip().upper()
        nama_standar = item.get("nama_standar", "").strip()
        qty_pcs = item.get("qty", 1)
        satuan_asli = item.get("satuan_asli", "pcs")
        nama_asli_seller = item.get("nama_asli_seller", "").strip() or pesan_seller

        if kode_sku == "UNKNOWN" or not kode_sku:
            unknown_items.append(nama_asli_seller or "Produk Tidak Dikenal")
            continue

        nama_resmi = sheets_manager.get_nama_standar_by_sku(kode_sku) or nama_standar
        current_qty, new_qty, is_stok_kurang, _ = sheets_manager.update_or_create_stok(
            kode_sku, qty_pcs, action="KELUAR"
        )

        ket = f"Resi: [{final_no_resi}] | {nama_asli_seller}" if final_no_resi else f"Order: {nama_asli_seller}"
        sheets_manager.record_mutasi(
            tipe="KELUAR",
            kode_sku=kode_sku,
            nama_barang=nama_resmi,
            qty=qty_pcs,
            sisa_stok=new_qty,
            pengirim=user_name,
            keterangan=ket
        )

        try:
            if nama_asli_seller and len(nama_asli_seller) < 150:
                sheets_manager.append_kamus_alias(nama_asli_seller, kode_sku)
        except Exception as e:
            logger.warning(f"Gagal append alias: {e}")

        processed_items.append({
            "sku": kode_sku,
            "nama": nama_resmi,
            "qty": qty_pcs,
            "sisa": new_qty,
            "is_kurang": is_stok_kurang,
            "satuan_asli": satuan_asli,
            "nama_asli": nama_asli_seller
        })

    if not processed_items and unknown_items:
        telegram_msg = (
            f"❓ AI tidak menemukan SKU yang cocok di Master Barang untuk:\n"
            f"<code>\"{', '.join(unknown_items)[:120]}\"</code>\n\n"
            f"Silakan periksa apakah barang sudah terdaftar di Tab Master_Barang."
        )
        telegram_client.send_message(chat_id, telegram_msg, reply_to_message_id=message_id, parse_mode="HTML")
        return False, "sku_unknown", {"unknown": unknown_items}

    is_resi = bool(final_no_resi or media_bytes)
    header_title = "📄 <b>BARANG KELUAR (RESI TERVERIFIKASI):</b>" if is_resi else "🤖 <b>BARANG KELUAR (GEMINI AI):</b>"

    lines = [
        header_title,
        "━━━━━━━━━━━━━━━━━━━━━━"
    ]
    if final_no_resi:
        lines.append(f"📋 <b>No. Resi:</b> <code>{final_no_resi}</code>")
    lines.append(f"👤 <b>Diproses oleh:</b> {user_name}\n")

    if len(processed_items) == 1:
        p = processed_items[0]
        satuan_info = f" <i>({p['satuan_asli']})</i>" if p.get('satuan_asli') and p.get('satuan_asli') != 'pcs' else ''
        lines.append(f"🏷️ <b>SKU:</b> <code>[{p['sku']}]</code>")
        lines.append(f"📦 <b>Barang:</b> {p['nama']}")
        lines.append(f"🔻 <b>Qty Keluar:</b> -{p['qty']} pcs{satuan_info}")
        lines.append(f"📊 <b>Sisa Stok:</b> {p['sisa']} pcs")
        if p["is_kurang"]:
            lines.append("⚠️ <b>PERINGATAN: Stok sebelumnya tidak mencukupi permintaan order ini!</b>")
        elif p["sisa"] < 10:
            lines.append("⚠️ <i>Perhatian: Sisa stok menipis (&lt; 10 pcs). Segera restok!</i>")
    else:
        lines.append(f"📦 <b>Daftar Barang Keluar ({len(processed_items)} Produk):</b>")
        for idx, p in enumerate(processed_items, 1):
            satuan_info = f" ({p['satuan_asli']})" if p.get('satuan_asli') and p.get('satuan_asli') != 'pcs' else ''
            lines.append(f"{idx}. <b>[{p['sku']}]</b> {p['nama']}")
            lines.append(f"   🔻 -{p['qty']} pcs{satuan_info} | Sisa: <b>{p['sisa']} pcs</b>")
            if p["is_kurang"]:
                lines.append("   ⚠️ <i>(Stok tidak mencukupi!)</i>")
            elif p["sisa"] < 10:
                lines.append("   ⚠️ <i>(Sisa stok menipis &lt; 10 pcs)</i>")

    if unknown_items:
        lines.append(f"\n⚠️ <i>Ada {len(unknown_items)} item tidak dikenali di Master: {', '.join(unknown_items)}</i>")

    lines.append("━━━━━━━━━━━━━━━━━━━━━━")
    lines.append("✅ <i>Stok seluruh produk berhasil dipotong &amp; dicatat di Riwayat Mutasi!</i>")

    telegram_msg = "\n".join(lines)
    telegram_client.send_message(chat_id, telegram_msg, reply_to_message_id=message_id, parse_mode="HTML")
    return True, "multi_item_outflow", {"items": processed_items, "no_resi": final_no_resi}


# =============================================================================
# 8. ROUTE UTAMA FLASK (/webhook & TELEGRAM MINI APP WEB PANEL)
# =============================================================================
@app.route("/", methods=["GET"])
def health_check():
    """Endpoint untuk verifikasi server berjalan."""
    return jsonify({
        "status": "online",
        "service": "AI Agen Gudang Telegram Bot (Tira)",
        "features": [
            "Pemisahan Mode (Grup Monitoring & Resi, Japri Panel Admin)",
            "Telegram Mini App (Web Panel Visual)",
            "Koreksi Stock Opname",
            "Leaderboard Seller Hari Ini",
            "Broadcast Notifikasi Restok Otomatis ke Grup",
            "Anti-Duplicate Resi",
            "Multi-Item dalam 1 Resi",
            "Ekspor Laporan Excel (.xlsx)"
        ],
        "web_panel": f"{WEB_APP_URL}",
        "timestamp": time.time()
    }), 200


@app.route("/panel", methods=["GET"])
def web_panel():
    """Menampilkan halaman Telegram Mini App (Web Panel Visual Gudang)."""
    return render_template("panel.html")


@app.route("/api/stok", methods=["GET"])
def api_get_stok():
    """API JSON mengambil seluruh daftar stok untuk Web Panel."""
    try:
        items = sheets_manager.get_all_products_list()
        return jsonify({"status": "ok", "items": items}), 200
    except Exception as e:
        logger.error(f"Error api_get_stok: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/restok", methods=["POST"])
def api_restok():
    """API JSON restok barang masuk + memicu broadcast otomatis ke seluruh grup Telegram."""
    try:
        data = request.get_json(force=True) or {}
        sku = data.get("sku", "").strip().upper()
        qty = int(data.get("qty", 0))
        admin = data.get("admin", "Admin Web Panel").strip()

        if not sku or qty <= 0:
            return jsonify({"status": "error", "message": "SKU atau Qty tidak valid."}), 400

        nama_resmi = sheets_manager.get_nama_standar_by_sku(sku)
        current_qty, new_qty, _, _ = sheets_manager.update_or_create_stok(sku, qty, action="MASUK")
        sheets_manager.record_mutasi(
            tipe="MASUK",
            kode_sku=sku,
            nama_barang=nama_resmi,
            qty=qty,
            sisa_stok=new_qty,
            pengirim=admin,
            keterangan="Restok Pabrik via Web Panel"
        )

        broadcast_restok_to_groups(sku, nama_resmi, qty, new_qty, admin_name=admin)
        return jsonify({
            "status": "ok",
            "message": f"Berhasil restok {sku} (+{qty} pcs).",
            "new_qty": new_qty
        }), 200
    except Exception as e:
        logger.error(f"Error api_restok: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/tambah_sku", methods=["POST"])
def api_tambah_sku():
    """API JSON pendaftaran Master SKU baru dari Web Panel."""
    try:
        data = request.get_json(force=True) or {}
        sku = data.get("sku", "").strip().upper()
        nama = data.get("nama", "").strip()
        stok_awal = int(data.get("stok_awal", 0))

        if not sku or not nama:
            return jsonify({"status": "error", "message": "SKU dan nama produk harus diisi."}), 400

        success = sheets_manager.add_master_sku(sku, nama, stok_awal)
        if success:
            return jsonify({"status": "ok", "message": f"SKU {sku} berhasil didaftarkan."}), 200
        else:
            return jsonify({"status": "error", "message": f"SKU {sku} sudah terdaftar sebelumnya."}), 400
    except Exception as e:
        logger.error(f"Error api_tambah_sku: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/opname", methods=["POST"])
def api_opname():
    """API JSON koreksi stock opname dari Web Panel."""
    try:
        data = request.get_json(force=True) or {}
        sku = data.get("sku", "").strip().upper()
        diff_qty = int(data.get("diff_qty", 0))
        alasan = data.get("alasan", "Koreksi Fisik").strip()
        admin = data.get("admin", "Admin Web Panel").strip()

        if not sku or diff_qty == 0 or not alasan:
            return jsonify({"status": "error", "message": "Parameter opname tidak lengkap."}), 400

        cur_qty, new_qty, success, nama = sheets_manager.perform_stock_opname(sku, diff_qty, alasan, admin_name=admin)
        if success:
            return jsonify({
                "status": "ok",
                "message": f"Stock opname {sku} berhasil ({'+' if diff_qty >= 0 else ''}{diff_qty} pcs).",
                "current_qty": cur_qty,
                "new_qty": new_qty
            }), 200
        else:
            return jsonify({"status": "error", "message": f"SKU {sku} tidak ditemukan."}), 404
    except Exception as e:
        logger.error(f"Error api_opname: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/leaderboard", methods=["GET"])
def api_leaderboard():
    """API JSON mengambil ranking leaderboard seller hari ini."""
    try:
        data = sheets_manager.get_leaderboard_data()
        return jsonify(data), 200
    except Exception as e:
        logger.error(f"Error api_leaderboard: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/webhook", methods=["POST"])
def telegram_webhook():
    """Endpoint webhook Telegram."""
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            return jsonify({"status": "ignored", "reason": "empty_payload"}), 200

        callback_query = data.get("callback_query")
        if callback_query:
            success, reason = handle_incoming_callback_query(callback_query)
            return jsonify({"status": "ok" if success else "handled", "reason": reason}), 200

        message = data.get("message")
        if not message:
            return jsonify({"status": "ignored", "reason": "no_message"}), 200

        success, reason, extra = handle_incoming_message(message)
        return jsonify({"status": "ok" if success else "handled", "reason": reason, **extra}), 200

    except Exception as e:
        logger.exception(f"Unhandled Exception pada webhook: {e}")
        return jsonify({"status": "error", "message": str(e)}), 200


# =============================================================================
# 8. HELPER ROUTES UNTUK VERCEL (WEBHOOK MANAGEMENT & CRON JOBS)
# =============================================================================
@app.route("/set_webhook", methods=["GET", "POST"])
def set_webhook_route():
    """
    Helper 1-klik untuk mendaftarkan URL Webhook Telegram setelah deploy ke Vercel.
    Cukup buka URL ini di browser: https://<domain-vercel>/set_webhook
    """
    custom_url = request.args.get("url")
    if not custom_url:
        host = request.headers.get("X-Forwarded-Host") or request.host
        proto = request.headers.get("X-Forwarded-Proto") or "https"
        custom_url = f"{proto}://{host}/webhook"

    ok, res = telegram_client.set_webhook(custom_url)
    return jsonify({
        "status": "ok" if ok else "error",
        "webhook_url": custom_url,
        "telegram_response": res,
        "message": "✅ Webhook Telegram berhasil didaftarkan!" if ok else "❌ Gagal mendaftarkan webhook."
    }), (200 if ok else 400)


@app.route("/get_webhook_info", methods=["GET"])
@app.route("/webhook_info", methods=["GET"])
def get_webhook_info_route():
    """Melihat status webhook Telegram yang sedang aktif."""
    info = telegram_client.get_webhook_info()
    return jsonify(info), 200


@app.route("/delete_webhook", methods=["GET", "POST"])
def delete_webhook_route():
    """Menghapus webhook dari Telegram (agar bisa kembali ke mode run_polling.py lokal)."""
    ok, res = telegram_client.delete_webhook()
    return jsonify({
        "status": "ok" if ok else "error",
        "telegram_response": res,
        "message": "✅ Webhook berhasil dihapus. Bot siap untuk polling lokal!" if ok else "❌ Gagal menghapus webhook."
    }), (200 if ok else 400)


@app.route("/api/cron/briefing", methods=["GET", "POST"])
def cron_morning_briefing():
    """
    Vercel Cron endpoint: Otomatis kirim briefing pagi (08:00 WIB / 01:00 UTC) ke grup.
    Dijalankan secara serverless oleh Vercel Scheduler tanpa perlu server lokal standby.
    """
    try:
        briefing_text = sheets_manager.get_morning_briefing_text()
        group_ids = get_active_group_ids()
        sent_count = 0
        for gid in group_ids:
            if telegram_client.send_message(gid, briefing_text, parse_mode="HTML"):
                sent_count += 1
        return jsonify({
            "status": "ok",
            "job": "morning_briefing",
            "sent_to_groups": sent_count,
            "total_groups": len(group_ids)
        }), 200
    except Exception as e:
        logger.error(f"Error cron morning briefing: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/cron/recap", methods=["GET", "POST"])
def cron_evening_recap():
    """
    Vercel Cron endpoint: Otomatis kirim rekap sore tutup buku (17:00 WIB / 10:00 UTC) ke grup.
    """
    try:
        recap_text = sheets_manager.get_evening_recap_text()
        group_ids = get_active_group_ids()
        sent_count = 0
        for gid in group_ids:
            if telegram_client.send_message(gid, recap_text, parse_mode="HTML"):
                sent_count += 1
        return jsonify({
            "status": "ok",
            "job": "evening_recap",
            "sent_to_groups": sent_count,
            "total_groups": len(group_ids)
        }), 200
    except Exception as e:
        logger.error(f"Error cron evening recap: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


# Pada serverless (Vercel), background scheduler thread tidak dijalankan karena digantikan oleh Vercel Crons
if not IS_VERCEL:
    start_background_scheduler()


# =============================================================================
# 9. MAIN ENTRYPOINT
# =============================================================================
if __name__ == "__main__":
    logger.info(f"Menjalankan Warehouse AI Bot di port {PORT} (Debug={DEBUG})...")
    app.run(host="0.0.0.0", port=PORT, debug=DEBUG)
