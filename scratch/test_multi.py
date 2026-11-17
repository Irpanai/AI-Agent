import requests
import json
import os
from dotenv import load_dotenv

load_dotenv()
key = os.getenv("GEMINI_API_KEY")
url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-lite-latest:generateContent?key={key}"

sys_prompt = """Kamu adalah sistem admin gudang ahli. Normalisasi pesanan/resi barang ke format JSON.
Daftar Master SKU:
- SKU: MHPD 021 | Nama: DEPHERO
- SKU: MHPD 003 | Nama: HABBAFIT
- SKU: MHPD 027 | Nama: G-MAX SACHET
- SKU: MHPD 004 | Nama: BIO BINAHONG

ATURAN KONVERSI SATUAN KEMASAN KE PCS:
- 1 Lusin = 12 pcs
- 1 Kodi = 20 pcs
- 1 Box/Kotak G-MAX = 10 Sachet (MHPD 027)
- 1 Box/Kotak V-MAX = 10 Sachet (MHPD 028)
- 1 Dus/Karton = 24 pcs

TUGAS:
1. Ekstrak 'no_resi' dari dokumen/foto/pesan jika ada (string kosong jika tidak ada).
2. Ekstrak SEMUA barang dalam pesanan (bisa 1 barang atau lebih / multi-item).
3. Hitung 'qty' akhir dalam satuan PCS (kalikan jika ada satuan box/lusin/dus).
4. Kembalikan HANYA JSON:
{
  "no_resi": "...",
  "items": [
    {
      "kode_sku": "...",
      "nama_standar": "...",
      "qty": 1,
      "satuan_asli": "pcs/box/lusin",
      "nama_asli_seller": "..."
    }
  ]
}
"""

user_text = "No. Resi: 11004683469987. Tolong dephero 1 pcs, 2 box g-max, dan 1 lusin bio binahong ya."

payload = {
    "system_instruction": {"parts": [{"text": sys_prompt}]},
    "contents": [{"parts": [{"text": user_text}]}],
    "generationConfig": {"responseMimeType": "application/json", "temperature": 0.1}
}

r = requests.post(url, json=payload, timeout=20)
print("Status:", r.status_code)
print("Response JSON:\n", r.json()["candidates"][0]["content"]["parts"][0]["text"])
