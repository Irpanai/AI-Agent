import os
import json
import requests
import pypdf
import io
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")

pdf_path = r"C:/Users/Irpanai/.gemini/antigravity/brain/dd0ef018-aaa8-4acf-a726-aa675b7d5802/.user_uploaded/media_1794808257866_451f5cec.pdf"

with open(pdf_path, "rb") as f:
    pdf_bytes = f.read()

reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
texts = [p.extract_text() for p in reader.pages if p.extract_text()]
full_text = "\n".join(texts)

daftar_master_sku = "- SKU: MHPD 021 | Nama: DEPHERO KAPSUL\n- SKU: MHPD 003 | Nama: HABBAFIT"

system_instruction = f"""Kamu adalah sistem AI admin logistik gudang ahli di Indonesia. Normalisasi pesanan/resi barang masuk ke format JSON standar.
DAFTAR MASTER SKU RESMI:
{daftar_master_sku}

ATURAN:
1. Ekstrak 'no_resi' dari teks/dokumen resi jika ada (misal nomor resi 11004683469987).
2. Ekstrak SEMUA barang dalam pesanan ke dalam list 'items'.
3. Cocokkan setiap barang ke SKU terdekat.
4. Output JSON:
{{
  "no_resi": "...",
  "items": [
    {{
      "kode_sku": "...",
      "nama_standar": "...",
      "qty": 1,
      "satuan_asli": "pcs",
      "nama_asli_seller": "..."
    }}
  ]
}}
"""

payload = {
    "system_instruction": {"parts": [{"text": system_instruction}]},
    "contents": [{"parts": [{"text": f"Isi Dokumen Resi:\n{full_text}"}]}],
    "generationConfig": {"responseMimeType": "application/json", "temperature": 0.1}
}

r = requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-lite-latest:generateContent?key={api_key}", json=payload, timeout=25)
print("Status:", r.status_code)
print(r.json()["candidates"][0]["content"]["parts"][0]["text"])
