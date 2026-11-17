import os
import json
import requests
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")

daftar_master_sku = """- SKU: MHPD 021 | Nama: DEPHERO KAPSUL
- SKU: MHPD 003 | Nama: HABBAFIT
- SKU: MHPD 027 | Nama: G-MAX SACHET
- SKU: MHPD 004 | Nama: BIO BINAHONG
- SKU: MHPD 010 | Nama: ARVOS"""

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
    "4. Nilai 'qty' HARUS bertipe integer dalam satuan PCS setelah dikonversi.\n"
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

test_inputs = [
    "No. Resi: 11004683469987. Kirim 1 Dephero dan 2 Habbafit ya",
    "Tolong siapin 2 box G-MAX dan 1 lusin Bio Binahong"
]

candidate_models = ["gemini-flash-lite-latest", "gemini-2.5-flash-lite", "gemini-flash-latest"]

for text in test_inputs:
    print(f"\n--- Testing: '{text}' ---")
    payload = {
        "system_instruction": {"parts": [{"text": system_instruction}]},
        "contents": [{"parts": [{"text": text}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.1
        }
    }
    success = False
    for model in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        try:
            resp = requests.post(url, json=payload, timeout=25)
            if resp.status_code == 200:
                data = resp.json()
                raw_json = data["candidates"][0]["content"]["parts"][0]["text"]
                parsed = json.loads(raw_json)
                print(f"Success with {model}!\nResult:", json.dumps(parsed, indent=2))
                success = True
                break
            else:
                print(f"Model {model} returned status {resp.status_code}")
        except Exception as e:
            print(f"Model {model} error: {e}")
    if not success:
        print("Failed all models for:", text)
