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
texts = []
for p in reader.pages:
    t = p.extract_text()
    if t:
        texts.append(t)
full_text = "\n".join(texts)
print(f"Extracted {len(full_text)} chars from PDF.")
print("Preview:\n", full_text[:400])

from scratch.test_gemini_complete import system_instruction

payload = {
    "system_instruction": {"parts": [{"text": system_instruction}]},
    "contents": [{"parts": [{"text": f"Isi Dokumen Resi:\n{full_text}"}]}],
    "generationConfig": {"responseMimeType": "application/json", "temperature": 0.1}
}

url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-lite-latest:generateContent?key={api_key}"
resp = requests.post(url, json=payload, timeout=25)
print("Status:", resp.status_code)
if resp.status_code == 200:
    res = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
    print("AI Parsed Result:\n", json.dumps(json.loads(res), indent=2))
else:
    print(resp.text)
