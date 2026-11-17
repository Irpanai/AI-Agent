import os
import time
import requests
from dotenv import load_dotenv

load_dotenv()
token = os.getenv("TELEGRAM_TOKEN")

for attempt in range(1, 4):
    try:
        print(f"Attempt {attempt} calling getMe...")
        r = requests.get(f"https://api.telegram.org/bot{token}/getMe", timeout=30)
        print("Response:", r.status_code, r.json())
        break
    except Exception as e:
        print(f"Attempt {attempt} error: {e}")
        time.sleep(2)
