import os
import sys
import io
from datetime import datetime, date, timedelta
from dotenv import load_dotenv

load_dotenv()

from app import sheets_manager, TELEGRAM_TOKEN, telegram_client

print("Testing Google Sheets connection...")
stok_overview = sheets_manager.get_stok_overview()
print("Stok overview preview:")
print(stok_overview[:200])

# Test duplicate resi check
print("\nTesting duplicate resi check...")
# Let's check a non-existent resi
res = sheets_manager.check_duplicate_resi("NON_EXISTENT_12345") if hasattr(sheets_manager, 'check_duplicate_resi') else None
print("Check duplicate result:", res)
