import os
import sys
from dotenv import load_dotenv

sys.path.insert(0, r"D:\AI Agent")
load_dotenv()

from app import sheets_manager

items = sheets_manager.get_all_products_list()
print(f"Total items in inventory: {len(items)}")

page_size = 10
total_pages = (len(items) + page_size - 1) // page_size
print(f"Total pages: {total_pages}")

# Test page 1
page = 1
start_idx = (page - 1) * page_size
end_idx = start_idx + page_size
page_items = items[start_idx:end_idx]
print(f"Page {page} items count: {len(page_items)}")
for it in page_items[:3]:
    print(f" - [{it['sku']}] {it['nama']} -> {it['qty']} pcs")
