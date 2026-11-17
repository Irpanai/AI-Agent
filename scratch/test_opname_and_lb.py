import os
import sys
import re
from datetime import datetime
from dotenv import load_dotenv

sys.path.insert(0, r"D:\AI Agent")
load_dotenv()

from app import sheets_manager

# Test Opname logic calculation
def test_opname():
    ws_stok = sheets_manager.get_worksheet("Stok_Barang")
    records = ws_stok.get_all_values()
    print("Found", len(records), "records in Stok_Barang.")
    if len(records) > 1:
        first_sku = records[1][0]
        cur_qty = int(records[1][1])
        print(f"Sample SKU: {first_sku}, Current Qty: {cur_qty}")
        diff = -2
        new_qty = max(0, cur_qty + diff)
        print(f"Opname sim: {cur_qty} + ({diff}) = {new_qty}")

test_opname()
