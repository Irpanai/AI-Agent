import os
import sys
from datetime import datetime
from dotenv import load_dotenv

sys.path.insert(0, r"D:\AI Agent")
load_dotenv()

from app import sheets_manager

print("Testing Leaderboard generator...")
lb = sheets_manager.get_leaderboard_text() if hasattr(sheets_manager, "get_leaderboard_text") else "Method not yet added"
print("Leaderboard result:\n", lb)
