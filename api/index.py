import os
import sys

# Tambahkan direktori root proyek ke sys.path agar modul di root dapat diimpor oleh Vercel
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app import app

# Entrypoint WSGI untuk Vercel Serverless Function
# Vercel akan otomatis mengenali objek 'app' ini sebagai Flask WSGI application.
app = app
