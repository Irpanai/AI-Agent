import os
import sys
import urllib.parse

# Tambahkan direktori root proyek ke sys.path agar modul di root dapat diimpor oleh Vercel
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app import app


class VercelPathFixMiddleware:
    """
    WSGI Middleware untuk memetakan path request Vercel rewrite ke PATH_INFO Flask.
    Di Vercel Serverless Python, URL asli yang dikunjungi user dikirimkan via header
    X-Matched-Path (HTTP_X_MATCHED_PATH) oleh reverse proxy Vercel.
    """
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        # 1. Cek header Vercel untuk URL / path asli yang diminta user
        matched_path = (
            environ.get("HTTP_X_MATCHED_PATH")
            or environ.get("HTTP_X_VERCEL_MATCHED_PATH")
            or environ.get("HTTP_X_NOW_ROUTE_MATCHES")
            or environ.get("HTTP_X_FORWARDED_URI")
            or environ.get("REQUEST_URI")
            or environ.get("RAW_URI")
        )

        # 2. Cek query parameter 'path' sebagai fallback
        qs = environ.get("QUERY_STRING", "")
        if not matched_path and qs:
            parsed = urllib.parse.parse_qs(qs, keep_blank_values=True)
            if "path" in parsed:
                p = parsed.get("path")
                if p and p[0]:
                    matched_path = p[0]

        if matched_path:
            # Pisahkan query string jika terselip di dalam matched_path
            if "?" in matched_path:
                clean_path, extracted_qs = matched_path.split("?", 1)
                if not environ.get("QUERY_STRING"):
                    environ["QUERY_STRING"] = extracted_qs
            else:
                clean_path = matched_path

            if not clean_path.startswith("/"):
                clean_path = "/" + clean_path

            # Jika path menunjuk ke file handler /api/index, itu adalah akses ke root '/'
            if clean_path in ("/api/index", "/api/index.py"):
                clean_path = "/"

            environ["PATH_INFO"] = clean_path

        return self.wsgi_app(environ, start_response)


# Pasang middleware ke objek WSGI Flask
app.wsgi_app = VercelPathFixMiddleware(app.wsgi_app)
