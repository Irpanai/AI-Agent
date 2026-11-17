import os
import sys

# Tambahkan direktori root proyek ke sys.path agar modul di root dapat diimpor oleh Vercel
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app import app


class VercelPathFixMiddleware:
    """
    WSGI Middleware untuk memulihkan PATH_INFO asli.
    Saat Vercel melakukan internal rewrite ke /api/index, header 'x-matched-path' atau
    'REQUEST_URI' menyimpan path asli yang dikunjungi pengguna (misal: /set_webhook).
    """
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        matched_path = (
            environ.get("HTTP_X_MATCHED_PATH") or
            environ.get("REQUEST_URI", "").split("?")[0] or
            environ.get("RAW_URI", "").split("?")[0]
        )

        current_path = environ.get("PATH_INFO", "")
        if matched_path and current_path in ["/api/index", "/api", "/api/index.py"]:
            environ["PATH_INFO"] = matched_path

        return self.wsgi_app(environ, start_response)


# Pasang middleware ke objek Flask
app.wsgi_app = VercelPathFixMiddleware(app.wsgi_app)
