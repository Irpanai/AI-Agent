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
    WSGI Middleware untuk memetakan path query Vercel rewrite ke PATH_INFO Flask.
    Ketika Vercel melakukan rewrite URL /set_webhook -> /api/index?path=set_webhook,
    middleware ini memulihkan PATH_INFO menjadi '/set_webhook' sehingga routing Flask
    berjalan 100% akurat.
    """
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        qs = environ.get("QUERY_STRING", "")
        if qs:
            parsed = urllib.parse.parse_qs(qs, keep_blank_values=True)
            if "path" in parsed:
                path_list = parsed.pop("path")
                path_val = path_list[0] if path_list else ""
                if path_val:
                    if not path_val.startswith("/"):
                        path_val = "/" + path_val
                    environ["PATH_INFO"] = path_val
                else:
                    environ["PATH_INFO"] = "/"

                environ["QUERY_STRING"] = urllib.parse.urlencode(parsed, doseq=True)

        return self.wsgi_app(environ, start_response)


# Pasang middleware ke objek Flask
app.wsgi_app = VercelPathFixMiddleware(app.wsgi_app)
