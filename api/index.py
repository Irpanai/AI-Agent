import os
import sys
import urllib.parse

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app import app


class VercelPathFixMiddleware:
    """
    WSGI Middleware untuk memetakan path request Vercel rewrite ke PATH_INFO Flask.
    Parameter __route__ dikirim via vercel.json rewrites: /api/index?__route__=/$1
    """
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        qs = environ.get("QUERY_STRING", "")
        route_path = None

        if qs:
            parsed = urllib.parse.parse_qs(qs, keep_blank_values=True)
            if "__route__" in parsed:
                raw_route = parsed.pop("__route__")[0]
                if raw_route and raw_route != "/$1":
                    route_path = "/" + raw_route.lstrip("/")
                environ["QUERY_STRING"] = urllib.parse.urlencode(parsed, doseq=True)

        # Fallback jika __route__ tidak ada
        if not route_path:
            fallback = (
                environ.get("HTTP_X_MATCHED_PATH")
                or environ.get("HTTP_X_VERCEL_MATCHED_PATH")
                or environ.get("HTTP_X_FORWARDED_URI")
            )
            if fallback and fallback not in ("/api/index", "/api/index.py"):
                clean = fallback.split("?")[0]
                route_path = "/" + clean.lstrip("/")

        if route_path:
            environ["PATH_INFO"] = route_path

        return self.wsgi_app(environ, start_response)


app.wsgi_app = VercelPathFixMiddleware(app.wsgi_app)
