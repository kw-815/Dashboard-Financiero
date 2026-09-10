"""Cierra la sesión: borra la cookie kw_session."""
from http.server import BaseHTTPRequestHandler


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header(
            "Set-Cookie",
            "kw_session=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax",
        )
        self.end_headers()
        self.wfile.write(b'{"ok": true}')
