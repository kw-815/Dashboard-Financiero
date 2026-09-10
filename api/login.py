"""
Login real del dashboard.

Verifica usuario/contraseña contra AUTH_USERS (variable de entorno de Vercel,
nunca en el código) y, si son correctos, entrega una cookie de sesión firmada
(JWT HS256) que `middleware.js` valida en cada visita antes de servir
index.html. La contraseña nunca se guarda como tal, solo su hash (scrypt).

Variables de entorno requeridas en Vercel:
  AUTH_USERS      JSON: {"usuario": {"hash": "<salt_hex>:<hash_hex>", "role": "admin"}, ...}
  AUTH_JWT_SECRET string aleatorio largo (ej. `openssl rand -hex 32`)
"""
import base64
import hashlib
import hmac
import json
import os
import time
from http.server import BaseHTTPRequestHandler

SCRYPT_N = 16384
SCRYPT_R = 8
SCRYPT_P = 1
DKLEN = 64
SESSION_SECONDS = 60 * 60 * 12  # 12 horas


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _verify_password(password: str, stored_hash: str) -> bool:
    try:
        salt_hex, hash_hex = stored_hash.split(":")
    except ValueError:
        return False
    salt = bytes.fromhex(salt_hex)
    expected = bytes.fromhex(hash_hex)
    derived = hashlib.scrypt(
        password.encode("utf-8"), salt=salt, n=SCRYPT_N, r=SCRYPT_R, p=SCRYPT_P, dklen=DKLEN
    )
    return hmac.compare_digest(derived, expected)


def _sign_jwt(payload: dict, secret: str) -> str:
    header = {"alg": "HS256", "typ": "JWT"}
    header_b64 = _b64url(json.dumps(header, separators=(",", ":")).encode())
    payload_b64 = _b64url(json.dumps(payload, separators=(",", ":")).encode())
    signing_input = f"{header_b64}.{payload_b64}".encode()
    signature = hmac.new(secret.encode("utf-8"), signing_input, hashlib.sha256).digest()
    return f"{header_b64}.{payload_b64}.{_b64url(signature)}"


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b"{}"
        try:
            body = json.loads(raw or b"{}")
        except json.JSONDecodeError:
            body = {}

        username = str(body.get("username", "")).strip()
        password = str(body.get("password", ""))

        users_raw = os.environ.get("AUTH_USERS", "{}")
        secret = os.environ.get("AUTH_JWT_SECRET", "")
        try:
            users = json.loads(users_raw)
        except json.JSONDecodeError:
            users = {}

        user = users.get(username)
        ok = bool(user and secret and _verify_password(password, user.get("hash", "")))

        if not ok:
            self.send_response(401)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": "Usuario o contraseña incorrectos."}).encode())
            return

        now = int(time.time())
        payload = {
            "sub": username,
            "role": user.get("role", "user"),
            "iat": now,
            "exp": now + SESSION_SECONDS,
        }
        token = _sign_jwt(payload, secret)

        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header(
            "Set-Cookie",
            f"kw_session={token}; Path=/; Max-Age={SESSION_SECONDS}; HttpOnly; Secure; SameSite=Lax",
        )
        self.end_headers()
        self.wfile.write(json.dumps({"ok": True}).encode())
