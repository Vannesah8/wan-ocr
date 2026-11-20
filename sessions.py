"""Accounts and revocable browser sessions for the local wan-ocr server."""
import hashlib
import hmac
import re
import secrets
import sqlite3
import time
from contextlib import contextmanager
from http.cookies import CookieError, SimpleCookie

LIFETIME = 30 * 24 * 3600
COOKIE = "wan_session"


@contextmanager
def database(root):
    root.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(root / "accounts.sqlite3", timeout=10)
    try:
        with connection:
            connection.execute("CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, email TEXT UNIQUE, salt BLOB, password BLOB)")
            connection.execute("CREATE TABLE IF NOT EXISTS sessions (digest BLOB PRIMARY KEY, user_id INTEGER, expires INTEGER)")
            yield connection
    finally:
        connection.close()


def digest(token):
    return hashlib.sha256(token.encode()).digest()


def token_from(headers):
    cookies = SimpleCookie()
    try:
        cookies.load(headers.get("Cookie", ""))
    except CookieError:
        return ""
    return cookies[COOKIE].value if COOKIE in cookies else ""


def identify(root, token):
    if not token:
        return None
    with database(root) as db:
        return db.execute("SELECT users.id, email FROM sessions JOIN users ON users.id = sessions.user_id "
                          "WHERE digest=? AND expires>?", (digest(token), time.time())).fetchone()


def authenticate(root, email, password, signup):
    if not isinstance(email, str) or not isinstance(password, str):
        raise ValueError("Enter your email and password.")
    email = email.strip().casefold()
    if len(email) > 254 or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
        raise ValueError("Enter a valid email address.")
    if len(password) > 256 or (signup and len(password) < 12):
        raise ValueError("Use a password between 12 and 256 characters.")
    with database(root) as db:
        row = db.execute("SELECT id,salt,password FROM users WHERE email=?", (email,)).fetchone()
        salt = row[1] if row and not signup else secrets.token_bytes(16)
        hashed = hashlib.scrypt(password.encode(), salt=salt, n=2**17, r=8, p=1, maxmem=256*1024**2)
        if signup:
            try:
                user = db.execute("INSERT INTO users(email,salt,password) VALUES (?,?,?)",
                                  (email, salt, hashed)).lastrowid
            except sqlite3.IntegrityError as exc:
                raise ValueError("Unable to create this account. Try signing in.") from exc
        elif row and hmac.compare_digest(row[2], hashed):
            user = row[0]
        else:
            raise ValueError("Incorrect email or password.")
        token = secrets.token_urlsafe(32)
        db.execute("DELETE FROM sessions WHERE expires<=?", (time.time(),))
        db.execute("INSERT INTO sessions VALUES (?,?,?)", (digest(token), user, int(time.time()) + LIFETIME))
        return token


def logout(root, token):
    with database(root) as db:
        db.execute("DELETE FROM sessions WHERE digest=?", (digest(token),))
