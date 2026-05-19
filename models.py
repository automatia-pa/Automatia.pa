import sqlite3
import threading
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

import os
DB_PATH = os.path.expanduser("~/.private_data/users.db")

_local = threading.local()

def get_conn():
    """Retorna una conexión SQLite dedicada al thread actual."""
    if not hasattr(_local, 'conn') or _local.conn is None:
        _local.conn = sqlite3.connect(DB_PATH, check_same_thread=True)
        _local.conn.execute("PRAGMA journal_mode=WAL")
        _local.conn.execute("PRAGMA foreign_keys=ON")
    return _local.conn

def db_init():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre      TEXT    NOT NULL,
            email       TEXT    UNIQUE NOT NULL COLLATE NOCASE,
            password    TEXT    NOT NULL,
            totp_secret TEXT    DEFAULT NULL,
            created_at  TEXT    DEFAULT (datetime('now'))
        )
    """)
    conn.commit()

    # Migración: agrega la columna si la DB ya existía sin ella
    try:
        conn.execute("ALTER TABLE users ADD COLUMN totp_secret TEXT DEFAULT NULL")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # columna ya existe, OK


class User(UserMixin):
    def __init__(self, id, nombre, email):
        self.id     = id
        self.nombre = nombre
        self.email  = email

    # ── CRUD básico ──────────────────────────────────────────

    @staticmethod
    def create(nombre, email, password):
        conn = get_conn()
        hashed = generate_password_hash(password)
        try:
            conn.execute(
                "INSERT INTO users (nombre, email, password) VALUES (?, ?, ?)",
                (nombre, email.lower(), hashed)
            )
            conn.commit()
            user_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            return User(user_id, nombre, email)
        except sqlite3.IntegrityError:
            return None

    @staticmethod
    def get(user_id):
        conn = get_conn()
        row = conn.execute(
            "SELECT id, nombre, email FROM users WHERE id=?", (user_id,)
        ).fetchone()
        return User(*row) if row else None

    @staticmethod
    def get_by_email(email):
        conn = get_conn()
        row = conn.execute(
            "SELECT id, nombre, email FROM users WHERE email=?", (email.lower(),)
        ).fetchone()
        return User(*row) if row else None

    @staticmethod
    def exists(email):
        conn = get_conn()
        row = conn.execute(
            "SELECT 1 FROM users WHERE email=?", (email.lower(),)
        ).fetchone()
        return row is not None

    @staticmethod
    def check_password(email, password):
        """Verifica credenciales con tiempo constante (anti timing-attack)."""
        conn = get_conn()
        row = conn.execute(
            "SELECT password FROM users WHERE email=?", (email.lower(),)
        ).fetchone()
        if row:
            return check_password_hash(row[0], password)
        check_password_hash("$dummy$", password)
        return False

    # ── MFA / TOTP ───────────────────────────────────────────

    @staticmethod
    def set_totp(user_id, secret):
        """
        Activa o desactiva TOTP para el usuario.
        Pasa secret=None para desactivar.
        """
        conn = get_conn()
        conn.execute(
            "UPDATE users SET totp_secret=? WHERE id=?",
            (secret, user_id)
        )
        conn.commit()

    @staticmethod
    def get_totp_secret(user_id):
        """Retorna el secret TOTP del usuario, o None si no tiene MFA."""
        conn = get_conn()
        row = conn.execute(
            "SELECT totp_secret FROM users WHERE id=?", (user_id,)
        ).fetchone()
        return row[0] if row else None

    @staticmethod
    def has_totp(user_id) -> bool:
        """True si el usuario tiene MFA configurado."""
        return bool(User.get_totp_secret(user_id))

    # ── RECUPERACIÓN DE CONTRASEÑA ────────────────────────────

    @staticmethod
    def _init_reset_table(conn):
        conn.execute("""
            CREATE TABLE IF NOT EXISTS password_reset_tokens (
                token      TEXT PRIMARY KEY,
                user_id    INTEGER NOT NULL,
                expires_at TEXT    NOT NULL,
                used       INTEGER DEFAULT 0
            )
        """)
        conn.commit()

    @staticmethod
    def create_reset_token(email: str) -> str | None:
        """
        Crea un token de recuperación válido por 1 hora.
        Retorna el token si el email existe, None si no.
        """
        import secrets as _secrets
        from datetime import datetime, timedelta
        conn = get_conn()
        User._init_reset_table(conn)
        row = conn.execute(
            "SELECT id FROM users WHERE email=?", (email.lower(),)
        ).fetchone()
        if not row:
            return None
        user_id = row[0]
        # Invalidar tokens anteriores del mismo usuario
        conn.execute(
            "DELETE FROM password_reset_tokens WHERE user_id=?", (user_id,)
        )
        token = _secrets.token_urlsafe(32)
        expires = (datetime.now() + timedelta(hours=1)).isoformat()
        conn.execute(
            "INSERT INTO password_reset_tokens (token, user_id, expires_at) VALUES (?, ?, ?)",
            (token, user_id, expires)
        )
        conn.commit()
        return token

    @staticmethod
    def validate_reset_token(token: str) -> int | None:
        """
        Valida el token. Retorna user_id si es válido y no expiró, None si no.
        """
        from datetime import datetime
        conn = get_conn()
        User._init_reset_table(conn)
        row = conn.execute(
            "SELECT user_id, expires_at, used FROM password_reset_tokens WHERE token=?",
            (token,)
        ).fetchone()
        if not row:
            return None
        user_id, expires_at, used = row
        if used:
            return None
        if datetime.now().isoformat() > expires_at:
            return None
        return user_id

    @staticmethod
    def consume_reset_token(token: str, new_password: str) -> bool:
        """
        Cambia la contraseña y marca el token como usado.
        Retorna True si tuvo éxito.
        """
        user_id = User.validate_reset_token(token)
        if not user_id:
            return False
        conn = get_conn()
        hashed = generate_password_hash(new_password)
        conn.execute("UPDATE users SET password=? WHERE id=?", (hashed, user_id))
        conn.execute(
            "UPDATE password_reset_tokens SET used=1 WHERE token=?", (token,)
        )
        conn.commit()
        return True
