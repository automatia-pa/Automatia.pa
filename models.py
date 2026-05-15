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
