import sqlite3
import threading
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

DB_PATH = "users.db"

# ── Pool de conexiones por thread (thread-local storage) ──────
# check_same_thread=False es inseguro en producción si se comparte una misma
# conexión entre threads. La solución correcta es una conexión por thread.
_local = threading.local()

def get_conn():
    """Retorna una conexión SQLite dedicada al thread actual."""
    if not hasattr(_local, 'conn') or _local.conn is None:
        _local.conn = sqlite3.connect(DB_PATH, check_same_thread=True)
        _local.conn.execute("PRAGMA journal_mode=WAL")   # mejor concurrencia
        _local.conn.execute("PRAGMA foreign_keys=ON")
    return _local.conn

def db_init():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre   TEXT    NOT NULL,
            email    TEXT    UNIQUE NOT NULL COLLATE NOCASE,
            password TEXT    NOT NULL,
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()


class User(UserMixin):
    def __init__(self, id, nombre, email):
        self.id     = id
        self.nombre = nombre
        self.email  = email

    @staticmethod
    def create(nombre, email, password):
        conn = get_conn()
        # pbkdf2:sha256 con 600 000 iteraciones (werkzeug >= 3 usa scrypt por defecto)
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
        """Verifica credenciales. Siempre ejecuta check_password_hash para
        evitar timing attacks (no retorna False prematuramente si el user
        no existe)."""
        conn = get_conn()
        row = conn.execute(
            "SELECT password FROM users WHERE email=?", (email.lower(),)
        ).fetchone()
        if row:
            return check_password_hash(row[0], password)
        # Ejecutar hash de todas formas para igualar tiempo de respuesta
        check_password_hash("$dummy$", password)
        return False
