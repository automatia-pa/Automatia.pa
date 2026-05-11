import sqlite3
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

DB_PATH = "users.db"

def get_conn():
    # FIX: check_same_thread=False para que Flask pueda usarlo en múltiples requests
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def db_init():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()

class User(UserMixin):
    def __init__(self, id, nombre, email):
        self.id = id
        self.nombre = nombre
        self.email = email

    @staticmethod
    def create(nombre, email, password):
        conn = get_conn()
        hashed = generate_password_hash(password)
        conn.execute("INSERT INTO users (nombre, email, password) VALUES (?, ?, ?)",
                     (nombre, email, hashed))
        conn.commit()
        user_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.close()
        return User(user_id, nombre, email)

    @staticmethod
    def get(user_id):
        conn = get_conn()
        row = conn.execute("SELECT id, nombre, email FROM users WHERE id=?", (user_id,)).fetchone()
        conn.close()
        return User(*row) if row else None

    @staticmethod
    def get_by_email(email):
        conn = get_conn()
        row = conn.execute("SELECT id, nombre, email FROM users WHERE email=?", (email,)).fetchone()
        conn.close()
        return User(*row) if row else None

    @staticmethod
    def exists(email):
        conn = get_conn()
        row = conn.execute("SELECT 1 FROM users WHERE email=?", (email,)).fetchone()
        conn.close()
        return row is not None

    @staticmethod
    def check_password(email, password):
        conn = get_conn()
        row = conn.execute("SELECT password FROM users WHERE email=?", (email,)).fetchone()
        conn.close()
        if row:
            return check_password_hash(row[0], password)
        return False