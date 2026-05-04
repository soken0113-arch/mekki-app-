import os
from datetime import datetime
from werkzeug.security import generate_password_hash

DATABASE_URL = os.environ.get("DATABASE_URL")
USE_PG = bool(DATABASE_URL)

if USE_PG:
    import psycopg2
    import psycopg2.extras
    import psycopg2.pool
    _db_pool = None

    def _get_pool():
        global _db_pool
        if _db_pool is None:
            _db_pool = psycopg2.pool.ThreadedConnectionPool(minconn=1, maxconn=4, dsn=DATABASE_URL)
        return _db_pool

    class _Conn:
        def __init__(self):
            self._pool = _get_pool()
            self._conn = self._pool.getconn()

        def execute(self, sql, params=()):
            sql = sql.replace("?", "%s")
            cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(sql, params)
            return cur

        def __enter__(self):
            return self

        def __exit__(self, exc_type, *_):
            if exc_type:
                self._conn.rollback()
                self._pool.putconn(self._conn, close=True)
            else:
                self._conn.commit()
                self._pool.putconn(self._conn)

else:
    import sqlite3

    class _Conn:
        def __init__(self):
            self._conn = sqlite3.connect(
                os.path.join(os.path.dirname(__file__), "orders.db")
            )
            self._conn.row_factory = sqlite3.Row

        def execute(self, sql, params=()):
            # SERIALをINTEGERに変換
            sql = sql.replace("SERIAL PRIMARY KEY", "INTEGER PRIMARY KEY AUTOINCREMENT")
            sql = sql.replace("BOOLEAN", "INTEGER")
            return self._conn.execute(sql, params)

        def __enter__(self):
            return self

        def __exit__(self, exc_type, *_):
            if exc_type:
                self._conn.rollback()
            else:
                self._conn.commit()
            self._conn.close()


def get_db():
    return _Conn()


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id SERIAL PRIMARY KEY,
                order_no TEXT NOT NULL,
                customer TEXT NOT NULL,
                product TEXT NOT NULL,
                part_no TEXT NOT NULL DEFAULT '',
                material TEXT NOT NULL DEFAULT '',
                quantity INTEGER NOT NULL,
                mekki_type TEXT NOT NULL,
                mekki_thickness TEXT NOT NULL DEFAULT '',
                thickness_data TEXT NOT NULL DEFAULT '不要',
                due_date TEXT NOT NULL,
                unit_price TEXT NOT NULL DEFAULT '',
                mekki_line TEXT NOT NULL DEFAULT '',
                process_note TEXT NOT NULL DEFAULT '',
                shipping_method TEXT NOT NULL DEFAULT '',
                note TEXT,
                assigned_to TEXT NOT NULL DEFAULT '',
                subcontractor_id INTEGER,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS customers (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                part_no TEXT NOT NULL DEFAULT '',
                unit_price TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                must_change_password INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS order_items (
                id SERIAL PRIMARY KEY,
                order_id INTEGER NOT NULL,
                part_no TEXT NOT NULL DEFAULT '',
                product TEXT NOT NULL,
                material TEXT NOT NULL DEFAULT '',
                quantity INTEGER NOT NULL,
                unit_price TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT ''
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS shipments (
                id SERIAL PRIMARY KEY,
                order_id INTEGER NOT NULL,
                shipped_at TEXT NOT NULL,
                shipped_by TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS subcontractors (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL
            )
        """)
        if not conn.execute("SELECT id FROM users WHERE username='admin'").fetchone():
            conn.execute(
                "INSERT INTO users (username, password_hash, must_change_password, created_at) VALUES (?, ?, ?, ?)",
                ("admin", generate_password_hash("admin1234", method="pbkdf2:sha256"), 1,
                 datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            )
