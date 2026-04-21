"""
SQLite (orders.db) のデータを PostgreSQL に移行するスクリプト
使い方:
  DATABASE_URL=postgresql://... python migrate_to_postgres.py
  SQLITE_PATH=orders.db DATABASE_URL=postgresql://... python migrate_to_postgres.py
"""
import sqlite3
import os
import psycopg2
import psycopg2.extras
from datetime import datetime
from werkzeug.security import generate_password_hash

SQLITE_PATH = os.environ.get("SQLITE_PATH", "orders.db")
DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    raise SystemExit("ERROR: DATABASE_URL 環境変数を設定してください")

sqlite_conn = sqlite3.connect(SQLITE_PATH)
sqlite_conn.row_factory = sqlite3.Row
pg_conn = psycopg2.connect(DATABASE_URL)
pg_cur = pg_conn.cursor()

# ── 1. テーブル作成 ──────────────────────────────────────────

pg_cur.execute("""
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
        note TEXT,
        created_at TEXT NOT NULL
    )
""")
pg_cur.execute("""
    CREATE TABLE IF NOT EXISTS customers (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL
    )
""")
pg_cur.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL
    )
""")
pg_cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
""")
pg_conn.commit()
print("テーブル作成完了")

# ── 2. データ移行 ────────────────────────────────────────────

now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# orders（旧カラムのみの場合に新カラムをデフォルト値で補完）
orders = sqlite_conn.execute("SELECT * FROM orders").fetchall()
if orders:
    sqlite_cols = set(orders[0].keys())
    for row in orders:
        pg_cur.execute("""
            INSERT INTO orders (
                id, order_no, customer, product, part_no, material, quantity,
                mekki_type, mekki_thickness, thickness_data, due_date,
                unit_price, mekki_line, process_note, note, created_at
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON CONFLICT (id) DO NOTHING
        """, (
            row["id"],
            row["order_no"],
            row["customer"],
            row["product"],
            row["part_no"] if "part_no" in sqlite_cols else "",
            row["material"] if "material" in sqlite_cols else "",
            row["quantity"],
            row["mekki_type"],
            row["mekki_thickness"] if "mekki_thickness" in sqlite_cols else "",
            row["thickness_data"] if "thickness_data" in sqlite_cols else "不要",
            row["due_date"],
            row["unit_price"] if "unit_price" in sqlite_cols else "",
            row["mekki_line"] if "mekki_line" in sqlite_cols else "",
            row["process_note"] if "process_note" in sqlite_cols else "",
            row["note"],
            row["created_at"],
        ))
    pg_cur.execute("SELECT setval(pg_get_serial_sequence('orders', 'id'), COALESCE(MAX(id), 1)) FROM orders")
    print(f"orders: {len(orders)} 件を移行しました")
else:
    print("orders: データなし、スキップ")

# customers
customers = sqlite_conn.execute("SELECT * FROM customers").fetchall()
if customers:
    for row in customers:
        pg_cur.execute(
            "INSERT INTO customers (id, name, created_at) VALUES (%s, %s, %s) ON CONFLICT DO NOTHING",
            (row["id"], row["name"], row["created_at"])
        )
    pg_cur.execute("SELECT setval(pg_get_serial_sequence('customers', 'id'), COALESCE(MAX(id), 1)) FROM customers")
    print(f"customers: {len(customers)} 件を移行しました")
else:
    print("customers: データなし、スキップ")

# products
products = sqlite_conn.execute("SELECT * FROM products").fetchall()
if products:
    for row in products:
        pg_cur.execute(
            "INSERT INTO products (id, name, created_at) VALUES (%s, %s, %s) ON CONFLICT DO NOTHING",
            (row["id"], row["name"], row["created_at"])
        )
    pg_cur.execute("SELECT setval(pg_get_serial_sequence('products', 'id'), COALESCE(MAX(id), 1)) FROM products")
    print(f"products: {len(products)} 件を移行しました")
else:
    print("products: データなし、スキップ")

# users（SQLiteのユーザーをそのまま移行。未登録なら admin を作成）
users = sqlite_conn.execute("SELECT * FROM users").fetchall()
if users:
    for row in users:
        pg_cur.execute(
            "INSERT INTO users (id, username, password_hash, created_at) VALUES (%s, %s, %s, %s) ON CONFLICT DO NOTHING",
            (row["id"], row["username"], row["password_hash"], row["created_at"])
        )
    pg_cur.execute("SELECT setval(pg_get_serial_sequence('users', 'id'), COALESCE(MAX(id), 1)) FROM users")
    print(f"users: {len(users)} 件を移行しました")
else:
    # デフォルト admin ユーザーを作成
    pg_cur.execute(
        "INSERT INTO users (username, password_hash, created_at) VALUES (%s, %s, %s) ON CONFLICT DO NOTHING",
        ("admin", generate_password_hash("admin1234", method="pbkdf2:sha256"), now)
    )
    print("users: admin ユーザーを新規作成しました")

pg_conn.commit()
pg_conn.close()
sqlite_conn.close()
print("\n移行完了！")
print(f"次のコマンドでアプリを起動してください:")
print(f"  DATABASE_URL={DATABASE_URL} python3 app.py")
