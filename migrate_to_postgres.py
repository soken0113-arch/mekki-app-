"""
SQLite (orders.db) のデータを PostgreSQL に移行するスクリプト
使い方:
  DATABASE_URL=postgresql://... python migrate_to_postgres.py
"""
import sqlite3
import os
import psycopg2
import psycopg2.extras

SQLITE_PATH = os.environ.get("SQLITE_PATH", "orders.db")
DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    raise SystemExit("ERROR: DATABASE_URL 環境変数を設定してください")

sqlite_conn = sqlite3.connect(SQLITE_PATH)
sqlite_conn.row_factory = sqlite3.Row
pg_conn = psycopg2.connect(DATABASE_URL)
pg_cur = pg_conn.cursor()

tables = ["orders", "customers", "products", "users"]

for table in tables:
    rows = sqlite_conn.execute(f"SELECT * FROM {table}").fetchall()
    if not rows:
        print(f"{table}: データなし、スキップ")
        continue

    cols = rows[0].keys()
    placeholders = ", ".join(["%s"] * len(cols))
    col_names = ", ".join(cols)

    for row in rows:
        pg_cur.execute(
            f"INSERT INTO {table} ({col_names}) VALUES ({placeholders}) ON CONFLICT DO NOTHING",
            tuple(row)
        )

    # SERIAL のシーケンスを最大IDに合わせる
    pg_cur.execute(f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), COALESCE(MAX(id), 1)) FROM {table}")
    print(f"{table}: {len(rows)} 件を移行しました")

pg_conn.commit()
pg_conn.close()
sqlite_conn.close()
print("移行完了！")
