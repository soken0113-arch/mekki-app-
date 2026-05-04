"""Initial schema - create all tables

Revision ID: 000
Revises:
Create Date: 2026-05-04
"""
from alembic import op
from sqlalchemy import text
from datetime import datetime
from werkzeug.security import generate_password_hash

revision = '000'
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    op.execute("""
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
    op.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL
        )
    """)
    op.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            part_no TEXT NOT NULL DEFAULT '',
            unit_price TEXT NOT NULL DEFAULT '',
            note TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)
    op.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            must_change_password BOOLEAN NOT NULL DEFAULT FALSE,
            created_at TEXT NOT NULL
        )
    """)
    op.execute("""
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
    op.execute("""
        CREATE TABLE IF NOT EXISTS shipments (
            id SERIAL PRIMARY KEY,
            order_id INTEGER NOT NULL,
            shipped_at TEXT NOT NULL,
            shipped_by TEXT NOT NULL DEFAULT '',
            note TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL
        )
    """)
    op.execute("""
        CREATE TABLE IF NOT EXISTS subcontractors (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL
        )
    """)
    op.execute(
        text(
            "INSERT INTO users (username, password_hash, must_change_password, created_at) "
            "VALUES (:u, :ph, TRUE, :ca) ON CONFLICT DO NOTHING"
        ).bindparams(
            u="admin",
            ph=generate_password_hash("admin1234", method="pbkdf2:sha256"),
            ca=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )
    )


def downgrade():
    op.execute("DROP TABLE IF EXISTS shipments")
    op.execute("DROP TABLE IF EXISTS order_items")
    op.execute("DROP TABLE IF EXISTS subcontractors")
    op.execute("DROP TABLE IF EXISTS orders")
    op.execute("DROP TABLE IF EXISTS products")
    op.execute("DROP TABLE IF EXISTS customers")
    op.execute("DROP TABLE IF EXISTS users")
