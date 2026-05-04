"""add columns to orders, products, users

Revision ID: 001
Revises:
Create Date: 2026-05-04

"""
from alembic import op

revision = '001'
down_revision = '000'
branch_labels = None
depends_on = None


def upgrade():
    # orders テーブル: 追加カラム
    op.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS mekki_thickness  TEXT NOT NULL DEFAULT ''")
    op.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS thickness_data   TEXT NOT NULL DEFAULT '不要'")
    op.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS unit_price       TEXT NOT NULL DEFAULT ''")
    op.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS mekki_line       TEXT NOT NULL DEFAULT ''")
    op.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS process_note     TEXT NOT NULL DEFAULT ''")
    op.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS shipping_method  TEXT NOT NULL DEFAULT ''")
    op.execute("ALTER TABLE orders ADD COLUMN IF NOT EXISTS subcontractor_id INTEGER")

    # orders テーブル: 廃止カラム削除
    for col in ["sub_part_no", "sub_part_name", "sub_qty", "sub_amount",
                "sub_total", "sub_due_date", "sub_request", "sub_note"]:
        op.execute(f"""
            DO $$ BEGIN
                ALTER TABLE orders DROP COLUMN IF EXISTS {col};
            END $$;
        """)

    # products テーブル
    op.execute("ALTER TABLE products ADD COLUMN IF NOT EXISTS unit_price TEXT NOT NULL DEFAULT ''")

    # users テーブル
    op.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS must_change_password BOOLEAN NOT NULL DEFAULT FALSE")


def downgrade():
    op.execute("ALTER TABLE orders DROP COLUMN IF EXISTS mekki_thickness")
    op.execute("ALTER TABLE orders DROP COLUMN IF EXISTS thickness_data")
    op.execute("ALTER TABLE orders DROP COLUMN IF EXISTS unit_price")
    op.execute("ALTER TABLE orders DROP COLUMN IF EXISTS mekki_line")
    op.execute("ALTER TABLE orders DROP COLUMN IF EXISTS process_note")
    op.execute("ALTER TABLE orders DROP COLUMN IF EXISTS shipping_method")
    op.execute("ALTER TABLE orders DROP COLUMN IF EXISTS subcontractor_id")
    op.execute("ALTER TABLE products DROP COLUMN IF EXISTS unit_price")
    op.execute("ALTER TABLE users DROP COLUMN IF EXISTS must_change_password")
