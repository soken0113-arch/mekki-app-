import os
import secrets
import sys
from datetime import timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _database_uri() -> str:
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        return "sqlite:///" + os.path.join(BASE_DIR, "lunch.db")
    # Heroku / Render 形式の postgres:// を SQLAlchemy 2.x 形式へ
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    return url


def _secret_key() -> str:
    key = os.environ.get("SECRET_KEY")
    if key:
        return key
    print(
        "[警告] SECRET_KEY が未設定です。ランダムキーを使用します"
        "（再起動するとログイン状態が失われます）",
        file=sys.stderr,
    )
    return secrets.token_hex(32)


class Config:
    SECRET_KEY = _secret_key()
    SQLALCHEMY_DATABASE_URI = _database_uri()
    SQLALCHEMY_ENGINE_OPTIONS = {"pool_pre_ping": True}
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    SESSION_TIMEOUT_MINUTES = int(os.environ.get("SESSION_TIMEOUT_MINUTES", "120"))
    PERMANENT_SESSION_LIFETIME = timedelta(minutes=SESSION_TIMEOUT_MINUTES)
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"

    # 初回起動時に作成する管理者
    ADMIN_USERNAME = os.environ.get("LUNCH_ADMIN_USERNAME", "admin")
    ADMIN_PASSWORD = os.environ.get("LUNCH_ADMIN_PASSWORD", "admin1234")
