import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import create_app  # noqa: E402
from config import Config  # noqa: E402
from extensions import db  # noqa: E402
from models import DailyMenu, MealType, User, set_setting  # noqa: E402
from timeutil import today_jst  # noqa: E402


class TestConfig(Config):
    SECRET_KEY = "test-secret"
    SQLALCHEMY_DATABASE_URI = "sqlite://"
    WTF_CSRF_ENABLED = False
    TESTING = True
    ADMIN_USERNAME = "admin"
    ADMIN_PASSWORD = "admin1234"


@pytest.fixture
def app():
    application = create_app(TestConfig)
    with application.app_context():
        admin = User.query.filter_by(username="admin").first()
        admin.must_change_password = False

        staff = User(
            username="taro", name="山田 太郎", employee_no="A001",
            department="製造部", sort_order=1, must_change_password=False,
        )
        staff.set_password("taro12345")
        other = User(
            username="hanako", name="鈴木 花子", employee_no="A002",
            department="総務部", sort_order=2, must_change_password=False,
        )
        other.set_password("hanako12345")
        db.session.add_all([staff, other])

        # 本日は「定食・丼」「うどん」「そば」、翌日は「定食・丼」「ラーメン」を提供
        codes = {t.code: t for t in MealType.query.all()}
        for code, dish in [("teishoku", "ハンバーグ"), ("udon", "冷やしうどん"), ("soba", "冷やしそば")]:
            db.session.add(DailyMenu(serve_date=today_jst(), meal_type_id=codes[code].id, dish_name=dish))

        set_setting("cutoff_days_before", "0")
        set_setting("cutoff_time", "23:59")  # 既定では「本日分は受付中」
        set_setting("period_start_day", "21")
        db.session.commit()

    # app context を張ったままにするとリクエスト間で g が共有されてしまうため、
    # 準備が終わったらコンテキストを抜けてから yield する
    # （インメモリ SQLite は StaticPool なのでコンテキストを抜けてもデータは残る）
    yield application


@pytest.fixture
def client(app):
    return app.test_client()


def login(client, username, password):
    return client.post(
        "/login", data={"username": username, "password": password}, follow_redirects=True
    )


@pytest.fixture
def staff_client(app):
    """一般社員としてログイン済みのクライアント。"""
    c = app.test_client()
    login(c, "taro", "taro12345")
    return c


@pytest.fixture
def admin_client(app):
    """管理者としてログイン済みのクライアント。"""
    c = app.test_client()
    login(c, "admin", "admin1234")
    return c
