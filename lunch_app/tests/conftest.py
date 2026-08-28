import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import create_app  # noqa: E402
from config import Config  # noqa: E402
from extensions import db  # noqa: E402
from models import DailyMenu, MenuItem, User, set_setting  # noqa: E402
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
        staff = User(username="taro", name="山田太郎", employee_no="A001", department="製造部")
        staff.set_password("taro12345")
        staff.must_change_password = False
        db.session.add(staff)

        item = MenuItem(name="から揚げ弁当", price=550, sort_order=1)
        limited = MenuItem(name="日替わり弁当", price=600, sort_order=2)
        db.session.add_all([item, limited])
        db.session.flush()

        db.session.add(DailyMenu(serve_date=today_jst(), menu_item_id=item.id, price=550))
        db.session.add(
            DailyMenu(serve_date=today_jst(), menu_item_id=limited.id, price=600, limit_count=2)
        )
        set_setting("cutoff_time", "23:59")  # 既定では「本日分は受付中」
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
