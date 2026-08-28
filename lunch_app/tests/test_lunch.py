import os
import sys
from datetime import timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from extensions import db  # noqa: E402
from models import DailyMenu, MenuItem, Order, User, set_setting  # noqa: E402
from timeutil import today_jst  # noqa: E402
from tests.conftest import login  # noqa: E402


def menu_ids(app):
    with app.app_context():
        rows = (
            DailyMenu.query.join(MenuItem)
            .filter(DailyMenu.serve_date == today_jst())
            .order_by(MenuItem.sort_order)
            .all()
        )
        return [dm.id for dm in rows]


# ── 認証 ──

def test_top_requires_login(client):
    res = client.get("/")
    assert res.status_code == 302
    assert "/login" in res.headers["Location"]


def test_login_and_logout(client):
    res = login(client, "taro", "taro12345")
    assert res.status_code == 200
    assert "昼食を注文" in res.get_data(as_text=True)
    assert client.get("/logout").status_code == 302


def test_login_rejects_bad_password(client):
    res = login(client, "taro", "wrong-password")
    assert "正しくありません" in res.get_data(as_text=True)


def test_staff_cannot_open_admin_pages(staff_client):
    res = staff_client.get("/admin/users", follow_redirects=True)
    assert "管理者のみ" in res.get_data(as_text=True)


# ── 注文 ──

def test_staff_can_place_and_update_order(app, staff_client):
    first, _ = menu_ids(app)
    staff_client.post(
        "/orders/save",
        data={"date": today_jst().isoformat(), f"qty_{first}": "2"},
        follow_redirects=True,
    )
    with app.app_context():
        order = Order.query.filter_by(daily_menu_id=first).one()
        assert order.quantity == 2
        assert order.subtotal == 1100

    # 数量を 0 にすると取り消される
    staff_client.post(
        "/orders/save",
        data={"date": today_jst().isoformat(), f"qty_{first}": "0"},
        follow_redirects=True,
    )
    with app.app_context():
        assert Order.query.count() == 0


def test_order_rejected_after_cutoff(app, staff_client):
    first, _ = menu_ids(app)
    with app.app_context():
        set_setting("cutoff_time", "00:00")  # 本日分は締切済み
        db.session.commit()

    res = staff_client.post(
        "/orders/save",
        data={"date": today_jst().isoformat(), f"qty_{first}": "1"},
        follow_redirects=True,
    )
    assert "締め切られています" in res.get_data(as_text=True)
    with app.app_context():
        assert Order.query.count() == 0


def test_order_respects_quantity_limit(app, staff_client):
    _, limited = menu_ids(app)
    res = staff_client.post(
        "/orders/save",
        data={"date": today_jst().isoformat(), f"qty_{limited}": "3"},  # 上限 2
        follow_redirects=True,
    )
    assert "残り" in res.get_data(as_text=True)
    with app.app_context():
        assert Order.query.count() == 0


def test_order_rejects_out_of_range_quantity(app, staff_client):
    first, _ = menu_ids(app)
    res = staff_client.post(
        "/orders/save",
        data={"date": today_jst().isoformat(), f"qty_{first}": "99"},
        follow_redirects=True,
    )
    assert "0〜20" in res.get_data(as_text=True)
    with app.app_context():
        assert Order.query.count() == 0


def test_history_shows_monthly_total(app, staff_client):
    first, _ = menu_ids(app)
    staff_client.post(
        "/orders/save",
        data={"date": today_jst().isoformat(), f"qty_{first}": "2"},
        follow_redirects=True,
    )
    res = staff_client.get(f"/history?month={today_jst():%Y-%m}")
    body = res.get_data(as_text=True)
    assert "1,100" in body
    assert "2 食" in body


# ── 集計 ──

def test_daily_report_aggregates_orders(app, staff_client, admin_client):
    first, _ = menu_ids(app)
    staff_client.post(
        "/orders/save",
        data={"date": today_jst().isoformat(), f"qty_{first}": "3"},
        follow_redirects=True,
    )
    res = admin_client.get(f"/reports/daily?date={today_jst().isoformat()}")
    body = res.get_data(as_text=True)
    assert "1,650" in body       # 550 円 × 3
    assert "山田太郎" in body


def test_daily_and_monthly_exports_return_xlsx(app, admin_client):
    for url in (
        f"/reports/daily/export.xlsx?date={today_jst().isoformat()}",
        f"/reports/monthly/export.xlsx?month={today_jst():%Y-%m}",
    ):
        res = admin_client.get(url)
        assert res.status_code == 200
        assert res.data[:2] == b"PK"  # xlsx は zip


def test_monthly_report_lists_each_employee(app, admin_client):
    res = admin_client.get(f"/reports/monthly?month={today_jst():%Y-%m}")
    assert "山田太郎" in res.get_data(as_text=True)


# ── 管理機能 ──

def test_admin_can_manage_menu_and_daily_menu(app, admin_client):
    admin_client.post(
        "/admin/menu-items/add",
        data={"name": "そぼろ丼", "price": "500", "sort_order": "3"},
        follow_redirects=True,
    )
    with app.app_context():
        item = MenuItem.query.filter_by(name="そぼろ丼").one()
        item_id = item.id

    tomorrow = today_jst() + timedelta(days=1)
    admin_client.post(
        "/admin/daily-menus/save",
        data={"date": tomorrow.isoformat(), f"use_{item_id}": "on", f"price_{item_id}": "480"},
        follow_redirects=True,
    )
    with app.app_context():
        daily = DailyMenu.query.filter_by(menu_item_id=item_id, serve_date=tomorrow).one()
        assert daily.price == 480


def test_admin_proxy_order_works_after_cutoff(app, admin_client):
    first, _ = menu_ids(app)
    with app.app_context():
        set_setting("cutoff_time", "00:00")
        db.session.commit()
        staff_id = User.query.filter_by(username="taro").one().id

    admin_client.post(
        "/admin/orders/proxy",
        data={
            "date": today_jst().isoformat(),
            "user_id": str(staff_id),
            "daily_menu_id": str(first),
            "quantity": "1",
        },
        follow_redirects=True,
    )
    with app.app_context():
        assert Order.query.filter_by(user_id=staff_id).one().quantity == 1


def test_admin_can_add_employee(app, admin_client):
    admin_client.post(
        "/admin/users/add",
        data={
            "username": "hanako",
            "name": "鈴木花子",
            "employee_no": "A002",
            "department": "総務部",
            "password": "hanako12345",
        },
        follow_redirects=True,
    )
    with app.app_context():
        user = User.query.filter_by(username="hanako").one()
        assert user.must_change_password is True
        assert user.check_password("hanako12345")


def test_last_admin_cannot_be_demoted(app, admin_client):
    with app.app_context():
        admin_id = User.query.filter_by(username="admin").one().id
    res = admin_client.post(
        f"/admin/users/{admin_id}/edit",
        data={"name": "管理者", "is_active": "on"},  # is_admin のチェックを外した状態
        follow_redirects=True,
    )
    assert "管理者が 0 人になる" in res.get_data(as_text=True)
    with app.app_context():
        assert User.query.filter_by(id=admin_id).one().is_admin is True


def test_settings_update_changes_cutoff(app, admin_client):
    admin_client.post(
        "/admin/settings",
        data={"cutoff_time": "09:30", "shop_name": "○○弁当店"},
        follow_redirects=True,
    )
    res = admin_client.get("/admin/settings")
    assert "09:30" in res.get_data(as_text=True)


def test_all_admin_pages_render(app, admin_client):
    """管理者向けの全画面が 200 で表示されること（テンプレートの取りこぼし検知）。"""
    urls = [
        "/",
        "/history",
        "/change-password",
        "/admin/menu-items",
        "/admin/daily-menus",
        "/admin/users",
        "/admin/settings",
        "/reports/daily",
        "/reports/monthly",
    ]
    for url in urls:
        assert admin_client.get(url).status_code == 200, url


def test_unknown_page_returns_404(staff_client):
    assert staff_client.get("/no-such-page").status_code == 404
