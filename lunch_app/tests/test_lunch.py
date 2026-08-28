import os
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from extensions import db  # noqa: E402
from models import DailyMenu, GuestOrder, MealType, Order, User, set_setting  # noqa: E402
from timeutil import period_of, period_range, today_jst, week_monday  # noqa: E402
from tests.conftest import login  # noqa: E402


def type_ids(app):
    with app.app_context():
        return {t.code: t.id for t in MealType.query.all()}


def this_month(app):
    with app.app_context():
        return period_of(today_jst(), 21).strftime("%Y-%m")


# ── 締め期間の計算（21日〜翌月20日） ──

def test_period_runs_from_21st_to_20th():
    assert period_range(date(2026, 9, 1), 21) == (date(2026, 8, 21), date(2026, 9, 21))
    assert period_of(date(2026, 8, 20), 21) == date(2026, 8, 1)   # 8月分の最終日
    assert period_of(date(2026, 8, 21), 21) == date(2026, 9, 1)   # 9月分の初日
    assert period_of(date(2026, 9, 20), 21) == date(2026, 9, 1)


def test_period_can_follow_the_calendar_month():
    assert period_range(date(2026, 9, 1), 1) == (date(2026, 9, 1), date(2026, 10, 1))


# ── 認証 ──

def test_top_requires_login(client):
    res = client.get("/")
    assert res.status_code == 302
    assert "/login" in res.headers["Location"]


def test_login_and_logout(client):
    res = login(client, "taro", "taro12345")
    assert "昼食を注文" in res.get_data(as_text=True)
    assert client.get("/logout").status_code == 302


def test_login_rejects_bad_password(client):
    res = login(client, "taro", "wrong-password")
    assert "正しくありません" in res.get_data(as_text=True)


def test_staff_cannot_open_admin_pages(staff_client):
    res = staff_client.get("/admin/users", follow_redirects=True)
    assert "管理者のみ" in res.get_data(as_text=True)


# ── 社員の注文（1日1食・まとめて入力） ──

def test_staff_picks_one_meal_type_per_day(app, staff_client):
    ids = type_ids(app)
    today = today_jst().isoformat()

    staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today}": str(ids["udon"])},
        follow_redirects=True,
    )
    with app.app_context():
        order = Order.query.one()
        assert order.meal_type_id == ids["udon"]

    # 選び直すと上書きされる（1 日 1 食）
    staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today}": str(ids["teishoku"])},
        follow_redirects=True,
    )
    with app.app_context():
        assert Order.query.count() == 1
        assert Order.query.one().meal_type_id == ids["teishoku"]

    # 「なし」で取り消し
    staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today}": ""},
        follow_redirects=True,
    )
    with app.app_context():
        assert Order.query.count() == 0


def test_order_rejected_after_cutoff(app, staff_client):
    ids = type_ids(app)
    with app.app_context():
        set_setting("cutoff_time", "00:00")  # 本日分は締切済み
        db.session.commit()

    res = staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today_jst().isoformat()}": str(ids["udon"])},
        follow_redirects=True,
    )
    assert "締切済み" in res.get_data(as_text=True)
    with app.app_context():
        assert Order.query.count() == 0


def test_cutoff_can_be_set_to_the_previous_day(app, staff_client):
    ids = type_ids(app)
    with app.app_context():
        set_setting("cutoff_days_before", "1")  # 前日締切 → 本日分はもう変更できない
        db.session.commit()

    res = staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today_jst().isoformat()}": str(ids["udon"])},
        follow_redirects=True,
    )
    assert "締切済み" in res.get_data(as_text=True)
    with app.app_context():
        assert Order.query.count() == 0


def test_order_rejects_meal_type_not_served_that_day(app, staff_client):
    ids = type_ids(app)  # ラーメンは本日の献立に無い
    staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today_jst().isoformat()}": str(ids["ramen"])},
        follow_redirects=True,
    )
    with app.app_context():
        assert Order.query.count() == 0


# ── 献立編成（週まとめ） ──

def test_admin_saves_a_week_of_menus(app, admin_client):
    monday = week_monday(today_jst() + timedelta(days=7))
    key = monday.isoformat()
    admin_client.post(
        "/admin/menus/save",
        data={
            "week": key,
            f"teishoku_{key}": "鶏肉の黒酢炒め",
            f"noodle_{key}": "冷やし肉味噌豆乳ラーメン",
            f"form_{key}": "ramen",
        },
        follow_redirects=True,
    )
    with app.app_context():
        rows = DailyMenu.query.filter_by(serve_date=monday).all()
        assert {r.meal_type.code for r in rows} == {"teishoku", "ramen"}
        assert {r.dish_name for r in rows} == {"鶏肉の黒酢炒め", "冷やし肉味噌豆乳ラーメン"}


def test_udon_soba_creates_both_meal_types(app, admin_client):
    monday = week_monday(today_jst() + timedelta(days=14))
    key = monday.isoformat()
    admin_client.post(
        "/admin/menus/save",
        data={"week": key, f"noodle_{key}": "冷やし豚しゃぶおろし", f"form_{key}": "udon_soba"},
        follow_redirects=True,
    )
    with app.app_context():
        rows = DailyMenu.query.filter_by(serve_date=monday).all()
        assert {r.meal_type.code for r in rows} == {"udon", "soba"}


def test_menu_with_orders_cannot_be_removed(app, staff_client, admin_client):
    ids = type_ids(app)
    today = today_jst()
    staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today.isoformat()}": str(ids["udon"])},
        follow_redirects=True,
    )

    monday = week_monday(today)
    key = today.isoformat()
    res = admin_client.post(
        "/admin/menus/save",
        data={"week": monday.isoformat(), f"teishoku_{key}": "ハンバーグ", f"form_{key}": "none"},
        follow_redirects=True,
    )
    assert "外せませんでした" in res.get_data(as_text=True)
    with app.app_context():
        assert DailyMenu.query.filter_by(serve_date=today, meal_type_id=ids["udon"]).first() is not None


# ── 集計 ──

def test_daily_report_counts_staff_and_guests(app, staff_client, admin_client):
    ids = type_ids(app)
    today = today_jst().isoformat()
    staff_client.post(
        "/orders/save",
        data={"month": this_month(app), f"choice_{today}": str(ids["udon"])},
        follow_redirects=True,
    )
    admin_client.post(
        "/admin/orders/guests",
        data={"date": today, f"guest_{ids['teishoku']}": "3"},
        follow_redirects=True,
    )

    res = admin_client.get(f"/reports/daily?date={today}")
    body = res.get_data(as_text=True)
    assert "山田 太郎" in body
    assert "鈴木 花子" in body  # 注文なしの一覧に出る
    with app.app_context():
        assert GuestOrder.query.one().count == 3


def test_admin_proxy_order_works_after_cutoff(app, admin_client):
    ids = type_ids(app)
    with app.app_context():
        set_setting("cutoff_time", "00:00")
        db.session.commit()
        staff_id = User.query.filter_by(username="taro").one().id

    admin_client.post(
        "/admin/orders/proxy",
        data={"date": today_jst().isoformat(), "user_id": str(staff_id), "meal_type_id": str(ids["soba"])},
        follow_redirects=True,
    )
    with app.app_context():
        assert Order.query.filter_by(user_id=staff_id).one().meal_type_id == ids["soba"]


def test_monthly_report_uses_the_21st_to_20th_period(app, admin_client):
    ids = type_ids(app)
    with app.app_context():
        staff_id = User.query.filter_by(username="taro").one().id
        # 8/20 は 8 月分、8/21 と 9/20 は 9 月分に入る
        for day in (date(2026, 8, 20), date(2026, 8, 21), date(2026, 9, 20)):
            db.session.add(Order(user_id=staff_id, serve_date=day, meal_type_id=ids["teishoku"]))
        db.session.commit()

    res = admin_client.get("/reports/monthly?month=2026-09")
    body = res.get_data(as_text=True)
    assert "2026年9月分（8/21〜9/20）" in body
    assert "山田 太郎" in body

    res_aug = admin_client.get("/reports/monthly?month=2026-08")
    assert "2026年8月分（7/21〜8/20）" in res_aug.get_data(as_text=True)


def test_exports_return_xlsx(app, admin_client):
    for url in (
        f"/reports/daily/export.xlsx?date={today_jst().isoformat()}",
        "/reports/monthly/export.xlsx?month=2026-09",
    ):
        res = admin_client.get(url)
        assert res.status_code == 200
        assert res.data[:2] == b"PK"  # xlsx は zip


# ── 管理機能 ──

def test_admin_can_add_employee(app, admin_client):
    admin_client.post(
        "/admin/users/add",
        data={
            "username": "jiro", "name": "佐藤 次郎", "employee_no": "A003",
            "department": "品質保証部", "sort_order": "3", "password": "jiro123456",
        },
        follow_redirects=True,
    )
    with app.app_context():
        user = User.query.filter_by(username="jiro").one()
        assert user.must_change_password is True
        assert user.check_password("jiro123456")


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


def test_settings_update_cutoff_and_period(app, admin_client):
    admin_client.post(
        "/admin/settings",
        data={"cutoff_time": "09:30", "cutoff_days_before": "1",
              "period_start_day": "21", "shop_name": "ハーベスト"},
        follow_redirects=True,
    )
    res = admin_client.get("/admin/settings")
    body = res.get_data(as_text=True)
    assert "09:30" in body
    assert "ハーベスト" in body
    assert "前日 09:30 まで" in body


def test_all_pages_render(app, admin_client):
    urls = [
        "/", "/change-password", "/admin/menus", "/admin/users",
        "/admin/settings", "/reports/daily", "/reports/monthly",
    ]
    for url in urls:
        assert admin_client.get(url).status_code == 200, url


def test_unknown_page_returns_404(staff_client):
    assert staff_client.get("/no-such-page").status_code == 404
