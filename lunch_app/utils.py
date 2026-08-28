from datetime import date, datetime, timedelta
from functools import wraps

from flask import flash, g, redirect, session, url_for

from extensions import db
from models import User, get_setting, get_setting_int
from timeutil import fmt_date_short, now_jst, parse_time


def current_user() -> User | None:
    """ログイン中のユーザー。1 リクエスト内では g にキャッシュする。"""
    if "user" not in g:
        user_id = session.get("user_id")
        g.user = db.session.get(User, user_id) if user_id else None
    return g.user


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        user = current_user()
        if user is None or not user.is_active:
            session.clear()
            return redirect(url_for("auth.login"))
        return f(*args, **kwargs)

    return decorated


def admin_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not current_user().is_admin:
            flash("この画面は管理者のみ利用できます。", "error")
            return redirect(url_for("orders.index"))
        return f(*args, **kwargs)

    return decorated


# ── 締切 ────────────────────────────────────────────────────

def cutoff_at(serve_date: date) -> datetime:
    """その日の注文締切。提供日の n 日前の指定時刻（n=0 なら当日）。"""
    days_before = max(get_setting_int("cutoff_days_before", 1), 0)
    return datetime.combine(serve_date - timedelta(days=days_before), parse_time(get_setting("cutoff_time")))


def is_open(serve_date: date, now: datetime | None = None) -> bool:
    """社員がその日の注文を追加・変更できるか。"""
    return (now or now_jst()) < cutoff_at(serve_date)


def cutoff_rule_text() -> str:
    """「前日 10:00 まで」のような、締切ルールの説明文。"""
    days_before = max(get_setting_int("cutoff_days_before", 1), 0)
    when = {0: "当日", 1: "前日", 2: "2日前"}.get(days_before, f"{days_before}日前")
    return f"{when} {get_setting('cutoff_time')} まで"


def order_status(serve_date: date) -> dict:
    """画面表示用の締切ステータス。"""
    deadline = cutoff_at(serve_date)
    return {
        "open": now_jst() < deadline,
        "deadline": deadline,
        "deadline_text": f"{fmt_date_short(deadline.date())} {deadline:%H:%M}",
        "rule_text": cutoff_rule_text(),
    }
