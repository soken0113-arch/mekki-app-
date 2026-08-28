from datetime import date, datetime
from functools import wraps

from flask import flash, g, redirect, session, url_for

from extensions import db
from models import User, get_setting
from timeutil import now_jst, parse_time


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


def cutoff_at(serve_date: date) -> datetime:
    """その日の注文締切時刻（例: 提供日の 10:00）。"""
    return datetime.combine(serve_date, parse_time(get_setting("cutoff_time")))


def is_open(serve_date: date, now: datetime | None = None) -> bool:
    """社員がその日の注文を追加・変更できるか。"""
    return (now or now_jst()) < cutoff_at(serve_date)


def order_status(serve_date: date) -> dict:
    """画面表示用の締切ステータス。"""
    deadline = cutoff_at(serve_date)
    open_ = now_jst() < deadline
    return {
        "open": open_,
        "deadline": deadline,
        "deadline_text": deadline.strftime("%m/%d %H:%M"),
        "cutoff_text": deadline.strftime("%H:%M"),
    }


def yen(value) -> str:
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return "0"
