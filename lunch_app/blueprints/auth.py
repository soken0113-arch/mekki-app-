from collections import defaultdict
from datetime import timedelta

from flask import Blueprint, flash, redirect, render_template, request, session, url_for

from extensions import db
from models import User
from timeutil import now_jst
from utils import current_user, login_required

bp = Blueprint("auth", __name__)

# ── 簡易ログイン試行制限（プロセス内メモリ） ──
MAX_ATTEMPTS = 10
LOCKOUT = timedelta(minutes=5)
_attempts: dict[str, list] = defaultdict(list)


def _too_many_attempts(key: str) -> bool:
    now = now_jst()
    recent = [t for t in _attempts[key] if now - t < LOCKOUT]
    _attempts[key] = recent
    return len(recent) >= MAX_ATTEMPTS


def _record_failure(key: str) -> None:
    _attempts[key].append(now_jst())


@bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect(url_for("orders.index"))

    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        key = f"{request.remote_addr}|{username}"

        if _too_many_attempts(key):
            error = "ログイン試行が多すぎます。5分ほど待ってから再度お試しください。"
        else:
            user = User.query.filter_by(username=username).first()
            if user and user.is_active and user.check_password(password):
                session.clear()
                session.permanent = True
                session["user_id"] = user.id
                session["last_active"] = now_jst().isoformat()
                session["must_change_password"] = user.must_change_password
                _attempts.pop(key, None)
                return redirect(request.args.get("next") or url_for("orders.index"))
            _record_failure(key)
            error = "ユーザー名またはパスワードが正しくありません。"

    return render_template("login.html", error=error)


@bp.route("/logout")
def logout():
    session.clear()
    flash("ログアウトしました。", "info")
    return redirect(url_for("auth.login"))


@bp.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    user = current_user()
    error = None
    success = None

    if request.method == "POST":
        current = request.form.get("current_password", "")
        new_pw = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")

        if not user.check_password(current):
            error = "現在のパスワードが正しくありません。"
        elif len(new_pw) < 8:
            error = "新しいパスワードは8文字以上で入力してください。"
        elif new_pw == current:
            error = "現在のパスワードとは別のパスワードを設定してください。"
        elif new_pw != confirm:
            error = "新しいパスワードと確認用パスワードが一致しません。"
        else:
            user.set_password(new_pw)
            user.must_change_password = False
            db.session.commit()
            session["must_change_password"] = False
            success = "パスワードを変更しました。"

    return render_template("change_password.html", error=error, success=success)
