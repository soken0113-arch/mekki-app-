from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from db import get_db
from utils import login_required
from extensions import csrf, limiter

bp = Blueprint("auth", __name__)


@bp.route("/login", methods=["GET", "POST"])
@csrf.exempt
@limiter.limit("5 per minute")
def login():
    if session.get("user_id"):
        return redirect(url_for("orders.index"))
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        with get_db() as conn:
            user = conn.execute(
                "SELECT * FROM users WHERE username=?", (username,)
            ).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session.permanent = True
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["last_active"] = datetime.now().isoformat()
            session["must_change_password"] = bool(user.get("must_change_password"))
            return redirect(url_for("orders.index"))
        error = "ユーザー名またはパスワードが正しくありません。"
    return render_template("login.html", error=error)


@bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("auth.login"))


@bp.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    error = None
    success = None
    if request.method == "POST":
        current = request.form.get("current_password", "")
        new_pw  = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")
        with get_db() as conn:
            user = conn.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
        if not check_password_hash(user["password_hash"], current):
            error = "現在のパスワードが正しくありません。"
        elif len(new_pw) < 6:
            error = "新しいパスワードは6文字以上で入力してください。"
        elif new_pw != confirm:
            error = "新しいパスワードと確認用パスワードが一致しません。"
        else:
            with get_db() as conn:
                conn.execute(
                    "UPDATE users SET password_hash=?, must_change_password=FALSE WHERE id=?",
                    (generate_password_hash(new_pw, method="pbkdf2:sha256"), session["user_id"])
                )
            session["must_change_password"] = False
            success = "パスワードを変更しました。"
    return render_template("change_password.html", error=error, success=success)
