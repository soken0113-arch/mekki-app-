import os
import sys
import secrets as _secrets_mod
from datetime import datetime, timedelta
from flask import Flask, redirect, url_for, session, flash, request, render_template
from extensions import csrf, limiter
from blueprints.auth import bp as auth_bp
from blueprints.orders import bp as orders_bp
from blueprints.masters import bp as masters_bp
from blueprints.shipments import bp as shipments_bp

app = Flask(__name__)

_sk = os.environ.get("SECRET_KEY")
if not _sk:
    _sk = _secrets_mod.token_hex(32)
    print("[警告] SECRET_KEY が未設定です。ランダムキーを使用します（再起動するとセッションが無効になります）", file=sys.stderr)
app.secret_key = _sk
app.permanent_session_lifetime = timedelta(minutes=30)

SESSION_TIMEOUT = 30

csrf.init_app(app)
limiter.init_app(app)

app.register_blueprint(auth_bp)
app.register_blueprint(orders_bp)
app.register_blueprint(masters_bp)
app.register_blueprint(shipments_bp)


@app.before_request
def check_session_timeout():
    if session.get("user_id"):
        last_active = session.get("last_active")
        if last_active:
            elapsed = datetime.now() - datetime.fromisoformat(last_active)
            if elapsed > timedelta(minutes=SESSION_TIMEOUT):
                session.clear()
                flash(f"{SESSION_TIMEOUT}分間操作がなかったため自動ログアウトしました。", "info")
                return redirect(url_for("auth.login"))
        session["last_active"] = datetime.now().isoformat()
        if session.get("must_change_password") and request.endpoint not in ("auth.change_password", "auth.logout", "static"):
            flash("初回ログインのため、パスワードを変更してください。", "info")
            return redirect(url_for("auth.change_password"))


@app.errorhandler(429)
def ratelimit_handler(e):
    return render_template("login.html", error="ログイン試行回数が多すぎます。1分後に再試行してください。"), 429


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)
