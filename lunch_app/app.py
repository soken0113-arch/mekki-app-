import os
from datetime import datetime, timedelta

import click
from flask import Flask, flash, redirect, render_template, request, session, url_for

from config import Config
from extensions import csrf, db
from timeutil import fmt_date, fmt_date_full, fmt_date_short, fmt_month, now_jst, today_jst, weekday_ja
from utils import current_user


def create_app(config_object=Config) -> Flask:
    app = Flask(__name__)
    app.config.from_object(config_object)

    db.init_app(app)
    csrf.init_app(app)

    from blueprints.admin import bp as admin_bp
    from blueprints.auth import bp as auth_bp
    from blueprints.orders import bp as orders_bp
    from blueprints.reports import bp as reports_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(orders_bp)
    app.register_blueprint(admin_bp)
    app.register_blueprint(reports_bp)

    _register_hooks(app)
    _register_template_helpers(app)
    _register_cli(app)

    with app.app_context():
        db.create_all()
        _ensure_meal_types()
        _ensure_admin(app)

    return app


def _register_hooks(app: Flask) -> None:
    timeout = app.config["SESSION_TIMEOUT_MINUTES"]

    @app.before_request
    def touch_session():
        if not session.get("user_id"):
            return None
        last = session.get("last_active")
        if last:
            elapsed = now_jst() - datetime.fromisoformat(last)
            if elapsed > timedelta(minutes=timeout):
                session.clear()
                flash(f"{timeout}分間操作がなかったため自動ログアウトしました。", "info")
                return redirect(url_for("auth.login"))
        session["last_active"] = now_jst().isoformat()

        # 初期パスワードのままなら変更画面へ誘導する
        allowed = ("auth.change_password", "auth.logout", "static")
        if session.get("must_change_password") and request.endpoint not in allowed:
            flash("初回ログインのため、パスワードを変更してください。", "info")
            return redirect(url_for("auth.change_password"))
        return None


def _register_template_helpers(app: Flask) -> None:
    app.jinja_env.filters["date_ja"] = fmt_date
    app.jinja_env.filters["date_ja_short"] = fmt_date_short
    app.jinja_env.filters["date_ja_full"] = fmt_date_full
    app.jinja_env.filters["month_ja"] = fmt_month
    app.jinja_env.filters["weekday_ja"] = weekday_ja

    @app.context_processor
    def inject_globals():
        return {"current_user": current_user(), "today": today_jst()}

    @app.errorhandler(404)
    def not_found(_e):
        return render_template("error.html", code=404, message="ページが見つかりません。"), 404

    @app.errorhandler(500)
    def server_error(_e):  # pragma: no cover - 例外時のみ
        db.session.rollback()
        return render_template("error.html", code=500, message="サーバーエラーが発生しました。"), 500


def _ensure_meal_types() -> None:
    """注文区分（定食・丼／うどん／そば／ラーメン／パスタ）を初回に投入する。"""
    from models import DEFAULT_MEAL_TYPES, MealType

    for code, name, is_noodle, order in DEFAULT_MEAL_TYPES:
        if MealType.query.filter_by(code=code).first() is None:
            db.session.add(
                MealType(code=code, name=name, is_noodle=is_noodle, sort_order=order)
            )
    db.session.commit()


def _ensure_admin(app: Flask) -> None:
    """管理者が 1 人もいなければ初期管理者を作る（初回起動用）。"""
    from models import User

    if db.session.query(User.id).first() is not None:
        return
    admin = User(
        username=app.config["ADMIN_USERNAME"],
        name="管理者",
        is_admin=True,
        must_change_password=True,
    )
    admin.set_password(app.config["ADMIN_PASSWORD"])
    db.session.add(admin)
    db.session.commit()
    app.logger.warning(
        "初期管理者を作成しました: %s（初回ログイン後にパスワードを変更してください）",
        admin.username,
    )


def _register_cli(app: Flask) -> None:
    from models import User

    @app.cli.command("create-admin")
    @click.argument("username")
    @click.argument("password")
    @click.option("--name", default="管理者", help="表示名")
    def create_admin(username, password, name):
        """管理者アカウントを追加する。"""
        if User.query.filter_by(username=username).first():
            raise click.ClickException(f"ユーザー名 {username} は既に使われています。")
        user = User(username=username, name=name, is_admin=True, must_change_password=True)
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        click.echo(f"管理者 {username} を作成しました。")

    @app.cli.command("add-users")
    @click.argument("csv_path", type=click.Path(exists=True))
    def add_users(csv_path):
        """CSV（社員番号,氏名,部署,ログインID,初期パスワード）から社員をまとめて登録する。"""
        import csv

        added = skipped = 0
        with open(csv_path, encoding="utf-8-sig", newline="") as fp:
            for i, row in enumerate(csv.reader(fp)):
                if not row or (i == 0 and "氏名" in "".join(row)):
                    continue
                employee_no, name, department, username, password = (list(row) + [""] * 5)[:5]
                username = username.strip()
                name = name.strip()
                if not username or not name or len(password.strip()) < 8:
                    skipped += 1
                    continue
                if User.query.filter_by(username=username).first():
                    skipped += 1
                    continue
                user = User(
                    username=username,
                    name=name,
                    employee_no=employee_no.strip(),
                    department=department.strip(),
                    sort_order=i,
                    must_change_password=True,
                )
                user.set_password(password.strip())
                db.session.add(user)
                added += 1
        db.session.commit()
        click.echo(f"{added} 名を登録しました。（{skipped} 行はスキップ）")


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5002)), debug=True)
