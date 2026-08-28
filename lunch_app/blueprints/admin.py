from flask import Blueprint, flash, redirect, render_template, request, url_for
from sqlalchemy import select
from sqlalchemy.exc import IntegrityError

from extensions import db
from models import DailyMenu, MenuItem, Order, User, get_setting, set_setting
from timeutil import parse_date, parse_time, today_jst
from utils import admin_required, current_user

bp = Blueprint("admin", __name__, url_prefix="/admin")

MAX_QTY_PER_ITEM = 20


def _int_or_none(raw: str):
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        value = int(raw)
    except ValueError:
        return None
    return value if value >= 0 else None


# ── メニューマスタ ──────────────────────────────────────────

@bp.route("/menu-items")
@admin_required
def menu_items():
    items = MenuItem.query.order_by(MenuItem.sort_order, MenuItem.name).all()
    return render_template("admin_menu_items.html", items=items)


@bp.route("/menu-items/add", methods=["POST"])
@admin_required
def add_menu_item():
    name = request.form.get("name", "").strip()
    price = _int_or_none(request.form.get("price", ""))
    if not name:
        flash("メニュー名を入力してください。", "error")
    elif price is None:
        flash("価格は 0 以上の数値で入力してください。", "error")
    elif MenuItem.query.filter_by(name=name).first():
        flash(f"「{name}」は既に登録されています。", "error")
    else:
        db.session.add(
            MenuItem(
                name=name,
                price=price,
                description=request.form.get("description", "").strip()[:200],
                sort_order=_int_or_none(request.form.get("sort_order", "")) or 0,
            )
        )
        db.session.commit()
        flash(f"「{name}」を登録しました。", "success")
    return redirect(url_for("admin.menu_items"))


@bp.route("/menu-items/<int:item_id>/edit", methods=["POST"])
@admin_required
def edit_menu_item(item_id):
    item = db.get_or_404(MenuItem, item_id)
    name = request.form.get("name", "").strip()
    price = _int_or_none(request.form.get("price", ""))
    if not name or price is None:
        flash("メニュー名と価格（0 以上の数値）は必須です。", "error")
        return redirect(url_for("admin.menu_items"))

    item.name = name
    item.price = price
    item.description = request.form.get("description", "").strip()[:200]
    item.sort_order = _int_or_none(request.form.get("sort_order", "")) or 0
    item.is_active = request.form.get("is_active") == "on"
    try:
        db.session.commit()
        flash(f"「{name}」を更新しました。", "success")
    except IntegrityError:
        db.session.rollback()
        flash(f"「{name}」は既に登録されています。", "error")
    return redirect(url_for("admin.menu_items"))


@bp.route("/menu-items/<int:item_id>/delete", methods=["POST"])
@admin_required
def delete_menu_item(item_id):
    item = db.get_or_404(MenuItem, item_id)
    if DailyMenu.query.filter_by(menu_item_id=item.id).first():
        flash(
            f"「{item.name}」は過去の献立で使用されているため削除できません。"
            "提供を止める場合は「表示」のチェックを外してください。",
            "error",
        )
        return redirect(url_for("admin.menu_items"))
    db.session.delete(item)
    db.session.commit()
    flash(f"「{item.name}」を削除しました。", "success")
    return redirect(url_for("admin.menu_items"))


# ── 日次メニュー（献立編成） ────────────────────────────────

@bp.route("/daily-menus")
@admin_required
def daily_menus():
    serve_date = parse_date(request.args.get("date", ""))
    items = MenuItem.query.filter_by(is_active=True).order_by(MenuItem.sort_order, MenuItem.name).all()
    current = {dm.menu_item_id: dm for dm in DailyMenu.query.filter_by(serve_date=serve_date).all()}
    ordered = {item_id: dm.ordered_count for item_id, dm in current.items()}
    return render_template(
        "admin_daily_menus.html",
        serve_date=serve_date,
        items=items,
        current=current,
        ordered=ordered,
        cutoff=get_setting("cutoff_time"),
    )


@bp.route("/daily-menus/save", methods=["POST"])
@admin_required
def save_daily_menus():
    serve_date = parse_date(request.form.get("date", ""))
    items = MenuItem.query.filter_by(is_active=True).all()
    current = {dm.menu_item_id: dm for dm in DailyMenu.query.filter_by(serve_date=serve_date).all()}

    added = removed = 0
    blocked = []
    for item in items:
        checked = request.form.get(f"use_{item.id}") == "on"
        price = _int_or_none(request.form.get(f"price_{item.id}", ""))
        limit = _int_or_none(request.form.get(f"limit_{item.id}", ""))
        existing = current.get(item.id)

        if checked:
            if existing:
                existing.price = item.price if price is None else price
                existing.limit_count = limit
            else:
                db.session.add(
                    DailyMenu(
                        serve_date=serve_date,
                        menu_item_id=item.id,
                        price=item.price if price is None else price,
                        limit_count=limit,
                    )
                )
                added += 1
        elif existing:
            if existing.ordered_count > 0:
                blocked.append(item.name)
            else:
                db.session.delete(existing)
                removed += 1

    db.session.commit()
    if blocked:
        flash(
            "既に注文が入っているため外せませんでした: " + "、".join(blocked)
            + "（注文を取り消してから外してください）",
            "error",
        )
    if added or removed or not blocked:
        flash(f"{serve_date:%m/%d} の献立を保存しました。（追加 {added} 件 / 削除 {removed} 件）", "success")
    return redirect(url_for("admin.daily_menus", date=serve_date.isoformat()))


@bp.route("/daily-menus/copy", methods=["POST"])
@admin_required
def copy_daily_menus():
    serve_date = parse_date(request.form.get("date", ""))
    source = parse_date(request.form.get("copy_from", ""), default=None)
    if source == serve_date:
        flash("コピー元とコピー先が同じ日付です。", "error")
        return redirect(url_for("admin.daily_menus", date=serve_date.isoformat()))

    existing_ids = {dm.menu_item_id for dm in DailyMenu.query.filter_by(serve_date=serve_date).all()}
    copied = 0
    for src in DailyMenu.query.filter_by(serve_date=source).all():
        if src.menu_item_id in existing_ids:
            continue
        db.session.add(
            DailyMenu(
                serve_date=serve_date,
                menu_item_id=src.menu_item_id,
                price=src.price,
                limit_count=src.limit_count,
            )
        )
        copied += 1
    db.session.commit()
    flash(f"{source:%m/%d} から {copied} 件コピーしました。", "success" if copied else "info")
    return redirect(url_for("admin.daily_menus", date=serve_date.isoformat()))


# ── 代理注文（締切後・欠席者の分を管理者が入力） ────────────

@bp.route("/orders/proxy", methods=["POST"])
@admin_required
def proxy_order():
    serve_date = parse_date(request.form.get("date", ""))
    user_id = _int_or_none(request.form.get("user_id", ""))
    daily_menu_id = _int_or_none(request.form.get("daily_menu_id", ""))
    qty = _int_or_none(request.form.get("quantity", ""))
    back = redirect(url_for("reports.daily", date=serve_date.isoformat()))

    if not user_id or not daily_menu_id or qty is None or qty > MAX_QTY_PER_ITEM:
        flash("社員・メニュー・数量（0〜20）を正しく指定してください。", "error")
        return back

    menu = db.session.get(DailyMenu, daily_menu_id)
    user = db.session.get(User, user_id)
    if menu is None or user is None or menu.serve_date != serve_date:
        flash("指定された社員またはメニューが見つかりません。", "error")
        return back

    order = Order.query.filter_by(user_id=user_id, daily_menu_id=daily_menu_id).first()
    if qty == 0:
        if order:
            db.session.delete(order)
            flash(f"{user.name} の「{menu.menu_item.name}」を取り消しました。", "success")
    elif order:
        order.quantity = qty
        flash(f"{user.name} の「{menu.menu_item.name}」を {qty} 個に変更しました。", "success")
    else:
        db.session.add(
            Order(
                user_id=user_id,
                daily_menu_id=daily_menu_id,
                serve_date=serve_date,
                quantity=qty,
                note="代理入力",
            )
        )
        flash(f"{user.name} の「{menu.menu_item.name}」を {qty} 個で登録しました。", "success")
    db.session.commit()
    return back


# ── 社員マスタ ──────────────────────────────────────────────

@bp.route("/users")
@admin_required
def users():
    rows = User.query.order_by(User.is_active.desc(), User.employee_no, User.name).all()
    return render_template("admin_users.html", users=rows)


@bp.route("/users/add", methods=["POST"])
@admin_required
def add_user():
    username = request.form.get("username", "").strip()
    name = request.form.get("name", "").strip()
    password = request.form.get("password", "")

    if not username or not name:
        flash("ログインIDと氏名は必須です。", "error")
    elif len(password) < 8:
        flash("初期パスワードは8文字以上で入力してください。", "error")
    elif User.query.filter_by(username=username).first():
        flash(f"ログインID「{username}」は既に使われています。", "error")
    else:
        user = User(
            username=username,
            name=name,
            employee_no=request.form.get("employee_no", "").strip()[:32],
            department=request.form.get("department", "").strip()[:64],
            is_admin=request.form.get("is_admin") == "on",
            must_change_password=True,
        )
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        flash(f"{name} さんを登録しました。初回ログイン時にパスワード変更が必要です。", "success")
    return redirect(url_for("admin.users"))


@bp.route("/users/<int:user_id>/edit", methods=["POST"])
@admin_required
def edit_user(user_id):
    user = db.get_or_404(User, user_id)
    name = request.form.get("name", "").strip()
    if not name:
        flash("氏名は必須です。", "error")
        return redirect(url_for("admin.users"))

    user.name = name
    user.employee_no = request.form.get("employee_no", "").strip()[:32]
    user.department = request.form.get("department", "").strip()[:64]
    is_admin = request.form.get("is_admin") == "on"
    is_active = request.form.get("is_active") == "on"

    # 最後の管理者から権限や利用を取り上げるとログインできなくなるため防ぐ
    if user.is_admin and (not is_admin or not is_active) and _admin_count() <= 1:
        flash("管理者が 0 人になるため、この操作はできません。", "error")
        return redirect(url_for("admin.users"))

    user.is_admin = is_admin
    user.is_active = is_active
    db.session.commit()
    flash(f"{user.name} さんの情報を更新しました。", "success")
    return redirect(url_for("admin.users"))


@bp.route("/users/<int:user_id>/reset-password", methods=["POST"])
@admin_required
def reset_password(user_id):
    user = db.get_or_404(User, user_id)
    password = request.form.get("password", "")
    if len(password) < 8:
        flash("パスワードは8文字以上で入力してください。", "error")
    else:
        user.set_password(password)
        user.must_change_password = True
        db.session.commit()
        flash(f"{user.name} さんのパスワードを再設定しました。", "success")
    return redirect(url_for("admin.users"))


@bp.route("/users/<int:user_id>/delete", methods=["POST"])
@admin_required
def delete_user(user_id):
    user = db.get_or_404(User, user_id)
    if user.id == current_user().id:
        flash("自分自身は削除できません。", "error")
    elif user.is_admin and _admin_count() <= 1:
        flash("管理者が 0 人になるため削除できません。", "error")
    elif Order.query.filter_by(user_id=user.id).first():
        flash(
            f"{user.name} さんには注文履歴があるため削除できません。"
            "「利用可」のチェックを外して無効化してください。",
            "error",
        )
    else:
        db.session.delete(user)
        db.session.commit()
        flash(f"{user.name} さんを削除しました。", "success")
    return redirect(url_for("admin.users"))


def _admin_count() -> int:
    return User.query.filter_by(is_admin=True, is_active=True).count()


# ── 設定 ────────────────────────────────────────────────────

@bp.route("/settings", methods=["GET", "POST"])
@admin_required
def settings():
    if request.method == "POST":
        raw = request.form.get("cutoff_time", "").strip()
        parsed = parse_time(raw, default=None)
        if parsed is None:
            flash("締切時刻は HH:MM 形式で入力してください（例: 10:00）。", "error")
        else:
            set_setting("cutoff_time", parsed.strftime("%H:%M"))
            set_setting("shop_name", request.form.get("shop_name", "").strip()[:100])
            db.session.commit()
            flash("設定を保存しました。", "success")
        return redirect(url_for("admin.settings"))

    return render_template(
        "admin_settings.html",
        cutoff_time=get_setting("cutoff_time"),
        shop_name=get_setting("shop_name"),
    )
