from datetime import timedelta

from flask import Blueprint, flash, redirect, render_template, request, url_for

from extensions import db
from models import (
    NOODLE_FORMS,
    DailyMenu,
    GuestOrder,
    MealType,
    Order,
    User,
    get_setting,
    get_setting_int,
    set_setting,
)
from timeutil import clamp_start_day, parse_date, parse_time, week_monday, weekdays_of
from utils import admin_required, cutoff_rule_text, current_user

bp = Blueprint("admin", __name__, url_prefix="/admin")


def _int_or_none(raw: str):
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        value = int(raw)
    except ValueError:
        return None
    return value if value >= 0 else None


def _meal_types():
    return MealType.query.filter_by(is_active=True).order_by(MealType.sort_order).all()


def _by_code(types):
    return {t.code: t for t in types}


# ── 献立編成（1 週間分をまとめて入力） ──────────────────────

@bp.route("/menus")
@admin_required
def menus():
    monday = week_monday(parse_date(request.args.get("week", "")))
    days = weekdays_of(monday)
    types = _meal_types()
    by_code = _by_code(types)

    rows = DailyMenu.query.filter(
        DailyMenu.serve_date >= days[0], DailyMenu.serve_date <= days[-1]
    ).all()

    # 画面用に「定食名 / 麺名 / 麺の提供形態」へ組み直す
    current: dict = {}
    for day in days:
        current[day] = {"teishoku": "", "noodle": "", "form": "none", "locked": set()}
    for row in rows:
        entry = current.get(row.serve_date)
        if entry is None:
            continue
        if row.meal_type.is_noodle:
            entry["noodle"] = row.dish_name
        else:
            entry["teishoku"] = row.dish_name
        if Order.query.filter_by(serve_date=row.serve_date, meal_type_id=row.meal_type_id).first():
            entry["locked"].add(row.meal_type.name)

    for day in days:
        codes = {r.meal_type.code for r in rows if r.serve_date == day and r.meal_type.is_noodle}
        for form, (_label, form_codes) in NOODLE_FORMS.items():
            if form_codes and codes == set(form_codes):
                current[day]["form"] = form
                break

    return render_template(
        "admin_menus.html",
        monday=monday,
        days=days,
        current=current,
        noodle_forms=NOODLE_FORMS,
        teishoku=by_code.get("teishoku"),
        prev_week=monday - timedelta(days=7),
        next_week=monday + timedelta(days=7),
        rule_text=cutoff_rule_text(),
    )


@bp.route("/menus/save", methods=["POST"])
@admin_required
def save_menus():
    monday = week_monday(parse_date(request.form.get("week", "")))
    days = weekdays_of(monday)
    types = _meal_types()
    by_code = _by_code(types)

    existing = {
        (row.serve_date, row.meal_type.code): row
        for row in DailyMenu.query.filter(
            DailyMenu.serve_date >= days[0], DailyMenu.serve_date <= days[-1]
        ).all()
    }

    blocked = []
    for day in days:
        key = day.isoformat()
        teishoku_name = request.form.get(f"teishoku_{key}", "").strip()[:120]
        noodle_name = request.form.get(f"noodle_{key}", "").strip()[:120]
        form = request.form.get(f"form_{key}", "none")
        if form not in NOODLE_FORMS:
            form = "none"

        wanted = {}
        if teishoku_name and "teishoku" in by_code:
            wanted["teishoku"] = teishoku_name
        if noodle_name:
            for code in NOODLE_FORMS[form][1]:
                if code in by_code:
                    wanted[code] = noodle_name

        for code, meal_type in by_code.items():
            row = existing.get((day, code))
            if code in wanted:
                if row:
                    row.dish_name = wanted[code]
                else:
                    db.session.add(
                        DailyMenu(serve_date=day, meal_type_id=meal_type.id, dish_name=wanted[code])
                    )
            elif row:
                if Order.query.filter_by(serve_date=day, meal_type_id=meal_type.id).first():
                    blocked.append(f"{day.month}/{day.day} {meal_type.name}")
                else:
                    db.session.delete(row)

    db.session.commit()
    if blocked:
        flash(
            "既に注文が入っているため外せませんでした：" + "、".join(blocked)
            + "（社員の注文を取り消してから外してください）",
            "error",
        )
    else:
        flash(f"{monday.month}/{monday.day} の週の献立を保存しました。", "success")
    return redirect(url_for("admin.menus", week=monday.isoformat()))


@bp.route("/menus/copy", methods=["POST"])
@admin_required
def copy_menus():
    monday = week_monday(parse_date(request.form.get("week", "")))
    source = week_monday(parse_date(request.form.get("copy_from", "")))
    if source == monday:
        flash("コピー元とコピー先が同じ週です。", "error")
        return redirect(url_for("admin.menus", week=monday.isoformat()))

    src_days = weekdays_of(source)
    dst_days = weekdays_of(monday)
    existing = {
        (row.serve_date, row.meal_type_id)
        for row in DailyMenu.query.filter(
            DailyMenu.serve_date >= dst_days[0], DailyMenu.serve_date <= dst_days[-1]
        ).all()
    }

    copied = 0
    for src_day, dst_day in zip(src_days, dst_days):
        for row in DailyMenu.query.filter_by(serve_date=src_day).all():
            if (dst_day, row.meal_type_id) in existing:
                continue
            db.session.add(
                DailyMenu(serve_date=dst_day, meal_type_id=row.meal_type_id, dish_name=row.dish_name)
            )
            copied += 1
    db.session.commit()
    flash(f"{source.month}/{source.day} の週から {copied} 件コピーしました。", "success" if copied else "info")
    return redirect(url_for("admin.menus", week=monday.isoformat()))


# ── 代理入力・来客用（発注集計画面から使う） ────────────────

@bp.route("/orders/proxy", methods=["POST"])
@admin_required
def proxy_order():
    serve_date = parse_date(request.form.get("date", ""))
    user_id = _int_or_none(request.form.get("user_id", ""))
    choice = (request.form.get("meal_type_id", "") or "").strip()
    back = redirect(url_for("reports.daily", date=serve_date.isoformat()))

    user = db.session.get(User, user_id) if user_id else None
    if user is None:
        flash("社員を選択してください。", "error")
        return back

    order = Order.query.filter_by(user_id=user.id, serve_date=serve_date).first()
    if choice == "":
        if order:
            db.session.delete(order)
            db.session.commit()
            flash(f"{user.name} さんの注文を取り消しました。", "success")
        else:
            flash(f"{user.name} さんの注文はもともとありません。", "info")
        return back

    meal_type_id = _int_or_none(choice)
    offered = DailyMenu.query.filter_by(serve_date=serve_date, meal_type_id=meal_type_id).first()
    if offered is None:
        flash("その日に提供されない区分は選べません。", "error")
        return back

    if order:
        order.meal_type_id = meal_type_id
    else:
        db.session.add(Order(user_id=user.id, serve_date=serve_date, meal_type_id=meal_type_id, note="代理入力"))
    db.session.commit()
    flash(f"{user.name} さんの注文を「{offered.meal_type.name}」にしました。", "success")
    return back


@bp.route("/orders/guests", methods=["POST"])
@admin_required
def save_guests():
    serve_date = parse_date(request.form.get("date", ""))
    existing = {g.meal_type_id: g for g in GuestOrder.query.filter_by(serve_date=serve_date).all()}

    total = 0
    for menu in DailyMenu.query.filter_by(serve_date=serve_date).all():
        count = _int_or_none(request.form.get(f"guest_{menu.meal_type_id}", "0")) or 0
        row = existing.get(menu.meal_type_id)
        if count == 0:
            if row:
                db.session.delete(row)
        elif row:
            row.count = count
        else:
            db.session.add(
                GuestOrder(serve_date=serve_date, meal_type_id=menu.meal_type_id, count=count)
            )
        total += count

    db.session.commit()
    flash(f"来客用を {total} 食で保存しました。", "success")
    return redirect(url_for("reports.daily", date=serve_date.isoformat()))


# ── 社員マスタ ──────────────────────────────────────────────

@bp.route("/users")
@admin_required
def users():
    rows = User.query.order_by(User.is_active.desc(), User.sort_order, User.name).all()
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
            sort_order=_int_or_none(request.form.get("sort_order", "")) or 0,
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

    is_admin = request.form.get("is_admin") == "on"
    is_active = request.form.get("is_active") == "on"
    # 最後の管理者から権限や利用を取り上げるとログインできなくなるため防ぐ
    if user.is_admin and (not is_admin or not is_active) and _admin_count() <= 1:
        flash("管理者が 0 人になるため、この操作はできません。", "error")
        return redirect(url_for("admin.users"))

    user.name = name
    user.employee_no = request.form.get("employee_no", "").strip()[:32]
    user.department = request.form.get("department", "").strip()[:64]
    user.sort_order = _int_or_none(request.form.get("sort_order", "")) or 0
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
        cutoff = parse_time(request.form.get("cutoff_time", ""), default=None)
        if cutoff is None:
            flash("締切時刻は HH:MM 形式で入力してください（例: 10:00）。", "error")
            return redirect(url_for("admin.settings"))

        set_setting("cutoff_time", cutoff.strftime("%H:%M"))
        set_setting("cutoff_days_before", str(_int_or_none(request.form.get("cutoff_days_before", "")) or 0))
        set_setting("period_start_day", str(clamp_start_day(_int_or_none(request.form.get("period_start_day", "")) or 1)))
        set_setting("shop_name", request.form.get("shop_name", "").strip()[:100])
        db.session.commit()
        flash("設定を保存しました。", "success")
        return redirect(url_for("admin.settings"))

    return render_template(
        "admin_settings.html",
        cutoff_time=get_setting("cutoff_time"),
        cutoff_days_before=get_setting_int("cutoff_days_before", 1),
        period_start_day=get_setting_int("period_start_day", 21),
        shop_name=get_setting("shop_name"),
        meal_types=MealType.query.order_by(MealType.sort_order).all(),
        rule_text=cutoff_rule_text(),
    )


@bp.route("/meal-types/<int:type_id>/edit", methods=["POST"])
@admin_required
def edit_meal_type(type_id):
    meal_type = db.get_or_404(MealType, type_id)
    name = request.form.get("name", "").strip()
    if not name:
        flash("区分名は必須です。", "error")
        return redirect(url_for("admin.settings"))

    is_active = request.form.get("is_active") == "on"
    if not is_active and Order.query.filter_by(meal_type_id=meal_type.id).first():
        flash(f"「{meal_type.name}」は注文実績があるため、使わない設定にはできません。", "error")
        return redirect(url_for("admin.settings"))

    meal_type.name = name[:40]
    meal_type.sort_order = _int_or_none(request.form.get("sort_order", "")) or 0
    meal_type.is_active = is_active
    db.session.commit()
    flash(f"区分「{name}」を更新しました。", "success")
    return redirect(url_for("admin.settings"))
