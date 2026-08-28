from flask import Blueprint, flash, redirect, render_template, request, url_for
from sqlalchemy import select

from extensions import db
from models import DailyMenu, MealType, Order, get_setting_int
from timeutil import add_months, parse_month, period_label, period_of, period_range, today_jst
from utils import current_user, is_open, login_required, order_status

bp = Blueprint("orders", __name__)


def menus_by_date(start, end_exclusive) -> dict:
    """期間内の献立を {日付: [DailyMenu, ...]} にまとめる。"""
    rows = (
        db.session.execute(
            select(DailyMenu)
            .join(MealType)
            .where(DailyMenu.serve_date >= start, DailyMenu.serve_date < end_exclusive)
            .order_by(DailyMenu.serve_date, MealType.sort_order)
        )
        .scalars()
        .all()
    )
    grouped: dict = {}
    for menu in rows:
        grouped.setdefault(menu.serve_date, []).append(menu)
    return grouped


def current_period(month_arg: str):
    """表示する集計期間を (月初日, 開始日, 終了日の翌日, 締め開始日) で返す。"""
    start_day = get_setting_int("period_start_day", 21)
    month = parse_month(month_arg, default=period_of(today_jst(), start_day))
    start, end = period_range(month, start_day)
    return month, start, end, start_day


@bp.route("/")
@login_required
def index():
    month, start, end, start_day = current_period(request.args.get("month", ""))
    grouped = menus_by_date(start, end)

    my_orders = {
        o.serve_date: o
        for o in Order.query.filter(
            Order.user_id == current_user().id,
            Order.serve_date >= start,
            Order.serve_date < end,
        ).all()
    }

    days = []
    for serve_date in sorted(grouped):
        days.append(
            {
                "date": serve_date,
                "menus": grouped[serve_date],
                "order": my_orders.get(serve_date),
                "open": is_open(serve_date),
            }
        )

    return render_template(
        "index.html",
        month=month,
        days=days,
        period_text=period_label(month, start_day),
        prev_month=add_months(month, -1),
        next_month=add_months(month, 1),
        my_count=len(my_orders),
        today=today_jst(),
        rule_text=order_status(today_jst())["rule_text"],
    )


@bp.route("/orders/save", methods=["POST"])
@login_required
def save():
    user = current_user()
    month, start, end, _ = current_period(request.form.get("month", ""))
    back = redirect(url_for("orders.index", month=month.strftime("%Y-%m")))

    grouped = menus_by_date(start, end)
    existing = {
        o.serve_date: o
        for o in Order.query.filter(
            Order.user_id == user.id, Order.serve_date >= start, Order.serve_date < end
        ).all()
    }

    changed = 0
    skipped = []
    for serve_date, menus in grouped.items():
        field = request.form.get(f"choice_{serve_date.isoformat()}")
        if field is None:
            continue  # 画面に出ていない日は触らない

        order = existing.get(serve_date)
        chosen = field.strip()
        current_id = str(order.meal_type_id) if order else ""
        if chosen == current_id:
            continue

        if not is_open(serve_date):
            skipped.append(serve_date)
            continue

        if chosen == "":
            if order:
                db.session.delete(order)
                changed += 1
            continue

        offered = {str(m.meal_type_id) for m in menus}
        if chosen not in offered:
            skipped.append(serve_date)
            continue

        if order:
            order.meal_type_id = int(chosen)
        else:
            db.session.add(
                Order(user_id=user.id, serve_date=serve_date, meal_type_id=int(chosen))
            )
        changed += 1

    db.session.commit()

    if skipped:
        dates = "、".join(f"{d.month}/{d.day}" for d in sorted(skipped))
        flash(f"締切済みのため変更できなかった日があります：{dates}", "error")
    flash(f"注文を保存しました。（{changed} 日分を変更）" if changed else "変更はありませんでした。",
          "success" if changed else "info")
    return back
