from flask import Blueprint, flash, redirect, render_template, request, url_for
from sqlalchemy import select

from extensions import db
from models import DailyMenu, MenuItem, Order
from timeutil import add_months, month_range, parse_date, parse_month, today_jst
from utils import current_user, is_open, login_required, order_status

bp = Blueprint("orders", __name__)

MAX_QTY_PER_ITEM = 20


def _menus_for(serve_date):
    return (
        db.session.execute(
            select(DailyMenu)
            .join(MenuItem)
            .where(DailyMenu.serve_date == serve_date)
            .order_by(MenuItem.sort_order, MenuItem.name)
        )
        .scalars()
        .all()
    )


def _upcoming_dates(limit: int = 7):
    return (
        db.session.execute(
            select(DailyMenu.serve_date)
            .where(DailyMenu.serve_date >= today_jst())
            .group_by(DailyMenu.serve_date)
            .order_by(DailyMenu.serve_date)
            .limit(limit)
        )
        .scalars()
        .all()
    )


@bp.route("/")
@login_required
def index():
    serve_date = parse_date(request.args.get("date", ""))
    menus = _menus_for(serve_date)
    my_orders = {
        o.daily_menu_id: o
        for o in Order.query.filter_by(user_id=current_user().id, serve_date=serve_date).all()
    }
    my_total = sum(o.quantity * o.daily_menu.price for o in my_orders.values())

    return render_template(
        "index.html",
        serve_date=serve_date,
        menus=menus,
        my_orders=my_orders,
        my_total=my_total,
        status=order_status(serve_date),
        upcoming=_upcoming_dates(),
        max_qty=MAX_QTY_PER_ITEM,
    )


@bp.route("/orders/save", methods=["POST"])
@login_required
def save():
    user = current_user()
    serve_date = parse_date(request.form.get("date", ""))
    redirect_to = redirect(url_for("orders.index", date=serve_date.isoformat()))

    if not is_open(serve_date):
        flash("この日の注文は締め切られています。変更が必要な場合は担当者にご連絡ください。", "error")
        return redirect_to

    menus = _menus_for(serve_date)
    existing = {
        o.daily_menu_id: o
        for o in Order.query.filter_by(user_id=user.id, serve_date=serve_date).all()
    }

    errors = []
    for menu in menus:
        raw = request.form.get(f"qty_{menu.id}", "0").strip()
        note = request.form.get(f"note_{menu.id}", "").strip()[:200]
        try:
            qty = int(raw or 0)
        except ValueError:
            errors.append(f"「{menu.menu_item.name}」の数量が数値ではありません。")
            continue
        if qty < 0 or qty > MAX_QTY_PER_ITEM:
            errors.append(f"「{menu.menu_item.name}」の数量は 0〜{MAX_QTY_PER_ITEM} で入力してください。")
            continue

        order = existing.get(menu.id)
        if menu.limit_count is not None:
            others = menu.ordered_count - (order.quantity if order else 0)
            if others + qty > menu.limit_count:
                errors.append(
                    f"「{menu.menu_item.name}」は残り {max(menu.limit_count - others, 0)} 個です。"
                )
                continue

        if qty == 0:
            if order:
                db.session.delete(order)
        elif order:
            order.quantity = qty
            order.note = note
        else:
            db.session.add(
                Order(
                    user_id=user.id,
                    daily_menu_id=menu.id,
                    serve_date=serve_date,
                    quantity=qty,
                    note=note,
                )
            )

    if errors:
        db.session.rollback()
        for message in errors:
            flash(message, "error")
    else:
        db.session.commit()
        flash("注文を保存しました。", "success")
    return redirect_to


@bp.route("/history")
@login_required
def history():
    month = parse_month(request.args.get("month", ""))
    start, end = month_range(month)
    orders = (
        Order.query.filter(
            Order.user_id == current_user().id,
            Order.serve_date >= start,
            Order.serve_date < end,
        )
        .order_by(Order.serve_date)
        .all()
    )

    by_date: dict = {}
    for order in orders:
        by_date.setdefault(order.serve_date, []).append(order)

    total_amount = sum(o.subtotal for o in orders)
    total_count = sum(o.quantity for o in orders)

    return render_template(
        "history.html",
        month=month,
        by_date=by_date,
        total_amount=total_amount,
        total_count=total_count,
        prev_month=add_months(month, -1),
        next_month=add_months(month, 1),
    )
