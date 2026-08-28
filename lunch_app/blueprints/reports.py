from io import BytesIO

from flask import Blueprint, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from sqlalchemy import and_, func, select

from extensions import db
from models import DailyMenu, MenuItem, Order, User, get_setting
from timeutil import add_months, fmt_date_full, fmt_month, month_range, parse_date, parse_month
from utils import admin_required, order_status

bp = Blueprint("reports", __name__, url_prefix="/reports")


# ── 日次（発注用） ──────────────────────────────────────────

def _daily_summary(serve_date):
    """その日のメニュー別 発注個数と金額。注文 0 のメニューも行として残す。"""
    rows = db.session.execute(
        select(
            DailyMenu.id,
            MenuItem.name,
            DailyMenu.price,
            func.coalesce(func.sum(Order.quantity), 0).label("qty"),
        )
        .select_from(DailyMenu)
        .join(MenuItem, MenuItem.id == DailyMenu.menu_item_id)
        .outerjoin(Order, Order.daily_menu_id == DailyMenu.id)
        .where(DailyMenu.serve_date == serve_date)
        .group_by(DailyMenu.id, MenuItem.name, DailyMenu.price, MenuItem.sort_order)
        .order_by(MenuItem.sort_order, MenuItem.name)
    ).all()
    return [
        {
            "daily_menu_id": r.id,
            "name": r.name,
            "price": r.price,
            "qty": r.qty,
            "amount": r.qty * r.price,
        }
        for r in rows
    ]


def _daily_by_user(serve_date):
    """社員別の内訳（誰が何を頼んだか）。"""
    orders = (
        Order.query.join(User, User.id == Order.user_id)
        .filter(Order.serve_date == serve_date)
        .order_by(User.employee_no, User.name, Order.id)
        .all()
    )
    grouped: dict = {}
    for order in orders:
        entry = grouped.setdefault(order.user_id, {"user": order.user, "items": [], "amount": 0})
        entry["items"].append(order)
        entry["amount"] += order.subtotal
    return list(grouped.values())


@bp.route("/daily")
@admin_required
def daily():
    serve_date = parse_date(request.args.get("date", ""))
    summary = _daily_summary(serve_date)
    return render_template(
        "report_daily.html",
        serve_date=serve_date,
        summary=summary,
        by_user=_daily_by_user(serve_date),
        total_qty=sum(r["qty"] for r in summary),
        total_amount=sum(r["amount"] for r in summary),
        users=User.query.filter_by(is_active=True).order_by(User.employee_no, User.name).all(),
        shop_name=get_setting("shop_name"),
        status=order_status(serve_date),
    )


@bp.route("/daily/export.xlsx")
@admin_required
def daily_export():
    serve_date = parse_date(request.args.get("date", ""))
    summary = _daily_summary(serve_date)
    shop_name = get_setting("shop_name")

    wb = Workbook()
    ws = wb.active
    ws.title = "発注書"
    ws.append([f"昼食発注書　{fmt_date_full(serve_date)}"])
    ws["A1"].font = Font(bold=True, size=14)
    if shop_name:
        ws.append([f"発注先: {shop_name}"])
    ws.append([])
    ws.append(["メニュー", "単価", "数量", "金額"])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    for row in summary:
        ws.append([row["name"], row["price"], row["qty"], row["amount"]])
    ws.append(["合計", "", sum(r["qty"] for r in summary), sum(r["amount"] for r in summary)])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    ws2 = wb.create_sheet("社員別内訳")
    ws2.append(["社員番号", "部署", "氏名", "メニュー", "数量", "金額", "備考"])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    for entry in _daily_by_user(serve_date):
        user = entry["user"]
        for order in entry["items"]:
            ws2.append(
                [
                    user.employee_no,
                    user.department,
                    user.name,
                    order.daily_menu.menu_item.name,
                    order.quantity,
                    order.subtotal,
                    order.note,
                ]
            )

    _autosize(ws)
    _autosize(ws2)
    return _xlsx_response(wb, f"lunch_order_{serve_date:%Y%m%d}.xlsx")


# ── 月次（個人別集計） ──────────────────────────────────────

def _monthly_rows(month):
    start, end = month_range(month)
    rows = db.session.execute(
        select(
            User.id,
            User.employee_no,
            User.department,
            User.name,
            User.is_active,
            func.coalesce(func.sum(Order.quantity), 0).label("qty"),
            func.coalesce(func.sum(Order.quantity * DailyMenu.price), 0).label("amount"),
        )
        .select_from(User)
        .outerjoin(
            Order,
            and_(Order.user_id == User.id, Order.serve_date >= start, Order.serve_date < end),
        )
        .outerjoin(DailyMenu, DailyMenu.id == Order.daily_menu_id)
        .group_by(User.id, User.employee_no, User.department, User.name, User.is_active)
        .order_by(User.employee_no, User.name)
    ).all()
    # 退職者などの無効ユーザーは、その月に注文があるときだけ表示する
    return [r for r in rows if r.is_active or r.qty]


@bp.route("/monthly")
@admin_required
def monthly():
    month = parse_month(request.args.get("month", ""))
    rows = _monthly_rows(month)
    start, end = month_range(month)

    by_date = db.session.execute(
        select(
            Order.serve_date,
            func.coalesce(func.sum(Order.quantity), 0).label("qty"),
            func.coalesce(func.sum(Order.quantity * DailyMenu.price), 0).label("amount"),
        )
        .join(DailyMenu, DailyMenu.id == Order.daily_menu_id)
        .where(Order.serve_date >= start, Order.serve_date < end)
        .group_by(Order.serve_date)
        .order_by(Order.serve_date)
    ).all()

    return render_template(
        "report_monthly.html",
        month=month,
        rows=rows,
        by_date=by_date,
        total_qty=sum(r.qty for r in rows),
        total_amount=sum(r.amount for r in rows),
        prev_month=add_months(month, -1),
        next_month=add_months(month, 1),
    )


@bp.route("/monthly/export.xlsx")
@admin_required
def monthly_export():
    month = parse_month(request.args.get("month", ""))
    start, end = month_range(month)
    rows = _monthly_rows(month)

    wb = Workbook()
    ws = wb.active
    ws.title = "社員別集計"
    ws.append([f"昼食代 個人別集計　{fmt_month(month)}"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(["社員番号", "部署", "氏名", "食数", "金額"])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    for r in rows:
        ws.append([r.employee_no, r.department, r.name, r.qty, r.amount])
    ws.append(["", "", "合計", sum(r.qty for r in rows), sum(r.amount for r in rows)])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    ws2 = wb.create_sheet("明細")
    ws2.append(["日付", "社員番号", "氏名", "メニュー", "数量", "金額"])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    detail = (
        Order.query.join(User, User.id == Order.user_id)
        .filter(Order.serve_date >= start, Order.serve_date < end)
        .order_by(Order.serve_date, User.employee_no, User.name)
        .all()
    )
    for order in detail:
        ws2.append(
            [
                order.serve_date.strftime("%Y-%m-%d"),
                order.user.employee_no,
                order.user.name,
                order.daily_menu.menu_item.name,
                order.quantity,
                order.subtotal,
            ]
        )

    _autosize(ws)
    _autosize(ws2)
    return _xlsx_response(wb, f"lunch_monthly_{month:%Y%m}.xlsx")


# ── 共通 ────────────────────────────────────────────────────

def _autosize(ws) -> None:
    for column in ws.columns:
        width = max((len(str(c.value)) for c in column if c.value is not None), default=0)
        ws.column_dimensions[column[0].column_letter].width = min(max(width + 4, 10), 40)
    for cell in ws[1]:
        cell.alignment = Alignment(vertical="center")


def _xlsx_response(wb, filename):
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
