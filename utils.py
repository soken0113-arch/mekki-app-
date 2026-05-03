from datetime import datetime, timedelta
from functools import wraps
import requests
from flask import session, redirect, url_for

MEKKI_TYPES = [
    "ニッケルメッキ",
    "クロムメッキ",
    "金メッキ",
    "銀メッキ",
    "銅メッキ",
    "無電解ニッケル",
    "硬質クロム",
    "スズメッキ",
    "三価クロメート",
    "黒ニッケル",
    "黒クロメート",
    "六価クロメート",
    "ユニクロ",
    "白アルマイト",
    "黒アルマイト",
    "三価化成処理",
    "その他",
]

MEKKI_LINES = ["銅", "２号機", "無Ⅰ", "無Ⅱ", "特殊", "アルマイト", "化成処理", "亜鉛", "Ni自動バレル", "外注"]


def build_thickness(from_val, to_val):
    f = (from_val or "").strip()
    t = (to_val or "").strip()
    if f and t:
        return f"{f}〜{t}μm"
    elif f:
        return f"{f}μm"
    return ""


def parse_thickness(value):
    if not value:
        return "", ""
    v = value.replace("μm", "").strip()
    if "〜" in v:
        parts = v.split("〜", 1)
        return parts[0].strip(), parts[1].strip()
    return v, ""


_holidays_cache: tuple = (set(), None)


def get_jp_holidays():
    global _holidays_cache
    cached_holidays, cached_at = _holidays_cache
    if cached_at and (datetime.now() - cached_at) < timedelta(hours=24):
        return cached_holidays
    try:
        res = requests.get("https://holidays-jp.github.io/api/v1/date.json", timeout=3)
        holidays = set(res.json().keys())
        _holidays_cache = (holidays, datetime.now())
        return holidays
    except Exception:
        return cached_holidays


def get_prev_business_day(target_date, holidays):
    d = target_date - timedelta(days=1)
    while d.weekday() >= 5 or d.strftime('%Y-%m-%d') in holidays:
        d -= timedelta(days=1)
    return d


def _build_order_list(rows, shipped_ids, holidays, today):
    orders = []
    for row in rows:
        if row["id"] in shipped_ids:
            continue
        order = dict(row)
        due = None
        try:
            due = datetime.strptime(order['due_date'], '%Y-%m-%d').date()
        except Exception:
            pass
        if due is None:
            order['alert'] = None
        elif due < today:
            order['alert'] = 'overdue'
        elif due == today:
            order['alert'] = 'today'
        elif get_prev_business_day(due, holidays) == today:
            order['alert'] = 'tomorrow'
        else:
            order['alert'] = None
        orders.append(order)
    return orders


def _register_shipment(conn, order_id, order):
    existing = conn.execute("SELECT id FROM shipments WHERE order_id=?", (order_id,)).fetchone()
    if existing:
        return False
    now = datetime.now()
    conn.execute(
        "INSERT INTO shipments (order_id, shipped_at, shipped_by, note, created_at) VALUES (?, ?, ?, ?, ?)",
        (order_id, now.strftime("%Y-%m-%d %H:%M:%S"), order["assigned_to"], "", now.strftime("%Y-%m-%d %H:%M:%S"))
    )
    return True


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("auth.login"))
        return f(*args, **kwargs)
    return decorated
