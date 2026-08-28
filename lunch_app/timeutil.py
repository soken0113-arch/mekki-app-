"""日本時間（JST）で日付・時刻を扱うためのヘルパー。

サーバーの TZ が UTC でも「今日」が日本の今日とずれないよう、
アプリ内の now / today はすべてここを通す。
"""
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

JST = ZoneInfo("Asia/Tokyo")

WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]


def now_jst() -> datetime:
    """タイムゾーンを外した（naive な）日本時間の現在時刻。DB 保存用。"""
    return datetime.now(JST).replace(tzinfo=None)


def today_jst() -> date:
    return now_jst().date()


def parse_date(value: str, default: date | None = None) -> date:
    """'YYYY-MM-DD' を date に。空・不正なら default（既定は今日）。"""
    try:
        return datetime.strptime((value or "").strip(), "%Y-%m-%d").date()
    except ValueError:
        return default if default is not None else today_jst()


def parse_month(value: str, default: date | None = None) -> date:
    """'YYYY-MM' をその月の 1 日の date に。空・不正なら default（既定は今月）。"""
    try:
        return datetime.strptime((value or "").strip(), "%Y-%m").date().replace(day=1)
    except ValueError:
        return default if default is not None else today_jst().replace(day=1)


def parse_time(value: str, default: time | None = time(10, 0)) -> time | None:
    """'HH:MM' を time に。空・不正なら default。"""
    try:
        return datetime.strptime((value or "").strip(), "%H:%M").time()
    except ValueError:
        return default


def month_range(first_day: date) -> tuple[date, date]:
    """月初日から (月初, 翌月初) を返す。期間検索に使う。"""
    if first_day.month == 12:
        nxt = first_day.replace(year=first_day.year + 1, month=1, day=1)
    else:
        nxt = first_day.replace(month=first_day.month + 1, day=1)
    return first_day, nxt


def add_months(first_day: date, delta: int) -> date:
    total = first_day.year * 12 + (first_day.month - 1) + delta
    return date(total // 12, total % 12 + 1, 1)


def fmt_date(d: date) -> str:
    return f"{d.month}月{d.day}日（{WEEKDAY_JA[d.weekday()]}）"


def fmt_date_full(d: date) -> str:
    return f"{d.year}年{d.month}月{d.day}日（{WEEKDAY_JA[d.weekday()]}）"


def fmt_month(d: date) -> str:
    return f"{d.year}年{d.month}月"


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def week_dates(around: date, days: int = 7) -> list[date]:
    return [around + timedelta(days=i) for i in range(days)]
