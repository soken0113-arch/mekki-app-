"""日本時間（JST）で日付を扱うためのヘルパー。

サーバーの TZ が UTC でも「今日」が日本の今日とずれないよう、
アプリ内の now / today はすべてここを通す。
締め日（21 日〜翌月 20 日など）の期間計算もここにまとめる。
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


# ── 変換 ────────────────────────────────────────────────────

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


def add_months(first_day: date, delta: int) -> date:
    """月初日を delta ヶ月ずらす。"""
    total = first_day.year * 12 + (first_day.month - 1) + delta
    return date(total // 12, total % 12 + 1, 1)


# ── 締め期間（例: 21 日〜翌月 20 日 = 「翌月分」） ───────────

MAX_START_DAY = 28  # どの月にも存在する日に限定する


def clamp_start_day(day: int) -> int:
    return min(max(day, 1), MAX_START_DAY)


def period_range(month_first: date, start_day: int) -> tuple[date, date]:
    """「month_first の月の分」の集計期間を (開始日, 終了日の翌日) で返す。

    start_day=21 なら 9 月分 = 8/21 〜 9/20（返り値は 8/21 と 9/21）。
    start_day=1 なら 9 月分 = 9/1 〜 9/30。
    """
    start_day = clamp_start_day(start_day)
    if start_day == 1:
        return month_first, add_months(month_first, 1)
    return (
        add_months(month_first, -1).replace(day=start_day),
        month_first.replace(day=start_day),
    )


def period_of(target: date, start_day: int) -> date:
    """その日が「何月分」に入るかを月初日で返す。"""
    start_day = clamp_start_day(start_day)
    month_first = target.replace(day=1)
    if start_day == 1 or target.day < start_day:
        return month_first
    return add_months(month_first, 1)


def period_label(month_first: date, start_day: int) -> str:
    start, end = period_range(month_first, start_day)
    last = end - timedelta(days=1)
    if start_day == 1:
        return f"{month_first.year}年{month_first.month}月分"
    return (
        f"{month_first.year}年{month_first.month}月分"
        f"（{start.month}/{start.day}〜{last.month}/{last.day}）"
    )


# ── 週（献立編成用） ────────────────────────────────────────

def week_monday(target: date) -> date:
    """その日を含む週の月曜日。"""
    return target - timedelta(days=target.weekday())


def weekdays_of(monday: date, days: int = 5) -> list[date]:
    """月曜から days 日分（既定は月〜金）。"""
    return [monday + timedelta(days=i) for i in range(days)]


def date_range(start: date, end_exclusive: date) -> list[date]:
    out = []
    d = start
    while d < end_exclusive:
        out.append(d)
        d += timedelta(days=1)
    return out


# ── 表示 ────────────────────────────────────────────────────

def fmt_date(d: date) -> str:
    return f"{d.month}月{d.day}日（{WEEKDAY_JA[d.weekday()]}）"


def fmt_date_short(d: date) -> str:
    return f"{d.month}/{d.day}（{WEEKDAY_JA[d.weekday()]}）"


def fmt_date_full(d: date) -> str:
    return f"{d.year}年{d.month}月{d.day}日（{WEEKDAY_JA[d.weekday()]}）"


def fmt_month(d: date) -> str:
    return f"{d.year}年{d.month}月"


def weekday_ja(d: date) -> str:
    return WEEKDAY_JA[d.weekday()]


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5
