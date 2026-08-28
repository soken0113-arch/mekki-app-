from sqlalchemy import UniqueConstraint

from werkzeug.security import check_password_hash, generate_password_hash

from extensions import db
from timeutil import now_jst


class User(db.Model):
    """社員（利用者）。管理者フラグを持つ。並び順は注文用紙の名簿順に合わせられる。"""

    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), nullable=False, unique=True)
    name = db.Column(db.String(64), nullable=False)
    employee_no = db.Column(db.String(32), nullable=False, default="")
    department = db.Column(db.String(64), nullable=False, default="")
    password_hash = db.Column(db.String(255), nullable=False)
    is_admin = db.Column(db.Boolean, nullable=False, default=False)
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    must_change_password = db.Column(db.Boolean, nullable=False, default=True)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)

    orders = db.relationship("Order", back_populates="user", cascade="all, delete-orphan")

    def set_password(self, raw: str) -> None:
        self.password_hash = generate_password_hash(raw, method="pbkdf2:sha256")

    def check_password(self, raw: str) -> bool:
        return check_password_hash(self.password_hash, raw)

    def __repr__(self) -> str:
        return f"<User {self.username}>"


class MealType(db.Model):
    """注文の区分。注文用紙の「定食・丼／うどん／そば／ラーメン／パスタ」に対応する。"""

    __tablename__ = "meal_types"

    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(20), nullable=False, unique=True)
    name = db.Column(db.String(40), nullable=False, unique=True)
    # 麺の区分は献立編成でまとめて扱うため、定食系と区別する
    is_noodle = db.Column(db.Boolean, nullable=False, default=False)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    is_active = db.Column(db.Boolean, nullable=False, default=True)

    def __repr__(self) -> str:
        return f"<MealType {self.name}>"


# 初回起動時に投入する区分（code は画面やコードから参照するため固定）
DEFAULT_MEAL_TYPES = [
    ("teishoku", "定食・丼", False, 1),
    ("udon",     "うどん",   True,  2),
    ("soba",     "そば",     True,  3),
    ("ramen",    "ラーメン", True,  4),
    ("pasta",    "パスタ",   True,  5),
]

# 麺の提供形態（献立編成でこの単位で選ぶ）
NOODLE_FORMS = {
    "none":  ("提供なし",       []),
    "udon_soba": ("うどん・そば", ["udon", "soba"]),
    "ramen": ("ラーメン",       ["ramen"]),
    "pasta": ("パスタ",         ["pasta"]),
}


class DailyMenu(db.Model):
    """その日に提供する区分と献立名。ここに行がある区分だけ注文できる。"""

    __tablename__ = "daily_menus"
    __table_args__ = (UniqueConstraint("serve_date", "meal_type_id", name="uq_daily_menu"),)

    id = db.Column(db.Integer, primary_key=True)
    serve_date = db.Column(db.Date, nullable=False, index=True)
    meal_type_id = db.Column(db.Integer, db.ForeignKey("meal_types.id"), nullable=False)
    dish_name = db.Column(db.String(120), nullable=False, default="")
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)

    meal_type = db.relationship("MealType")


class Order(db.Model):
    """社員 1 人につき 1 日 1 食。区分を 1 つ選ぶ（注文用紙の 1 マスに相当）。"""

    __tablename__ = "orders"
    __table_args__ = (UniqueConstraint("user_id", "serve_date", name="uq_order"),)

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    serve_date = db.Column(db.Date, nullable=False, index=True)
    meal_type_id = db.Column(db.Integer, db.ForeignKey("meal_types.id"), nullable=False)
    note = db.Column(db.String(200), nullable=False, default="")
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)
    updated_at = db.Column(db.DateTime, nullable=False, default=now_jst, onupdate=now_jst)

    user = db.relationship("User", back_populates="orders")
    meal_type = db.relationship("MealType")


class GuestOrder(db.Model):
    """来客用のまとめ注文。担当者が日付・区分ごとに人数を入力する。"""

    __tablename__ = "guest_orders"
    __table_args__ = (UniqueConstraint("serve_date", "meal_type_id", name="uq_guest_order"),)

    id = db.Column(db.Integer, primary_key=True)
    serve_date = db.Column(db.Date, nullable=False, index=True)
    meal_type_id = db.Column(db.Integer, db.ForeignKey("meal_types.id"), nullable=False)
    count = db.Column(db.Integer, nullable=False, default=0)
    note = db.Column(db.String(200), nullable=False, default="")
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)
    updated_at = db.Column(db.DateTime, nullable=False, default=now_jst, onupdate=now_jst)

    meal_type = db.relationship("MealType")


class Setting(db.Model):
    """締切時刻などのアプリ設定を key/value で保持する。"""

    __tablename__ = "settings"

    key = db.Column(db.String(64), primary_key=True)
    value = db.Column(db.String(255), nullable=False, default="")
    updated_at = db.Column(db.DateTime, nullable=False, default=now_jst, onupdate=now_jst)


DEFAULT_SETTINGS = {
    "cutoff_time": "10:00",       # 締切時刻
    "cutoff_days_before": "1",    # 提供日の何日前で締め切るか（0 = 当日）
    "period_start_day": "21",     # 月次集計の開始日（21 なら 21 日〜翌月 20 日）
    "shop_name": "",              # 発注書に出す仕出し屋名
    "access_url": "",             # 社員に案内する接続先 URL（空なら自動判定）
}


def get_setting(key: str, default: str = "") -> str:
    row = db.session.get(Setting, key)
    if row is None:
        return DEFAULT_SETTINGS.get(key, default)
    return row.value


def get_setting_int(key: str, default: int = 0) -> int:
    try:
        return int(get_setting(key))
    except ValueError:
        return default


def set_setting(key: str, value: str) -> None:
    row = db.session.get(Setting, key)
    if row is None:
        db.session.add(Setting(key=key, value=value))
    else:
        row.value = value
        row.updated_at = now_jst()
