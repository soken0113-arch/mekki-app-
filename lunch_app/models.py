from sqlalchemy import UniqueConstraint
from werkzeug.security import check_password_hash, generate_password_hash

from extensions import db
from timeutil import now_jst


class User(db.Model):
    """社員（利用者）。管理者フラグを持つ。"""

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
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)

    orders = db.relationship("Order", back_populates="user", cascade="all, delete-orphan")

    def set_password(self, raw: str) -> None:
        self.password_hash = generate_password_hash(raw, method="pbkdf2:sha256")

    def check_password(self, raw: str) -> bool:
        return check_password_hash(self.password_hash, raw)

    def __repr__(self) -> str:
        return f"<User {self.username}>"


class MenuItem(db.Model):
    """メニューマスタ。日々の献立はここから選んで組む。"""

    __tablename__ = "menu_items"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False, unique=True)
    price = db.Column(db.Integer, nullable=False, default=0)
    description = db.Column(db.String(200), nullable=False, default="")
    is_active = db.Column(db.Boolean, nullable=False, default=True)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)

    daily_menus = db.relationship("DailyMenu", back_populates="menu_item")

    def __repr__(self) -> str:
        return f"<MenuItem {self.name}>"


class DailyMenu(db.Model):
    """ある日に提供するメニュー。価格はその日の値段を保持する（後からマスタを変えても履歴が狂わない）。"""

    __tablename__ = "daily_menus"
    __table_args__ = (UniqueConstraint("serve_date", "menu_item_id", name="uq_daily_menu"),)

    id = db.Column(db.Integer, primary_key=True)
    serve_date = db.Column(db.Date, nullable=False, index=True)
    menu_item_id = db.Column(db.Integer, db.ForeignKey("menu_items.id"), nullable=False)
    price = db.Column(db.Integer, nullable=False, default=0)
    limit_count = db.Column(db.Integer, nullable=True)  # 数量上限（未設定なら無制限）
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)

    menu_item = db.relationship("MenuItem", back_populates="daily_menus")
    orders = db.relationship("Order", back_populates="daily_menu", cascade="all, delete-orphan")

    @property
    def ordered_count(self) -> int:
        return sum(o.quantity for o in self.orders)

    @property
    def remaining(self):
        if self.limit_count is None:
            return None
        return self.limit_count - self.ordered_count


class Order(db.Model):
    """社員 1 人 × その日の 1 メニュー = 1 行。数量で複数個に対応する。"""

    __tablename__ = "orders"
    __table_args__ = (UniqueConstraint("user_id", "daily_menu_id", name="uq_order"),)

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    daily_menu_id = db.Column(db.Integer, db.ForeignKey("daily_menus.id"), nullable=False, index=True)
    serve_date = db.Column(db.Date, nullable=False, index=True)
    quantity = db.Column(db.Integer, nullable=False, default=1)
    note = db.Column(db.String(200), nullable=False, default="")
    created_at = db.Column(db.DateTime, nullable=False, default=now_jst)
    updated_at = db.Column(db.DateTime, nullable=False, default=now_jst, onupdate=now_jst)

    user = db.relationship("User", back_populates="orders")
    daily_menu = db.relationship("DailyMenu", back_populates="orders")

    @property
    def subtotal(self) -> int:
        return self.quantity * self.daily_menu.price


class Setting(db.Model):
    """締切時刻などのアプリ設定を key/value で保持する。"""

    __tablename__ = "settings"

    key = db.Column(db.String(64), primary_key=True)
    value = db.Column(db.String(255), nullable=False, default="")
    updated_at = db.Column(db.DateTime, nullable=False, default=now_jst, onupdate=now_jst)


DEFAULT_SETTINGS = {
    "cutoff_time": "10:00",   # この時刻を過ぎるとその日の注文は締切
    "shop_name": "",          # 発注書に出す仕出し屋名
}


def get_setting(key: str, default: str = "") -> str:
    row = db.session.get(Setting, key)
    if row is None:
        return DEFAULT_SETTINGS.get(key, default)
    return row.value


def set_setting(key: str, value: str) -> None:
    row = db.session.get(Setting, key)
    if row is None:
        row = Setting(key=key, value=value)
        db.session.add(row)
    else:
        row.value = value
        row.updated_at = now_jst()
