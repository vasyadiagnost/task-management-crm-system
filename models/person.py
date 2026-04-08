from sqlalchemy import Date, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from database import Base


class Person(Base):
    __tablename__ = "persons"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)

    # Базовые поля системы
    fio: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    position: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    department: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    phone: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    circle: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    birthday: Mapped[Date | None] = mapped_column(Date, nullable=True)
    responsible: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    notes: Mapped[str] = mapped_column(Text, nullable=False, default="")

    # Расширенная структура по "Таблица ВВ"
    level: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    subcategory: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    legacy_column_1: Mapped[str] = mapped_column(Text, nullable=False, default="")
    track_calls: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    gender: Mapped[str] = mapped_column(String(50), nullable=False, default="")
    gift_category: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    legacy_column1: Mapped[str] = mapped_column(Text, nullable=False, default="")
    birthday_today_flag: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    call_periodicity_legacy: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    spouse_birthday: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    children_birthdays: Mapped[str] = mapped_column(Text, nullable=False, default="")
    served: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    prof_holiday: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    legacy_column2: Mapped[str] = mapped_column(Text, nullable=False, default="")
    prof_holiday_flag: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    religion: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    hobby: Mapped[str] = mapped_column(Text, nullable=False, default="")
    phone_contact_person: Mapped[str] = mapped_column(Text, nullable=False, default="")
    food_preferences: Mapped[str] = mapped_column(Text, nullable=False, default="")
    avoid_with: Mapped[str] = mapped_column(Text, nullable=False, default="")
    gifts_already_given: Mapped[str] = mapped_column(Text, nullable=False, default="")
    last_meeting_legacy: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    next_meeting_legacy: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    meeting_purpose_legacy: Mapped[str] = mapped_column(Text, nullable=False, default="")
    need_meeting_flag: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    last_call_legacy: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    next_call_legacy: Mapped[str] = mapped_column(String(100), nullable=False, default="")
    call_purpose_legacy: Mapped[str] = mapped_column(Text, nullable=False, default="")
    need_call_flag: Mapped[str] = mapped_column(String(100), nullable=False, default="")

    def __repr__(self) -> str:
        return f"<Person(id={self.id}, fio='{self.fio}')>"