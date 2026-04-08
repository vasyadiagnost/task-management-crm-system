from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database import Base


MEETING_STATUS_SCHEDULED = "Запланирована"
MEETING_STATUS_COMPLETED = "Проведена"
MEETING_STATUS_CANCELLED = "Отменена"
MEETING_STATUS_OVERDUE = "Просрочена"

MEETING_STATUS_VALUES = [
    MEETING_STATUS_SCHEDULED,
    MEETING_STATUS_COMPLETED,
    MEETING_STATUS_CANCELLED,
    MEETING_STATUS_OVERDUE,
]


class Meeting(Base):
    __tablename__ = "meetings"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    person_id: Mapped[int | None] = mapped_column(ForeignKey("persons.id"), nullable=True)
    subject: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    location: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    start_datetime: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    end_datetime: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    recurrence_rule: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    status: Mapped[str] = mapped_column(String(50), nullable=False, default=MEETING_STATUS_SCHEDULED)
    notes: Mapped[str] = mapped_column(Text, nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    person = relationship("Person", backref="meetings")

    def __repr__(self) -> str:
        return f"<Meeting(id={self.id}, subject='{self.subject}')>"
