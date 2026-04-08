from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from database import Base


REGISTRY_TASK_STATUS_NEW = "Новая"
REGISTRY_TASK_STATUS_IN_PROGRESS = "В работе"
REGISTRY_TASK_STATUS_DONE = "Выполнена"
REGISTRY_TASK_STATUS_CANCELLED = "Отменена"
REGISTRY_TASK_STATUS_OVERDUE = "Просрочена"

REGISTRY_TASK_STATUS_VALUES = [
    REGISTRY_TASK_STATUS_NEW,
    REGISTRY_TASK_STATUS_IN_PROGRESS,
    REGISTRY_TASK_STATUS_DONE,
    REGISTRY_TASK_STATUS_CANCELLED,
]


class RegistryTask(Base):
    __tablename__ = "registry_tasks"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    title: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    description: Mapped[str] = mapped_column(Text, nullable=False, default="")
    source: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    main_responsible: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    co_executors: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    controller: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    due_date: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    status: Mapped[str] = mapped_column(String(50), nullable=False, default=REGISTRY_TASK_STATUS_NEW)
    comment: Mapped[str] = mapped_column(Text, nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    def __repr__(self) -> str:
        return f"<RegistryTask(id={self.id}, title='{self.title}')>"
