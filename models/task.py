from datetime import datetime

from sqlalchemy import Column, Integer, String, Text, Date, DateTime, ForeignKey
from sqlalchemy.orm import relationship

from database import Base


TASK_STATUS_NEW = "Новая"
TASK_STATUS_IN_PROGRESS = "В работе"
TASK_STATUS_DONE = "Выполнена"
TASK_STATUS_CANCELLED = "Отменена"
TASK_STATUS_OVERDUE = "Просрочена"

TASK_STATUS_VALUES = [
    TASK_STATUS_NEW,
    TASK_STATUS_IN_PROGRESS,
    TASK_STATUS_DONE,
    TASK_STATUS_CANCELLED,
    TASK_STATUS_OVERDUE,
]


class Task(Base):
    __tablename__ = "tasks"

    id = Column(Integer, primary_key=True)
    person_id = Column(Integer, ForeignKey("persons.id"), nullable=True)
    title = Column(String(255), nullable=False)
    description = Column(Text, nullable=True)
    main_responsible = Column(String(255), nullable=False, default="")
    co_executors = Column(Text, nullable=True)
    controller = Column(String(255), nullable=True)
    due_date = Column(Date, nullable=True)
    status = Column(String(50), nullable=False, default=TASK_STATUS_NEW)
    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)
    updated_at = Column(DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)

    person = relationship("Person", backref="tasks")