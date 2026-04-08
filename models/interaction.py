from sqlalchemy import Date, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from database import Base


class Interaction(Base):
    __tablename__ = "interactions"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)

    person_id: Mapped[int] = mapped_column(ForeignKey("persons.id"), nullable=False)

    interaction_type: Mapped[str] = mapped_column(String(50), nullable=False, default="Звонок")
    interaction_date: Mapped[Date | None] = mapped_column(Date, nullable=True)

    purpose: Mapped[str] = mapped_column(Text, nullable=False, default="")
    result: Mapped[str] = mapped_column(Text, nullable=False, default="")
    next_date: Mapped[Date | None] = mapped_column(Date, nullable=True)

    responsible: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    comment: Mapped[str] = mapped_column(Text, nullable=False, default="")

    is_active: Mapped[int] = mapped_column(Integer, nullable=False, default=1)
    completed_at: Mapped[DateTime | None] = mapped_column(DateTime, nullable=True)

    def __repr__(self) -> str:
        return (
            f"<Interaction(id={self.id}, person_id={self.person_id}, "
            f"type='{self.interaction_type}', is_active={self.is_active})>"
        )