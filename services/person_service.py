from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any

from sqlalchemy import or_

from database import get_session
from models.interaction import Interaction
from models.person import Person
from ui.statuses import (
    STATUS_NO_DATE,
    STATUS_OVERDUE,
    STATUS_PLANNED,
    STATUS_TODAY,
    STATUS_7_DAYS,
)


@dataclass
class NextActionInfo:
    person_id: int
    interaction_id: int | None
    action_type: str
    next_date: date | None
    status: str
    responsible: str
    purpose: str


class PersonService:
    """Сервис работы с людьми.

    Сейчас сделан максимально простым и прозрачным:
    - без репозиториев
    - без сложных DTO
    - с обычной работой через SQLAlchemy session
    """

    def list_persons(
        self,
        search: str | None = None,
        circle: str | None = None,
        responsible: str | None = None,
        only_without_birthday: bool = False,
    ) -> list[Person]:
        session = get_session()
        try:
            query = session.query(Person)

            if search:
                value = f"%{search.strip()}%"
                query = query.filter(Person.fio.ilike(value))

            if circle and circle != "Все":
                query = query.filter(Person.circle == circle)

            if responsible and responsible != "Все":
                query = query.filter(Person.responsible == responsible)

            if only_without_birthday:
                query = query.filter(Person.birthday.is_(None))

            return query.order_by(Person.fio.asc()).all()
        finally:
            session.close()

    def search_persons(self, search: str) -> list[Person]:
        return self.list_persons(search=search)

    def get_person(self, person_id: int) -> Person | None:
        session = get_session()
        try:
            return session.query(Person).filter(Person.id == person_id).first()
        finally:
            session.close()

    def create_person(self, data: dict[str, Any]) -> Person:
        session = get_session()
        try:
            payload = self._normalize_person_payload(data)
            person = Person(**payload)
            session.add(person)
            session.commit()
            session.refresh(person)
            return person
        finally:
            session.close()

    def update_person(self, person_id: int, data: dict[str, Any]) -> Person | None:
        session = get_session()
        try:
            person = session.query(Person).filter(Person.id == person_id).first()
            if not person:
                return None

            payload = self._normalize_person_payload(data)
            for key, value in payload.items():
                if hasattr(person, key):
                    setattr(person, key, value)

            session.commit()
            session.refresh(person)
            return person
        finally:
            session.close()

    def delete_person(self, person_id: int) -> bool:
        session = get_session()
        try:
            person = session.query(Person).filter(Person.id == person_id).first()
            if not person:
                return False

            session.delete(person)
            session.commit()
            return True
        finally:
            session.close()

    def get_person_next_action(self, person_id: int) -> NextActionInfo | None:
        session = get_session()
        try:
            interaction = (
                session.query(Interaction)
                .filter(
                    Interaction.person_id == person_id,
                    Interaction.is_active == 1,
                )
                .order_by(Interaction.next_date.asc().nullsfirst(), Interaction.id.desc())
                .first()
            )

            if not interaction:
                return None

            responsible = (interaction.responsible or "").strip()
            status = self._compute_status(interaction.next_date)
            purpose = (interaction.purpose or "").strip()

            return NextActionInfo(
                person_id=person_id,
                interaction_id=interaction.id,
                action_type=interaction.interaction_type or "",
                next_date=interaction.next_date,
                status=status,
                responsible=responsible,
                purpose=purpose,
            )
        finally:
            session.close()

    def get_upcoming_birthdays(self, days: int = 7) -> list[Person]:
        people = self.list_persons()
        today = date.today()
        result: list[tuple[int, Person]] = []

        for person in people:
            if not person.birthday:
                continue
            next_bday = self._next_birthday_date(person.birthday, today)
            delta = (next_bday - today).days
            if 0 <= delta <= days:
                result.append((delta, person))

        result.sort(key=lambda item: ((self._next_birthday_date(item[1].birthday, today) - today).days, item[1].fio.lower()))
        return [item[1] for item in result]

    def get_birthdays_today(self) -> list[Person]:
        return self.get_upcoming_birthdays(days=0)

    def count_persons(self, search: str | None = None) -> int:
        session = get_session()
        try:
            query = session.query(Person)
            if search:
                value = f"%{search.strip()}%"
                query = query.filter(Person.fio.ilike(value))
            return query.count()
        finally:
            session.close()

    @staticmethod
    def _normalize_person_payload(data: dict[str, Any]) -> dict[str, Any]:
        payload = dict(data)

        # Совместимость UI -> текущая модель Person
        if "comment" in payload and "notes" not in payload:
            payload["notes"] = payload.pop("comment")

        # Обрезаем пробелы в строковых полях
        for key, value in list(payload.items()):
            if isinstance(value, str):
                payload[key] = value.strip()

        # Не передаём в SQLAlchemy поля, которых нет в модели
        valid_keys = set(Person.__table__.columns.keys())
        payload = {k: v for k, v in payload.items() if k in valid_keys}
        return payload

    @staticmethod
    def _compute_status(next_date: date | None) -> str:
        if not next_date:
            return STATUS_NO_DATE

        today = date.today()
        if next_date < today:
            return STATUS_OVERDUE
        if next_date == today:
            return STATUS_TODAY
        if today < next_date <= today + timedelta(days=7):
            return STATUS_7_DAYS
        return STATUS_PLANNED

    @staticmethod
    def _next_birthday_date(birthday: date | None, today: date) -> date:
        if birthday is None:
            return today

        month = birthday.month
        day = birthday.day

        # 29 февраля -> 28 февраля в невисокосный год
        try:
            candidate = date(today.year, month, day)
        except ValueError:
            candidate = date(today.year, 2, 28)

        if candidate < today:
            try:
                candidate = date(today.year + 1, month, day)
            except ValueError:
                candidate = date(today.year + 1, 2, 28)

        return candidate
