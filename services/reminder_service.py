from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any

from database import get_session
from models.interaction import Interaction
from models.person import Person
from models.reference import Circle


@dataclass
class ReminderItem:
    kind: str  # 'birthday' | 'contact'
    title: str
    person_id: int | None
    fio: str
    event_date: date
    responsible: str
    circle: str
    subtitle: str = ""
    status: str = ""
    interaction_id: int | None = None


class ReminderService:
    def __init__(self, interaction_service=None):
        self.interaction_service = interaction_service

    @staticmethod
    def _safe_strip(value: Any) -> str:
        return (value or "").strip() if isinstance(value, str) else ("" if value is None else str(value).strip())

    @staticmethod
    def _normalize_birthday_for_year(birthday_value: date | None, year: int) -> date | None:
        if not birthday_value:
            return None
        try:
            return date(year, birthday_value.month, birthday_value.day)
        except ValueError:
            if birthday_value.month == 2 and birthday_value.day == 29:
                return date(year, 2, 28)
            return None

    def _birthday_items_between(self, start_day: date, end_day: date) -> list[ReminderItem]:
        session = get_session()
        people = session.query(Person).all()
        session.close()

        result: list[ReminderItem] = []
        current = start_day
        while current <= end_day:
            for person in people:
                bday = self._normalize_birthday_for_year(person.birthday, current.year)
                if bday != current:
                    continue
                result.append(
                    ReminderItem(
                        kind="birthday",
                        title="День рождения",
                        person_id=person.id,
                        fio=self._safe_strip(person.fio) or "Неизвестный контакт",
                        event_date=current,
                        responsible=self._safe_strip(person.responsible),
                        circle=self._safe_strip(person.circle),
                        subtitle="День рождения",
                        status="Сегодня" if current == start_day else f"Через {(current - start_day).days} дн.",
                    )
                )
            current += timedelta(days=1)

        result.sort(key=lambda item: (item.event_date, item.fio.lower()))
        return result

    def _contact_items_between(self, start_day: date, end_day: date) -> list[ReminderItem]:
        session = get_session()
        interactions = session.query(Interaction).filter(Interaction.is_active == 1).all()

        person_ids = {item.person_id for item in interactions if item.person_id}
        persons = session.query(Person).filter(Person.id.in_(person_ids)).all() if person_ids else []
        person_map = {person.id: person for person in persons}
        session.close()

        result: list[ReminderItem] = []
        for interaction in interactions:
            if not interaction.next_date:
                continue
            next_day = interaction.next_date
            if not (start_day <= next_day <= end_day):
                continue

            person = person_map.get(interaction.person_id)
            fio = self._safe_strip(person.fio if person else "") or "Неизвестный контакт"
            responsible = self._safe_strip(interaction.responsible) or self._safe_strip(person.responsible if person else "")
            circle = self._safe_strip(person.circle if person else "")
            delta_days = (next_day - start_day).days
            if delta_days < 0:
                status = "Просрочен"
            elif delta_days == 0:
                status = "Сегодня"
            elif delta_days <= 7:
                status = "7 дней"
            else:
                status = "Запланирован"

            result.append(
                ReminderItem(
                    kind="contact",
                    title="Контакт",
                    person_id=interaction.person_id,
                    fio=fio,
                    event_date=next_day,
                    responsible=responsible,
                    circle=circle,
                    subtitle=self._safe_strip(interaction.interaction_type) or "Контакт",
                    status=status,
                    interaction_id=interaction.id,
                )
            )

        result.sort(key=lambda item: (item.event_date, item.fio.lower()))
        return result

    def get_today(self) -> dict[str, list[ReminderItem]]:
        today = date.today()
        return {
            "birthdays": self._birthday_items_between(today, today),
            "contacts": self._contact_items_between(today, today),
        }

    def get_next_days(self, days: int = 7) -> dict[str, list[ReminderItem]]:
        today = date.today()
        end_day = today + timedelta(days=days)
        return {
            "birthdays": self._birthday_items_between(today, end_day),
            "contacts": self._contact_items_between(today, end_day),
        }

    def get_dashboard_payload(self, days: int = 7, responsible: str = "Все", circle: str = "Все") -> dict[str, Any]:
        today_payload = self.get_today()
        week_payload = self.get_next_days(days=days)

        def _filter(items: list[ReminderItem]) -> list[ReminderItem]:
            filtered = items
            if responsible and responsible != "Все":
                filtered = [item for item in filtered if item.responsible == responsible]
            if circle and circle != "Все":
                filtered = [item for item in filtered if item.circle == circle]
            return filtered

        birthdays_today = _filter(today_payload["birthdays"])
        contacts_today = _filter(today_payload["contacts"])
        birthdays_week = _filter(week_payload["birthdays"])
        contacts_week = _filter(week_payload["contacts"])

        return {
            "birthdays_today": birthdays_today,
            "contacts_today": contacts_today,
            "birthdays_week": birthdays_week,
            "contacts_week": contacts_week,
            "today_total": len(birthdays_today) + len(contacts_today),
            "week_total": len(birthdays_week) + len(contacts_week),
        }

    def get_filter_values(self) -> dict[str, list[str]]:
        session = get_session()
        responsibles = {"Все"}
        circles = {"Все"}

        for person in session.query(Person).all():
            responsible = self._safe_strip(person.responsible)
            if responsible:
                responsibles.add(responsible)

        for circle in session.query(Circle).order_by(Circle.id).all():
            name = self._safe_strip(getattr(circle, "name", None))
            if name:
                circles.add(name)

        session.close()
        return {
            "responsibles": ["Все"] + sorted(v for v in responsibles if v != "Все"),
            "circles": ["Все"] + [v for v in sorted(circles, key=lambda x: (x == "Все", x)) if v != "Все"],
        }

    def generate_day_order_message(self, responsible: str = "Все", circle: str = "Все") -> str:
        payload = self.get_dashboard_payload(days=7, responsible=responsible, circle=circle)
        birthdays_today = payload["birthdays_today"]
        contacts_today = payload["contacts_today"]

        lines = ["Добрый день! Разнарядка по контактам на сегодня.", ""]

        if birthdays_today:
            lines.append(f"Дни рождения сегодня: {len(birthdays_today)}")
            for item in birthdays_today:
                responsible_text = item.responsible or "-"
                lines.append(f"— {item.fio} ({item.subtitle or 'День рождения'}, {responsible_text})")
            lines.append("")
        else:
            lines.append("Дней рождения на сегодня не запланировано.")
            lines.append("")

        if contacts_today:
            lines.append(f"Контакты на сегодня: {len(contacts_today)}")
            for item in contacts_today:
                responsible_text = item.responsible or "-"
                lines.append(f"— {item.fio} ({item.subtitle or 'Контакт'}, {responsible_text})")
        else:
            lines.append("На сегодня контактов не запланировано.")

        return "\n".join(lines).strip() + "\n"
