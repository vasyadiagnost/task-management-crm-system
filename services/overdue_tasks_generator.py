from __future__ import annotations

from collections import defaultdict
from datetime import datetime, date, time
from typing import Iterable

from database import get_session
from models.task import Task


TASK_STATUS_NEW = "Новая"
TASK_STATUS_DONE = "Выполнена"
TASK_STATUS_CANCELLED = "Отменена"


class OverdueTasksGenerator:
    """
    Генератор задач из просроченных контактов.
    Максимально мягко адаптируется:
    - к разным типам строк (dict / dataclass-like / object with attributes)
    - к разным вариантам модели Task
    """

    def build_groups(self, interaction_rows: Iterable[object]) -> dict[str, list[str]]:
        groups: dict[str, list[str]] = defaultdict(list)

        for row in interaction_rows:
            person = self._row_value(row, "person")
            interaction = self._row_value(row, "interaction")
            responsible = (self._row_value(row, "responsible") or "").strip() or "Без ответственного"

            if person and getattr(person, "fio", None):
                fio = (person.fio or "").strip()
            else:
                person_id = getattr(interaction, "person_id", None) if interaction else None
                fio = f"[ID {person_id}]" if person_id else "Неизвестный контакт"

            if fio:
                groups[responsible].append(fio)

        return {k: v for k, v in groups.items() if v}

    def create_tasks_from_groups(
        self,
        groups: dict[str, list[str]],
        controller: str = "",
        due_date: date | None = None,
        title: str = "Отработать просроченные контакты",
    ) -> dict:
        session = get_session()
        created = 0
        skipped = 0
        details: list[str] = []

        try:
            if due_date is None:
                due_date = date.today()

            due_datetime = datetime.combine(due_date, time.min)

            for responsible, names in groups.items():
                names = sorted(set(name.strip() for name in names if name.strip()))
                if not names:
                    continue

                if self._has_open_duplicate(session, title=title, main_responsible=responsible):
                    skipped += 1
                    details.append(f"{responsible}: пропущено, уже есть открытая задача")
                    continue

                task = Task()

                description = self._build_description(names)
                self._set_any(task, ["title", "name", "task_name"], title)
                self._set_any(task, ["description", "notes", "comment"], description)
                self._set_any(task, ["main_responsible", "responsible", "owner"], responsible)
                self._set_any(task, ["co_executors", "coexecutors", "participants"], "")
                self._set_any(task, ["controller", "control", "supervisor"], controller or "")
                self._set_any(task, ["due_date", "deadline", "planned_date"], due_datetime)
                self._set_any(task, ["status"], TASK_STATUS_NEW)
                self._set_any(task, ["contact_name", "contact", "person_name"], "")
                self._set_any(task, ["created_at"], datetime.utcnow())
                self._set_any(task, ["updated_at"], datetime.utcnow())

                session.add(task)
                session.flush()

                created += 1
                details.append(f"{responsible}: создано {len(names)} контактов")

            session.commit()
            return {
                "created": created,
                "skipped": skipped,
                "details": details,
            }
        finally:
            session.close()

    def _has_open_duplicate(self, session, title: str, main_responsible: str) -> bool:
        query = session.query(Task)

        title_field = self._first_existing_attr(Task, ["title", "name", "task_name"])
        responsible_field = self._first_existing_attr(Task, ["main_responsible", "responsible", "owner"])
        status_field = self._first_existing_attr(Task, ["status"])
        id_field = self._first_existing_attr(Task, ["id"])

        if title_field is not None:
            query = query.filter(title_field == title)
        if responsible_field is not None:
            query = query.filter(responsible_field == main_responsible)
        if status_field is not None:
            query = query.filter(~status_field.in_([TASK_STATUS_DONE, TASK_STATUS_CANCELLED]))

        if id_field is not None:
            task = query.order_by(id_field.desc()).first()
        else:
            task = query.first()

        if not task:
            return False

        status = (getattr(task, "status", "") or "").strip()
        return status not in (TASK_STATUS_DONE, TASK_STATUS_CANCELLED)

    def _build_description(self, names: list[str]) -> str:
        lines = [
            "Прошу отработать просроченные контакты:",
            "",
            "Список:",
        ]
        for fio in names:
            lines.append(f"— {fio}")
        return "\n".join(lines).strip()

    def _set_any(self, obj, field_names: list[str], value):
        for field_name in field_names:
            if hasattr(obj, field_name):
                setattr(obj, field_name, value)
                return True
        return False

    def _first_existing_attr(self, cls, field_names: list[str]):
        for field_name in field_names:
            if hasattr(cls, field_name):
                return getattr(cls, field_name)
        return None

    def _row_value(self, row: object, field_name: str):
        if row is None:
            return None

        if isinstance(row, dict):
            return row.get(field_name)

        if hasattr(row, field_name):
            return getattr(row, field_name)

        try:
            return row[field_name]
        except Exception:
            return None
