from __future__ import annotations

from datetime import date, timedelta
from typing import Any

from sqlalchemy.orm import joinedload

from database import get_session
from models.person import Person
from models.reference import Responsible
from models.task import (
    Task,
    TASK_STATUS_NEW,
    TASK_STATUS_IN_PROGRESS,
    TASK_STATUS_DONE,
    TASK_STATUS_CANCELLED,
    TASK_STATUS_OVERDUE,
)


_META_START = "[[EXECUTION_META]]"
_META_END = "[[/EXECUTION_META]]"


class TaskService:
    def list_tasks(
        self,
        search: str = "",
        status: str = "Все",
        responsible: str = "Все",
    ) -> list[dict[str, Any]]:
        session = get_session()
        try:
            query = (
                session.query(Task)
                .options(joinedload(Task.person))
                .outerjoin(Person, Task.person_id == Person.id)
            )

            if search:
                search_like = f"%{search.strip()}%"
                query = query.filter(
                    (Task.title.ilike(search_like))
                    | (Task.main_responsible.ilike(search_like))
                    | (Task.controller.ilike(search_like))
                    | (Person.fio.ilike(search_like))
                )

            items = query.order_by(Task.due_date.is_(None), Task.due_date.asc(), Task.id.desc()).all()
            rows: list[dict[str, Any]] = []
            for item in items:
                actual_status = self.compute_status(item)

                if status and status != "Все" and actual_status != status:
                    continue
                if responsible and responsible != "Все" and item.main_responsible != responsible:
                    continue

                rows.append(
                    {
                        "task": item,
                        "person": item.person,
                        "status": actual_status,
                    }
                )
            return rows
        finally:
            session.close()

    def get_task(self, task_id: int) -> Task | None:
        session = get_session()
        try:
            try:
                task_id = int(task_id)
            except (TypeError, ValueError):
                return None

            task = (
                session.query(Task)
                .options(joinedload(Task.person))
                .filter(Task.id == task_id)
                .first()
            )
            if not task:
                return None

            _ = task.id
            _ = task.person_id
            _ = task.title
            _ = task.description
            _ = task.main_responsible
            _ = task.co_executors
            _ = task.controller
            _ = task.due_date
            _ = task.status
            if task.person:
                _ = task.person.id
                _ = task.person.fio

            session.expunge(task)
            if task.person:
                session.expunge(task.person)
            return task
        finally:
            session.close()

    def create_task(self, payload: dict[str, Any]) -> Task:
        session = get_session()
        try:
            task = Task(
                person_id=payload.get("person_id"),
                title=(payload.get("title") or "").strip(),
                description=(payload.get("description") or "").strip(),
                main_responsible=(payload.get("main_responsible") or "").strip(),
                co_executors=(payload.get("co_executors") or "").strip(),
                controller=(payload.get("controller") or "").strip(),
                due_date=payload.get("due_date"),
                status=(payload.get("status") or TASK_STATUS_NEW).strip(),
            )
            self.validate_task(task)
            session.add(task)
            session.commit()
            session.refresh(task)
            return task
        finally:
            session.close()

    def update_task(self, task_id: int, payload: dict[str, Any]) -> Task | None:
        session = get_session()
        try:
            try:
                task_id = int(task_id)
            except (TypeError, ValueError):
                return None

            task = session.get(Task, task_id)
            if not task:
                return None

            task.person_id = payload.get("person_id")
            task.title = (payload.get("title") or "").strip()
            task.description = (payload.get("description") or "").strip()
            task.main_responsible = (payload.get("main_responsible") or "").strip()
            task.co_executors = (payload.get("co_executors") or "").strip()
            task.controller = (payload.get("controller") or "").strip()
            task.due_date = payload.get("due_date")
            task.status = (payload.get("status") or TASK_STATUS_NEW).strip()

            self.validate_task(task)
            session.commit()
            session.refresh(task)
            return task
        finally:
            session.close()

    def delete_task(self, task_id: int) -> bool:
        session = get_session()
        try:
            try:
                task_id = int(task_id)
            except (TypeError, ValueError):
                return False

            task = session.get(Task, task_id)
            if not task:
                return False

            session.delete(task)
            session.commit()
            return True
        finally:
            session.close()

    def validate_task(self, task: Task) -> None:
        if not (task.title or "").strip():
            raise ValueError("Нужно указать название задачи")
        if not (task.main_responsible or "").strip():
            raise ValueError("Нужно указать основного ответственного")

    def compute_status(self, task: Task) -> str:
        if task.status in (TASK_STATUS_DONE, TASK_STATUS_CANCELLED):
            return task.status
        if task.due_date and task.due_date < date.today():
            return TASK_STATUS_OVERDUE
        return task.status or TASK_STATUS_NEW

    def get_status_counts(self) -> dict[str, int]:
        rows = self.list_tasks()
        counts = {
            "Все": len(rows),
            TASK_STATUS_NEW: 0,
            TASK_STATUS_IN_PROGRESS: 0,
            TASK_STATUS_DONE: 0,
            TASK_STATUS_CANCELLED: 0,
            TASK_STATUS_OVERDUE: 0,
        }
        for row in rows:
            counts[row["status"]] = counts.get(row["status"], 0) + 1
        return counts

    def get_responsible_values(self) -> list[str]:
        session = get_session()
        try:
            values = {"Все"}
            for item in session.query(Responsible).order_by(Responsible.name).all():
                if (item.name or "").strip():
                    values.add(item.name.strip())
            for item in session.query(Task).all():
                if (item.main_responsible or "").strip():
                    values.add(item.main_responsible.strip())
            return ["Все"] + sorted(v for v in values if v != "Все")
        finally:
            session.close()

    def get_person_records(self) -> list[dict[str, Any]]:
        session = get_session()
        try:
            items = session.query(Person).order_by(Person.fio).all()
            return [
                {"id": item.id, "fio": (item.fio or "").strip()}
                for item in items
                if (item.fio or "").strip()
            ]
        finally:
            session.close()

    def get_overdue_tasks(self) -> list[dict[str, Any]]:
        return self.list_tasks(status=TASK_STATUS_OVERDUE)

    def get_today_tasks(self) -> list[dict[str, Any]]:
        session = get_session()
        try:
            items = (
                session.query(Task)
                .options(joinedload(Task.person))
                .filter(Task.due_date == date.today())
                .order_by(Task.id.desc())
                .all()
            )
            rows = []
            for item in items:
                rows.append({"task": item, "person": item.person, "status": self.compute_status(item)})
            return rows
        finally:
            session.close()

    def get_upcoming_tasks(self, days: int = 7) -> list[dict[str, Any]]:
        start_day = date.today()
        end_day = start_day + timedelta(days=days)
        session = get_session()
        try:
            items = (
                session.query(Task)
                .options(joinedload(Task.person))
                .filter(Task.due_date.is_not(None))
                .filter(Task.due_date >= start_day)
                .filter(Task.due_date <= end_day)
                .order_by(Task.due_date.asc(), Task.id.desc())
                .all()
            )
            rows = []
            for item in items:
                rows.append({"task": item, "person": item.person, "status": self.compute_status(item)})
            return rows
        finally:
            session.close()

    def parse_execution_people(self, description: str) -> list[str]:
        state = self.parse_execution_state(description)
        return self.merge_execution_people(state["remaining"], state["done"])

    def merge_execution_people(self, remaining_people: list[str], done_people: list[str]) -> list[str]:
        seen = set()
        ordered: list[str] = []
        for item in list(remaining_people) + list(done_people):
            fio = (item or "").strip()
            if fio and fio not in seen:
                seen.add(fio)
                ordered.append(fio)
        return ordered

    def parse_execution_state(self, description: str) -> dict[str, Any]:
        text = description or ""
        metadata_text = ""
        visible_text = text

        if _META_START in text and _META_END in text:
            before, _, tail = text.partition(_META_START)
            meta_block, _, _ = tail.partition(_META_END)
            visible_text = before.strip()
            metadata_text = meta_block.strip()

        remaining: list[str] = []
        done: list[str] = []
        interaction_ids: dict[str, int] = {}
        interaction_dates: dict[str, str] = {}

        mode = "legacy"
        for raw_line in visible_text.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            lowered = line.lower()
            if "к отработке" in lowered:
                mode = "remaining"
                continue
            if "уже отработано" in lowered:
                mode = "done"
                continue

            if line.startswith("—"):
                fio = line[1:].strip()
                if not fio or fio.lower() == "список закрыт":
                    continue
                if mode == "done":
                    done.append(fio)
                else:
                    remaining.append(fio)
            elif line.startswith("-"):
                fio = line[1:].strip()
                if not fio or fio.lower() == "список закрыт":
                    continue
                if mode == "done":
                    done.append(fio)
                else:
                    remaining.append(fio)
            elif line.startswith("✓"):
                fio = line[1:].strip()
                if fio:
                    done.append(fio)

        for raw_line in metadata_text.splitlines():
            line = raw_line.strip()
            if not line or "|||" not in line:
                continue
            parts = [part.strip() for part in line.split("|||")]
            if len(parts) < 2:
                continue
            fio = parts[0]
            if not fio:
                continue
            try:
                interaction_ids[fio] = int(parts[1])
            except ValueError:
                continue
            if len(parts) >= 3 and parts[2]:
                interaction_dates[fio] = parts[2]

        remaining = self._unique_ordered(remaining)
        done = [fio for fio in self._unique_ordered(done) if fio not in remaining]

        return {
            "remaining": remaining,
            "done": done,
            "interaction_ids": interaction_ids,
            "interaction_dates": interaction_dates,
        }

    def build_execution_description(
        self,
        remaining_people: list[str],
        done_people: list[str],
        interaction_ids: dict[str, int] | None = None,
        interaction_dates: dict[str, str] | None = None,
    ) -> str:
        remaining = self._unique_ordered(remaining_people)
        done = [fio for fio in self._unique_ordered(done_people) if fio not in remaining]
        interaction_ids = interaction_ids or {}
        interaction_dates = interaction_dates or {}

        parts = [
            "Прошу отработать просроченные контакты:",
            "",
            "К отработке:",
        ]

        if remaining:
            for fio in remaining:
                parts.append(f"— {fio}")
        else:
            parts.append("— список закрыт")

        if done:
            parts.extend(["", "Уже отработано:"])
            for fio in done:
                parts.append(f"✓ {fio}")

        meta_lines = []
        for fio in done:
            interaction_id = interaction_ids.get(fio)
            if interaction_id:
                raw_date = interaction_dates.get(fio, "")
                meta_lines.append(f"{fio}|||{interaction_id}|||{raw_date}")

        if meta_lines:
            parts.extend(["", _META_START])
            parts.extend(meta_lines)
            parts.append(_META_END)

        return "\n".join(parts).strip()

    def is_execution_task(self, task: Task) -> bool:
        state = self.parse_execution_state(task.description or "")
        return len(state["remaining"]) > 0 or len(state["done"]) > 0

    def mark_task_done(self, task_id: int) -> Task | None:
        session = get_session()
        try:
            try:
                task_id = int(task_id)
            except (TypeError, ValueError):
                return None

            task = session.get(Task, task_id)
            if not task:
                return None

            task.status = TASK_STATUS_DONE
            session.commit()
            session.refresh(task)
            return task
        finally:
            session.close()

    def save_execution_progress(
        self,
        task_id: int,
        remaining_people: list[str],
        done_people: list[str],
        interaction_ids: dict[str, int] | None = None,
        interaction_dates: dict[str, str] | None = None,
        mark_done: bool = False,
    ) -> Task | None:
        session = get_session()
        try:
            try:
                task_id = int(task_id)
            except (TypeError, ValueError):
                return None

            task = session.get(Task, task_id)
            if not task:
                return None

            remaining = self._unique_ordered(remaining_people)
            done = [fio for fio in self._unique_ordered(done_people) if fio not in remaining]
            interaction_ids = interaction_ids or {}
            interaction_dates = interaction_dates or {}

            task.description = self.build_execution_description(remaining, done, interaction_ids, interaction_dates)

            if mark_done or not remaining:
                task.status = TASK_STATUS_DONE
            elif done:
                task.status = TASK_STATUS_IN_PROGRESS
            else:
                task.status = TASK_STATUS_NEW

            session.commit()
            session.refresh(task)
            return task
        finally:
            session.close()

    def _unique_ordered(self, items: list[str]) -> list[str]:
        seen = set()
        result: list[str] = []
        for item in items:
            value = (item or "").strip()
            if value and value not in seen:
                seen.add(value)
                result.append(value)
        return result
