from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any

from database import get_session
from models.reference import Responsible
from models.registry_task import (
    RegistryTask,
    REGISTRY_TASK_STATUS_CANCELLED,
    REGISTRY_TASK_STATUS_DONE,
    REGISTRY_TASK_STATUS_IN_PROGRESS,
    REGISTRY_TASK_STATUS_NEW,
    REGISTRY_TASK_STATUS_OVERDUE,
)


@dataclass
class RegistryTaskRow:
    task: RegistryTask
    status: str

    def __getitem__(self, key: str):
        mapping = {
            "task": self.task,
            "status": self.status,
        }
        return mapping[key]


class RegistryTaskService:
    def compute_status(self, task: RegistryTask) -> str:
        if task.status in (REGISTRY_TASK_STATUS_DONE, REGISTRY_TASK_STATUS_CANCELLED):
            return task.status
        if task.due_date and task.due_date.date() < date.today():
            return REGISTRY_TASK_STATUS_OVERDUE
        return task.status or REGISTRY_TASK_STATUS_NEW

    def list_tasks(self, search: str = "", status: str = "Все", responsible: str = "Все") -> list[RegistryTaskRow]:
        session = get_session()
        try:
            items = (
                session.query(RegistryTask)
                .order_by(RegistryTask.due_date.is_(None), RegistryTask.due_date.asc(), RegistryTask.id.desc())
                .all()
            )
            rows: list[RegistryTaskRow] = []
            for item in items:
                actual_status = self.compute_status(item)
                if search:
                    search_l = search.strip().lower()
                    haystack = " ".join([
                        item.title or "",
                        item.description or "",
                        item.source or "",
                        item.main_responsible or "",
                        item.co_executors or "",
                        item.controller or "",
                        item.comment or "",
                    ]).lower()
                    if search_l not in haystack:
                        continue
                if status and status != "Все" and actual_status != status:
                    continue
                if responsible and responsible != "Все":
                    values = " ".join([item.main_responsible or "", item.co_executors or "", item.controller or ""])
                    if responsible not in values:
                        continue
                rows.append(RegistryTaskRow(task=item, status=actual_status))
            return rows
        finally:
            session.close()

    def get_task(self, task_id: int) -> RegistryTask | None:
        session = get_session()
        try:
            try:
                task_id = int(task_id)
            except (TypeError, ValueError):
                return None
            task = session.get(RegistryTask, task_id)
            if not task:
                return None
            _ = task.id
            _ = task.title
            _ = task.description
            _ = task.source
            _ = task.main_responsible
            _ = task.co_executors
            _ = task.controller
            _ = task.due_date
            _ = task.status
            _ = task.comment
            session.expunge(task)
            return task
        finally:
            session.close()

    def create_task(self, payload: dict[str, Any]) -> RegistryTask:
        session = get_session()
        try:
            task = RegistryTask(
                title=(payload.get("title") or "").strip(),
                description=(payload.get("description") or "").strip(),
                source=(payload.get("source") or "").strip(),
                main_responsible=(payload.get("main_responsible") or "").strip(),
                co_executors=(payload.get("co_executors") or "").strip(),
                controller=(payload.get("controller") or "").strip(),
                due_date=payload.get("due_date"),
                status=(payload.get("status") or REGISTRY_TASK_STATUS_NEW).strip(),
                comment=(payload.get("comment") or "").strip(),
            )
            self.validate_task(task)
            session.add(task)
            session.commit()
            session.refresh(task)
            return task
        finally:
            session.close()

    def update_task(self, task_id: int, payload: dict[str, Any]) -> RegistryTask | None:
        session = get_session()
        try:
            try:
                task_id = int(task_id)
            except (TypeError, ValueError):
                return None
            task = session.get(RegistryTask, task_id)
            if not task:
                return None
            task.title = (payload.get("title") or "").strip()
            task.description = (payload.get("description") or "").strip()
            task.source = (payload.get("source") or "").strip()
            task.main_responsible = (payload.get("main_responsible") or "").strip()
            task.co_executors = (payload.get("co_executors") or "").strip()
            task.controller = (payload.get("controller") or "").strip()
            task.due_date = payload.get("due_date")
            task.status = (payload.get("status") or REGISTRY_TASK_STATUS_NEW).strip()
            task.comment = (payload.get("comment") or "").strip()
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
            task = session.get(RegistryTask, task_id)
            if not task:
                return False
            session.delete(task)
            session.commit()
            return True
        finally:
            session.close()

    def validate_task(self, task: RegistryTask) -> None:
        if not (task.title or "").strip():
            raise ValueError("Нужно указать название поручения")
        if not (task.main_responsible or "").strip():
            raise ValueError("Нужно указать основного ответственного")

    def get_responsible_values(self) -> list[str]:
        session = get_session()
        try:
            values = {"Все"}
            for item in session.query(Responsible).order_by(Responsible.name).all():
                if (item.name or "").strip():
                    values.add(item.name.strip())
            for item in session.query(RegistryTask).all():
                for value in [item.main_responsible, item.co_executors, item.controller]:
                    if (value or "").strip():
                        values.add(value.strip())
            return ["Все"] + sorted(v for v in values if v != "Все")
        finally:
            session.close()

    def get_status_counts(self) -> dict[str, int]:
        rows = self.list_tasks()
        counts = {
            REGISTRY_TASK_STATUS_NEW: 0,
            REGISTRY_TASK_STATUS_IN_PROGRESS: 0,
            REGISTRY_TASK_STATUS_DONE: 0,
            REGISTRY_TASK_STATUS_CANCELLED: 0,
            REGISTRY_TASK_STATUS_OVERDUE: 0,
        }
        for row in rows:
            counts[row["status"]] = counts.get(row["status"], 0) + 1
        return counts
