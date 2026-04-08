from __future__ import annotations

from datetime import date
from typing import Any

from sqlalchemy.orm import joinedload

from database import get_session
from models.meeting import Meeting
from models.person import Person


MEETING_STATUS_ACTIVE = "Активна"
MEETING_STATUS_PAUSED = "На паузе"
MEETING_STATUS_DONE = "Проведена"
MEETING_STATUS_CANCELLED = "Отменена"
MEETING_STATUS_OVERDUE = "Просрочена"

MEETING_STATUS_VALUES = [
    MEETING_STATUS_ACTIVE,
    MEETING_STATUS_PAUSED,
    MEETING_STATUS_DONE,
    MEETING_STATUS_CANCELLED,
]


class MeetingService:
    def _resolve_contact_name(self, meeting: Meeting) -> str:
        if getattr(meeting, "person", None) and (meeting.person.fio or "").strip():
            return meeting.person.fio.strip()
        if (meeting.subject or "").strip():
            return meeting.subject.strip()
        return ""

    def _row_sort_key(self, row: dict[str, Any]):
        meeting = row["meeting"]
        days_left = row["days_left"]
        status = row["status"]
        if status == MEETING_STATUS_OVERDUE:
            return (0, meeting.start_datetime or date.max, self._resolve_contact_name(meeting).lower())
        if days_left is not None and days_left <= 9:
            return (1, meeting.start_datetime or date.max, self._resolve_contact_name(meeting).lower())
        return (2, meeting.start_datetime or date.max, self._resolve_contact_name(meeting).lower())

    def list_meetings(self, search: str = "", status: str = "Все") -> list[dict[str, Any]]:
        session = get_session()
        try:
            meetings = (
                session.query(Meeting)
                .options(joinedload(Meeting.person))
                .order_by(Meeting.start_datetime.asc().nulls_last(), Meeting.id.desc())
                .all()
            )
            rows: list[dict[str, Any]] = []
            for item in meetings:
                contact_name = self._resolve_contact_name(item)
                actual_status = self.compute_status(item)
                recurrence = (item.recurrence_rule or "").strip()

                if search:
                    search_value = search.strip().lower()
                    haystack = " ".join([contact_name, recurrence, item.notes or "", item.subject or "", item.location or ""]).lower()
                    if search_value not in haystack:
                        continue

                if status and status != "Все" and actual_status != status:
                    continue

                rows.append(
                    {
                        "meeting": item,
                        "contact_name": contact_name,
                        "status": actual_status,
                        "days_left": self.get_days_left(item),
                    }
                )
            rows.sort(key=self._row_sort_key)
            return rows
        finally:
            session.close()

    def get_meeting(self, meeting_id: int) -> Meeting | None:
        session = get_session()
        try:
            try:
                meeting_id = int(meeting_id)
            except (TypeError, ValueError):
                return None

            meeting = (
                session.query(Meeting)
                .options(joinedload(Meeting.person))
                .filter(Meeting.id == meeting_id)
                .first()
            )
            if not meeting:
                return None

            _ = meeting.id
            _ = meeting.person_id
            _ = meeting.subject
            _ = meeting.location
            _ = meeting.start_datetime
            _ = meeting.end_datetime
            _ = meeting.recurrence_rule
            _ = meeting.status
            _ = meeting.notes
            if meeting.person:
                _ = meeting.person.id
                _ = meeting.person.fio
                session.expunge(meeting.person)
            session.expunge(meeting)
            return meeting
        finally:
            session.close()

    def create_meeting(self, payload: dict[str, Any]) -> Meeting:
        session = get_session()
        try:
            meeting = Meeting(
                person_id=payload.get("person_id"),
                subject=(payload.get("subject") or "").strip(),
                location="",
                start_datetime=payload.get("planning_deadline"),
                end_datetime=None,
                recurrence_rule=(payload.get("recurrence_rule") or "").strip(),
                status=(payload.get("status") or MEETING_STATUS_ACTIVE).strip(),
                notes=(payload.get("notes") or "").strip(),
            )
            self.validate_meeting(meeting)
            session.add(meeting)
            session.commit()
            session.refresh(meeting)
            return meeting
        finally:
            session.close()

    def update_meeting(self, meeting_id: int, payload: dict[str, Any]) -> Meeting | None:
        session = get_session()
        try:
            try:
                meeting_id = int(meeting_id)
            except (TypeError, ValueError):
                return None

            meeting = session.get(Meeting, meeting_id)
            if not meeting:
                return None

            meeting.person_id = payload.get("person_id")
            meeting.subject = (payload.get("subject") or "").strip()
            meeting.start_datetime = payload.get("planning_deadline")
            meeting.recurrence_rule = (payload.get("recurrence_rule") or "").strip()
            meeting.status = (payload.get("status") or MEETING_STATUS_ACTIVE).strip()
            meeting.notes = (payload.get("notes") or "").strip()
            meeting.location = meeting.location or ""
            meeting.end_datetime = None

            self.validate_meeting(meeting)
            session.commit()
            session.refresh(meeting)
            return meeting
        finally:
            session.close()

    def delete_meeting(self, meeting_id: int) -> bool:
        session = get_session()
        try:
            try:
                meeting_id = int(meeting_id)
            except (TypeError, ValueError):
                return False
            meeting = session.get(Meeting, meeting_id)
            if not meeting:
                return False
            session.delete(meeting)
            session.commit()
            return True
        finally:
            session.close()

    def get_upcoming_meetings(self, days: int = 9) -> list[dict[str, Any]]:
        rows = self.list_meetings()
        result = []
        for row in rows:
            days_left = row["days_left"]
            if days_left is None:
                continue
            if days_left <= days and row["status"] not in (MEETING_STATUS_DONE, MEETING_STATUS_CANCELLED):
                result.append(row)
        result.sort(key=self._row_sort_key)
        return result

    def get_status_counts(self) -> dict[str, int]:
        rows = self.list_meetings()
        counts = {
            MEETING_STATUS_OVERDUE: 0,
            MEETING_STATUS_ACTIVE: 0,
            MEETING_STATUS_PAUSED: 0,
            MEETING_STATUS_DONE: 0,
            MEETING_STATUS_CANCELLED: 0,
        }
        for row in rows:
            counts[row["status"]] = counts.get(row["status"], 0) + 1
        return counts

    def get_person_records(self) -> list[dict[str, Any]]:
        session = get_session()
        try:
            items = session.query(Person).order_by(Person.fio).all()
            return [{"id": item.id, "fio": (item.fio or "").strip()} for item in items if (item.fio or "").strip()]
        finally:
            session.close()

    def generate_summary(self, days: int = 9) -> str:
        upcoming = self.get_upcoming_meetings(days=days)
        lines = ["Добрый день! Напоминание по регулярным встречам.", ""]
        if not upcoming:
            lines.append("На ближайшие 9 дней встреч для планирования нет.")
            return "\n".join(lines) + "\n"

        lines.append(f"Встречи ближайшей недели: {len(upcoming)}")
        lines.append("")
        for row in upcoming:
            meeting = row["meeting"]
            contact_name = row["contact_name"] or "Без контакта"
            recurrence = (meeting.recurrence_rule or "-").strip()
            days_left = row["days_left"]
            days_text = "сегодня" if days_left == 0 else f"через {days_left} дн."
            lines.append(f"— {contact_name} | периодичность: {recurrence} | планировать {days_text}")
        return "\n".join(lines).strip() + "\n"

    def validate_meeting(self, meeting: Meeting) -> None:
        if not (meeting.recurrence_rule or "").strip():
            raise ValueError("Нужно указать периодичность встречи")
        if not meeting.start_datetime:
            raise ValueError("Нужно указать дедлайн планирования")

    def compute_status(self, meeting: Meeting) -> str:
        manual_status = (meeting.status or "").strip() or MEETING_STATUS_ACTIVE
        if manual_status in (MEETING_STATUS_DONE, MEETING_STATUS_CANCELLED):
            return manual_status
        if meeting.start_datetime and meeting.start_datetime.date() < date.today():
            return MEETING_STATUS_OVERDUE
        if manual_status == MEETING_STATUS_PAUSED:
            return MEETING_STATUS_PAUSED
        return MEETING_STATUS_ACTIVE

    def get_days_left(self, meeting: Meeting) -> int | None:
        if not meeting.start_datetime:
            return None
        return (meeting.start_datetime.date() - date.today()).days
