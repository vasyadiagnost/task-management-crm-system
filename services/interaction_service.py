from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any
import re

from database import get_session
from models.interaction import Interaction
from models.person import Person
from ui.statuses import (
    STATUS_7_DAYS,
    STATUS_ALL,
    STATUS_NO_DATE,
    STATUS_OVERDUE,
    STATUS_PLANNED,
    STATUS_TODAY,
)


_ID_ARTIFACT_RE = re.compile(r"^\s*\[ID\s+\d+\]\s*$", re.IGNORECASE)


@dataclass
class DisplayPerson:
    fio: str
    responsible: str = ""


@dataclass
class InteractionRow:
    interaction: Interaction
    person: Person | DisplayPerson | None
    display_status: str
    effective_responsible: str

    def __getitem__(self, key: str):
        mapping = {
            "interaction": self.interaction,
            "person": self.person,
            "status": self.display_status,
            "responsible": self.effective_responsible,
            "display_status": self.display_status,
            "effective_responsible": self.effective_responsible,
        }
        return mapping[key]


class InteractionService:
    """Сервис логики контактов."""

    def __init__(self, settings_service=None):
        self.settings_service = settings_service

    def list_active_interactions(
        self,
        search: str | None = None,
        status: str | None = None,
        type_: str | None = None,
        responsible: str | None = None,
    ) -> list[InteractionRow]:
        session = get_session()
        try:
            interactions = (
                session.query(Interaction)
                .filter(Interaction.is_active == 1)
                .order_by(Interaction.next_date.asc().nullsfirst(), Interaction.id.desc())
                .all()
            )

            rows: list[InteractionRow] = []
            for interaction in interactions:
                person = self._resolve_person_for_display(session, interaction)
                row = self._build_row(interaction, person)
                if self._matches_filters(row, search, status, type_, responsible):
                    rows.append(row)
            return rows
        finally:
            session.close()

    def list_all_interactions(
        self,
        search: str | None = None,
        status: str | None = None,
        type_: str | None = None,
        responsible: str | None = None,
    ) -> list[InteractionRow]:
        session = get_session()
        try:
            interactions = (
                session.query(Interaction)
                .order_by(Interaction.interaction_date.desc().nullslast(), Interaction.id.desc())
                .all()
            )

            rows: list[InteractionRow] = []
            for interaction in interactions:
                person = self._resolve_person_for_display(session, interaction)
                row = self._build_row(interaction, person)
                if self._matches_filters(row, search, status, type_, responsible):
                    rows.append(row)
            return rows
        finally:
            session.close()

    def get_interaction(self, interaction_id: int) -> Interaction | None:
        session = get_session()
        try:
            return session.query(Interaction).filter(Interaction.id == interaction_id).first()
        finally:
            session.close()

    def create_interaction(self, data: dict[str, Any], meeting_equals_call: bool = True) -> Interaction:
        session = get_session()
        try:
            payload = self._normalize_payload(data)
            person_id = int(payload["person_id"])
            interaction_type = payload.get("interaction_type", "Звонок")

            self._deactivate_previous_active(
                session=session,
                person_id=person_id,
                interaction_type=interaction_type,
                meeting_equals_call=meeting_equals_call,
            )

            interaction = Interaction(**payload)
            session.add(interaction)
            session.commit()
            session.refresh(interaction)
            return interaction
        finally:
            session.close()

    def update_interaction(self, interaction_id: int, data: dict[str, Any]) -> Interaction | None:
        session = get_session()
        try:
            interaction = session.query(Interaction).filter(Interaction.id == interaction_id).first()
            if not interaction:
                return None

            payload = self._normalize_payload(data)
            for key, value in payload.items():
                if hasattr(interaction, key):
                    setattr(interaction, key, value)

            session.commit()
            session.refresh(interaction)
            return interaction
        finally:
            session.close()

    def delete_interaction(self, interaction_id: int, meeting_equals_call: bool = True) -> bool:
        session = get_session()
        try:
            interaction = session.query(Interaction).filter(Interaction.id == interaction_id).first()
            if not interaction:
                return False

            person_id = interaction.person_id
            interaction_type = interaction.interaction_type
            was_active = interaction.is_active == 1

            session.delete(interaction)
            session.commit()

            if was_active:
                self._restore_previous_active(
                    session=session,
                    person_id=person_id,
                    interaction_type=interaction_type,
                    meeting_equals_call=meeting_equals_call,
                )
                session.commit()

            return True
        finally:
            session.close()

    def compute_status(self, next_date: date | None) -> str:
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

    def get_overdue_interactions(self) -> list[InteractionRow]:
        return self.list_active_interactions(status=STATUS_OVERDUE)

    def get_today_interactions(self) -> list[InteractionRow]:
        return self.list_active_interactions(status=STATUS_TODAY)

    def get_upcoming_interactions(self, days: int = 7) -> list[InteractionRow]:
        session = get_session()
        try:
            interactions = (
                session.query(Interaction)
                .filter(Interaction.is_active == 1)
                .order_by(Interaction.next_date.asc().nullsfirst(), Interaction.id.desc())
                .all()
            )
            rows: list[InteractionRow] = []
            today = date.today()
            limit = today + timedelta(days=days)
            for interaction in interactions:
                if not interaction.next_date:
                    continue
                if today <= interaction.next_date <= limit:
                    person = self._resolve_person_for_display(session, interaction)
                    rows.append(self._build_row(interaction, person))
            return rows
        finally:
            session.close()

    def get_dashboard_counts(self) -> dict[str, int]:
        rows = self.list_active_interactions()
        counts = {
            STATUS_OVERDUE: 0,
            STATUS_TODAY: 0,
            STATUS_7_DAYS: 0,
            STATUS_PLANNED: 0,
            STATUS_NO_DATE: 0,
        }
        for row in rows:
            counts[row.display_status] = counts.get(row.display_status, 0) + 1
        return counts

    @staticmethod
    def is_track_calls_enabled(track_calls_value: str | None) -> bool:
        if track_calls_value is None:
            return True
        normalized = str(track_calls_value).strip().lower()
        return normalized not in {"нет", "no", "false", "0"}

    def _matches_filters(
        self,
        row: InteractionRow,
        search: str | None,
        status: str | None,
        type_: str | None,
        responsible: str | None,
    ) -> bool:
        if search:
            fio = (getattr(row.person, "fio", "") if row.person else "") or ""
            if search.strip().lower() not in fio.lower():
                return False

        if status and status not in {"", STATUS_ALL} and row.display_status != status:
            return False

        if type_ and type_ not in {"", "Все"} and row.interaction.interaction_type != type_:
            return False

        if responsible and responsible not in {"", "Все"}:
            if row.effective_responsible != responsible:
                return False

        return True

    def _build_row(self, interaction: Interaction, person: Person | DisplayPerson | None) -> InteractionRow:
        effective_responsible = (interaction.responsible or "").strip() or ((getattr(person, "responsible", "") or "").strip() if person else "")
        return InteractionRow(
            interaction=interaction,
            person=person,
            display_status=self.compute_status(interaction.next_date),
            effective_responsible=effective_responsible,
        )

    def _resolve_person_for_display(self, session, interaction: Interaction) -> Person | DisplayPerson | None:
        person = session.query(Person).filter(Person.id == interaction.person_id).first()
        if person and self._is_valid_person_name(person.fio):
            return person

        # Legacy-safe display object: не показываем [ID ...] в интерфейсе.
        fallback_name = "Неизвестный контакт"
        fallback_responsible = ""
        if person:
            fallback_responsible = (person.responsible or "").strip()
            if self._is_valid_person_name(person.fio):
                fallback_name = person.fio.strip()
        return DisplayPerson(fio=fallback_name, responsible=fallback_responsible)

    @staticmethod
    def _is_valid_person_name(value: str | None) -> bool:
        text = (value or "").strip()
        if not text:
            return False
        if _ID_ARTIFACT_RE.match(text):
            return False
        return True

    def _deactivate_previous_active(
        self,
        session,
        person_id: int,
        interaction_type: str,
        meeting_equals_call: bool,
    ) -> None:
        query = session.query(Interaction).filter(
            Interaction.person_id == person_id,
            Interaction.is_active == 1,
        )

        if not meeting_equals_call:
            query = query.filter(Interaction.interaction_type == interaction_type)

        for item in query.all():
            item.is_active = 0

    def _restore_previous_active(
        self,
        session,
        person_id: int,
        interaction_type: str,
        meeting_equals_call: bool,
    ) -> None:
        query = session.query(Interaction).filter(Interaction.person_id == person_id)

        if not meeting_equals_call:
            query = query.filter(Interaction.interaction_type == interaction_type)

        previous = (
            query.order_by(Interaction.interaction_date.desc().nullslast(), Interaction.id.desc())
            .first()
        )
        if previous:
            previous.is_active = 1

    def _normalize_payload(self, data: dict[str, Any]) -> dict[str, Any]:
        payload = dict(data)

        if "person_id" in payload:
            payload["person_id"] = int(payload["person_id"])

        for date_key in ("interaction_date", "next_date"):
            value = payload.get(date_key)
            if isinstance(value, str):
                value = value.strip()
                payload[date_key] = datetime.strptime(value, "%d.%m.%Y").date() if value else None

        payload.setdefault("purpose", "")
        payload.setdefault("result", "")
        payload.setdefault("comment", "")
        payload.setdefault("responsible", "")
        payload.setdefault("interaction_type", "Звонок")
        payload.setdefault("is_active", 1)
        return payload


# local import to avoid unused at module top after refactor
from datetime import datetime
