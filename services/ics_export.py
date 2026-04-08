from __future__ import annotations

import os
from datetime import datetime, timedelta
from typing import Iterable


EXPORT_DIR = "exports"


def ensure_export_dir() -> None:
    os.makedirs(EXPORT_DIR, exist_ok=True)


def escape_ics_text(value: str) -> str:
    if not value:
        return ""

    return (
        value.replace("\\", "\\\\")
        .replace(";", r"\;")
        .replace(",", r"\,")
        .replace("\n", r"\n")
    )


def format_ics_datetime(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%S")


def build_uid(prefix: str, item_id: int, timestamp: str) -> str:
    return f"{prefix}-{item_id}-{timestamp}@crm-system"


def export_interactions_to_ics(interactions_with_persons: Iterable[tuple], include_status: bool = True) -> str:
    """
    interactions_with_persons: iterable of tuples (interaction, person, status)
    Экспортирует активные контакты с next_date в ICS.
    include_status=False убирает строку "Статус: ..." из DESCRIPTION.
    """

    ensure_export_dir()

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_path = os.path.join(EXPORT_DIR, f"contacts_export_{timestamp}.ics")

    lines: list[str] = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//CRM System//Contacts Export//RU",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]

    now_utc = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    for interaction, person, status in interactions_with_persons:
        if not interaction.next_date:
            continue

        start_dt = datetime.combine(interaction.next_date, datetime.min.time()).replace(hour=9, minute=0, second=0)
        end_dt = start_dt + timedelta(minutes=30)

        person_fio = getattr(person, "fio", None) or "Неизвестный контакт"
        summary = f"{interaction.interaction_type}: {person_fio}"
        description_parts = []
        if include_status and status:
            description_parts.append(f"Статус: {status}")
        description_parts.append(f"Ответственный: {interaction.responsible or '-'}")

        if interaction.purpose:
            description_parts.append(f"Цель: {interaction.purpose}")

        if interaction.result:
            description_parts.append(f"Результат: {interaction.result}")

        if interaction.comment:
            description_parts.append(f"Комментарий: {interaction.comment}")

        description = "\n".join(description_parts)

        uid = build_uid("interaction", interaction.id, timestamp)

        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{now_utc}",
            f"DTSTART:{format_ics_datetime(start_dt)}",
            f"DTEND:{format_ics_datetime(end_dt)}",
            f"SUMMARY:{escape_ics_text(summary)}",
            f"DESCRIPTION:{escape_ics_text(description)}",
            "END:VEVENT",
        ])

    lines.append("END:VCALENDAR")

    with open(output_path, "w", encoding="utf-8", newline="\r\n") as f:
        f.write("\r\n".join(lines) + "\r\n")

    return output_path
