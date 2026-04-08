from __future__ import annotations

import os
from datetime import datetime, timedelta, date
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import from_excel

from database import get_session
from models.person import Person
from models.interaction import Interaction
from models.reference import Circle, Responsible
from models.app_setting import AppSetting


EXPORT_DIR = "exports"


IMPORT_HEADERS = [
    "ФИО",
    "Должность",
    "Круг",
    "Уровень",
    "Подкатегория",
    "Столбец 1",
    "Отслеживать звонки?",
    "пол",
    "Дата рождения",
    "Категория подарка",
    "Столбец1",
    "День рождения сегодня?",
    "Периодичность звонка",
    "Дата рождения жены",
    "Дата рождения детей",
    "Служил",
    "Проф праздник ",
    "Столбец2",
    "ПРОФ ПРАЗДНИК?",
    "Религия",
    "Хобби",
    "Телефон/Контактное лицо",
    "Аллергия/вкусы/предпочтения в еде",
    "С кем лучше не пересекать",
    "Что уже дарили ",
    "Последняя встреча ",
    "Дата следующей встречи",
    "Цель встречи ",
    "Пора встретиться?",
    "Последний звонок",
    "Дата следующего звонка",
    "Цель звонка",
    "Пора позвонить",
]


HEADER_TO_ATTR = {
    "ФИО": "fio",
    "Должность": "position",
    "Круг": "circle",
    "Уровень": "level",
    "Подкатегория": "subcategory",
    "Столбец 1": "legacy_column_1",
    "Отслеживать звонки?": "track_calls",
    "пол": "gender",
    "Дата рождения": "birthday",
    "Категория подарка": "gift_category",
    "Столбец1": "legacy_column1",
    "День рождения сегодня?": "birthday_today_flag",
    "Периодичность звонка": "call_periodicity_legacy",
    "Дата рождения жены": "spouse_birthday",
    "Дата рождения детей": "children_birthdays",
    "Служил": "served",
    "Проф праздник ": "prof_holiday",
    "Столбец2": "legacy_column2",
    "ПРОФ ПРАЗДНИК?": "prof_holiday_flag",
    "Религия": "religion",
    "Хобби": "hobby",
    "Телефон/Контактное лицо": "phone_contact_person",
    "Аллергия/вкусы/предпочтения в еде": "food_preferences",
    "С кем лучше не пересекать": "avoid_with",
    "Что уже дарили ": "gifts_already_given",
    "Последняя встреча ": "last_meeting_legacy",
    "Дата следующей встречи": "next_meeting_legacy",
    "Цель встречи ": "meeting_purpose_legacy",
    "Пора встретиться?": "need_meeting_flag",
    "Последний звонок": "last_call_legacy",
    "Дата следующего звонка": "next_call_legacy",
    "Цель звонка": "call_purpose_legacy",
    "Пора позвонить": "need_call_flag",
}


def ensure_export_dir() -> None:
    os.makedirs(EXPORT_DIR, exist_ok=True)


def create_people_import_template() -> str:
    ensure_export_dir()

    wb = Workbook()
    ws = wb.active
    ws.title = "Импорт людей"

    ws.append(IMPORT_HEADERS)

    ws.append([
        "Иванов Иван Иванович",
        "начальник отдела",
        "1",
        "регион",
        "основной контакт",
        "Шафоростов",
        "Да",
        "М",
        "04.04.1990",
        "стандарт",
        "",
        "",
        "14",
        "",
        "",
        "Нет",
        "",
        "",
        "",
        "",
        "",
        "+7 900 000-00-00",
        "",
        "",
        "",
        "01.03.2026",
        "",
        "поддержание контакта",
        "Да",
        "05.03.2026",
        "",
        "уточнить позицию",
        "Да",
    ])

    output_path = os.path.join(EXPORT_DIR, "people_import_template.xlsx")
    wb.save(output_path)
    return output_path


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def stringify_cell(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")

    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")

    return str(value).strip()


def normalize_yes_no(value: str) -> str:
    raw = (value or "").strip().lower()
    if raw in {"да", "yes", "y", "true", "1"}:
        return "Да"
    if raw in {"нет", "no", "n", "false", "0"}:
        return "Нет"
    return (value or "").strip()


def is_track_calls_enabled(value: str) -> bool:
    return normalize_yes_no(value) == "Да"


def parse_excel_serial_date(value: Any):
    if isinstance(value, (int, float)):
        try:
            converted = from_excel(value)
            if isinstance(converted, datetime):
                return converted.date()
            if isinstance(converted, date):
                return converted
        except Exception:
            return None
    return None


def parse_date_flexible(value: Any):
    if value in (None, ""):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    excel_serial_date = parse_excel_serial_date(value)
    if excel_serial_date:
        return excel_serial_date

    raw = str(value).strip()
    if not raw:
        return None

    date_formats = [
        "%d.%m.%Y",
        "%d.%m.%y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m.%d.%Y",
        "%m.%d.%y",
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue

    return None


def get_circle_period_days(session, circle_name: str) -> int:
    if not circle_name:
        return 0

    circle = session.query(Circle).filter(Circle.name == circle_name).first()
    if not circle:
        return 0

    return circle.contact_period_days or 0


def get_meeting_equals_call_setting(session) -> bool:
    setting = session.query(AppSetting).filter(AppSetting.key == "meeting_equals_call").first()
    if not setting:
        return True
    return setting.value == "1"


def deactivate_previous_active_contacts(session, person_id: int, interaction_type: str):
    query = session.query(Interaction).filter(
        Interaction.person_id == person_id,
        Interaction.is_active == 1,
    )

    if not get_meeting_equals_call_setting(session):
        query = query.filter(Interaction.interaction_type == interaction_type)

    old_items = query.all()

    for item in old_items:
        item.is_active = 0
        item.completed_at = datetime.now()


def ensure_responsible_exists(session, responsible_name: str) -> None:
    value = (responsible_name or "").strip()
    if not value:
        return

    exists = session.query(Responsible).filter(Responsible.name == value).first()
    if not exists:
        session.add(Responsible(name=value))
        session.flush()


def create_legacy_interaction_if_needed(
    session,
    person: Person,
    interaction_type: str,
    last_date_raw: Any,
    next_date_raw: Any,
    purpose_raw: str,
    result_text: str,
    comment_text: str,
) -> bool:
    interaction_date = parse_date_flexible(last_date_raw)
    if not interaction_date:
        return False

    existing = (
        session.query(Interaction)
        .filter(
            Interaction.person_id == person.id,
            Interaction.interaction_type == interaction_type,
            Interaction.interaction_date == interaction_date,
        )
        .first()
    )
    if existing:
        return False

    next_date = parse_date_flexible(next_date_raw)
    if not next_date:
        period_days = get_circle_period_days(session, person.circle)
        next_date = interaction_date + timedelta(days=period_days) if period_days > 0 else None

    deactivate_previous_active_contacts(session, person.id, interaction_type)

    interaction = Interaction(
        person_id=person.id,
        interaction_type=interaction_type,
        interaction_date=interaction_date,
        next_date=next_date,
        responsible=person.responsible or "",
        purpose=(purpose_raw or "").strip(),
        result=result_text,
        comment=comment_text,
        is_active=1,
        completed_at=None,
    )
    session.add(interaction)
    return True


def resolve_responsible_from_row(row, header_indexes: dict[str, int]) -> str:
    """
    В текущем шаблоне ответственный физически лежит в столбце 'Столбец 1'
    (это бывший column F исходной таблицы).
    Если позже появится явный столбец 'Ответственный', он тоже можно будет поддержать.
    """
    if "Ответственный" in header_indexes:
        return stringify_cell(row[header_indexes["Ответственный"]])

    if "Столбец 1" in header_indexes:
        return stringify_cell(row[header_indexes["Столбец 1"]])

    return ""


def import_people_from_excel(file_path: str) -> dict:
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    headers = [normalize_header(cell.value) for cell in ws[1]]
    header_indexes = {header: idx for idx, header in enumerate(headers) if header}

    required_headers = ["ФИО"]
    for req in required_headers:
        if req not in header_indexes:
            raise ValueError(f"В файле отсутствует обязательный столбец: {req}")

    session = get_session()

    created_count = 0
    updated_count = 0
    skipped_count = 0
    interactions_created = 0
    responsibles_created = 0

    existing_responsibles_before = {
        x.name.strip()
        for x in session.query(Responsible).all()
        if (x.name or "").strip()
    }

    for row in ws.iter_rows(min_row=2, values_only=True):
        fio = stringify_cell(row[header_indexes["ФИО"]]) if "ФИО" in header_indexes else ""
        if not fio:
            skipped_count += 1
            continue

        person = session.query(Person).filter(Person.fio == fio).first()
        is_new = person is None

        if is_new:
            person = Person(fio=fio)
            session.add(person)

        for header, attr_name in HEADER_TO_ATTR.items():
            if header not in header_indexes:
                continue

            value = row[header_indexes[header]]

            if attr_name == "birthday":
                setattr(person, attr_name, parse_date_flexible(value))
            else:
                text_value = stringify_cell(value)

                if attr_name == "track_calls":
                    text_value = normalize_yes_no(text_value)

                setattr(person, attr_name, text_value)

        # Ответственный берём из column F / "Столбец 1"
        responsible_value = resolve_responsible_from_row(row, header_indexes).strip()
        if responsible_value:
            person.responsible = responsible_value
            ensure_responsible_exists(session, responsible_value)

            if responsible_value not in existing_responsibles_before:
                existing_responsibles_before.add(responsible_value)
                responsibles_created += 1

        if person.phone_contact_person and not person.phone:
            person.phone = person.phone_contact_person

        if is_new:
            created_count += 1
        else:
            updated_count += 1

        session.flush()

        if create_legacy_interaction_if_needed(
            session=session,
            person=person,
            interaction_type="Встреча",
            last_date_raw=person.last_meeting_legacy,
            next_date_raw=person.next_meeting_legacy,
            purpose_raw=person.meeting_purpose_legacy,
            result_text="Импортировано из исходной базы (последняя встреча)",
            comment_text="Создано автоматически при импорте людей из Excel",
        ):
            interactions_created += 1

        if is_track_calls_enabled(person.track_calls):
            if create_legacy_interaction_if_needed(
                session=session,
                person=person,
                interaction_type="Звонок",
                last_date_raw=person.last_call_legacy,
                next_date_raw=person.next_call_legacy,
                purpose_raw=person.call_purpose_legacy,
                result_text="Импортировано из исходной базы (последний звонок)",
                comment_text="Создано автоматически при импорте людей из Excel",
            ):
                interactions_created += 1

    session.commit()
    session.close()

    return {
        "created": created_count,
        "updated": updated_count,
        "skipped": skipped_count,
        "interactions_created": interactions_created,
        "responsibles_created": responsibles_created,
    }