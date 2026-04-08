from __future__ import annotations

import sqlite3
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

from database import get_session
from models.person import Person


# -------------------------------------------------------------------------
# Пути
# -------------------------------------------------------------------------
# Ожидаемая структура:
# C:\projects\CRM 1.1\app\
#   ├── crm\
#   └── birthday_reminder\

BIRTHDAY_APP_DIR_RELATIVE_PATH = Path("..") / "birthday_reminder"
BIRTHDAY_DB_RELATIVE_PATH = BIRTHDAY_APP_DIR_RELATIVE_PATH / "data" / "birthday.db"
BIRTHDAY_EXE_RELATIVE_PATH = BIRTHDAY_APP_DIR_RELATIVE_PATH / "birthday_reminder.exe"
BIRTHDAY_MAIN_RELATIVE_PATH = BIRTHDAY_APP_DIR_RELATIVE_PATH / "main.py"


def _get_runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_birthday_app_dir() -> Path:
    return (_get_runtime_dir() / BIRTHDAY_APP_DIR_RELATIVE_PATH).resolve()


def get_birthday_db_path() -> Path:
    return (_get_runtime_dir() / BIRTHDAY_DB_RELATIVE_PATH).resolve()


def get_birthday_exe_path() -> Path:
    return (_get_runtime_dir() / BIRTHDAY_EXE_RELATIVE_PATH).resolve()


def get_birthday_main_path() -> Path:
    return (_get_runtime_dir() / BIRTHDAY_MAIN_RELATIVE_PATH).resolve()


# -------------------------------------------------------------------------
# Нормализация
# -------------------------------------------------------------------------

def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_birth_date(value) -> str:
    """
    Приводим дату рождения к формату:
    - dd.mm.yyyy если значение date/datetime
    - либо возвращаем строку как есть, если она уже в базе/тексте
    """
    if value is None:
        return ""

    if isinstance(value, datetime):
        value = value.date()

    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")

    raw = str(value).strip()
    if not raw:
        return ""

    return raw


# -------------------------------------------------------------------------
# SQLite Birthday Reminder
# -------------------------------------------------------------------------

def ensure_birthday_table_exists(conn: sqlite3.Connection) -> None:
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS persons_birthdays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            crm_person_id INTEGER,
            full_name TEXT NOT NULL,
            birth_date TEXT NOT NULL,
            position TEXT DEFAULT '',
            department TEXT DEFAULT '',
            circle TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            congratulate TEXT DEFAULT 'Да',
            gift_type TEXT DEFAULT '',
            greeting_text TEXT DEFAULT '',
            comment TEXT DEFAULT '',
            source TEXT DEFAULT '',
            source_file TEXT DEFAULT '',
            source_row INTEGER DEFAULT 0,
            created_at TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )
    """)

    cursor.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_persons_birthdays_name_birth
        ON persons_birthdays (full_name, birth_date)
    """)

    cursor.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_persons_birthdays_crm_person_id
        ON persons_birthdays (crm_person_id)
        WHERE crm_person_id IS NOT NULL
    """)

    conn.commit()


# -------------------------------------------------------------------------
# Поиск записей
# -------------------------------------------------------------------------

def find_row_by_crm_person_id(cursor: sqlite3.Cursor, crm_person_id: int | None):
    if crm_person_id is None:
        return None

    cursor.execute("""
        SELECT *
        FROM persons_birthdays
        WHERE crm_person_id = ?
        LIMIT 1
    """, (crm_person_id,))
    return cursor.fetchone()


def find_row_by_name_and_birth(cursor: sqlite3.Cursor, full_name: str, birth_date: str):
    cursor.execute("""
        SELECT *
        FROM persons_birthdays
        WHERE full_name = ? AND birth_date = ?
        LIMIT 1
    """, (full_name, birth_date))
    return cursor.fetchone()


def find_row_by_name_only(cursor: sqlite3.Cursor, full_name: str):
    """
    Фолбэк-сценарий:
    если имя совпадает, а дата в Birthday Reminder была неполной/устаревшей,
    всё равно находим запись для актуализации.
    """
    cursor.execute("""
        SELECT *
        FROM persons_birthdays
        WHERE full_name = ?
        ORDER BY id ASC
        LIMIT 1
    """, (full_name,))
    return cursor.fetchone()


# -------------------------------------------------------------------------
# Подготовка данных из CRM
# -------------------------------------------------------------------------

def build_comment(person: Person) -> str:
    comments = ["Импорт из основной базы CRM"]

    if not normalize_text(getattr(person, "circle", "")):
        comments.append("круг общения не был определен")

    return "; ".join(comments)


def extract_person_payload(person: Person) -> dict:
    """
    В CRM дата рождения хранится в поле `birthday`.
    """
    full_name = normalize_text(getattr(person, "fio", ""))
    birthday_value = getattr(person, "birthday", None)
    birth_date = normalize_birth_date(birthday_value)

    if not full_name or not birth_date:
        return {}

    phone_value = normalize_text(
        getattr(person, "phone_contact_person", "") or getattr(person, "phone", "")
    )

    return {
        "crm_person_id": getattr(person, "id", None),
        "full_name": full_name,
        "birth_date": birth_date,
        "position": normalize_text(getattr(person, "position", "")),
        "department": normalize_text(getattr(person, "department", "")),
        "circle": normalize_text(getattr(person, "circle", "")),
        "phone": phone_value,
        "congratulate": "Да",
        "gift_type": "",
        "greeting_text": "",
        "comment": build_comment(person),
        "source": "crm_sync",
        "source_file": "",
        "source_row": 0,
    }


# -------------------------------------------------------------------------
# INSERT / UPDATE
# -------------------------------------------------------------------------

def insert_payload(cursor: sqlite3.Cursor, payload: dict) -> None:
    now = datetime.now().isoformat(timespec="seconds")

    cursor.execute("""
        INSERT INTO persons_birthdays (
            crm_person_id,
            full_name,
            birth_date,
            position,
            department,
            circle,
            phone,
            congratulate,
            gift_type,
            greeting_text,
            comment,
            source,
            source_file,
            source_row,
            created_at,
            updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        payload["crm_person_id"],
        payload["full_name"],
        payload["birth_date"],
        payload["position"],
        payload["department"],
        payload["circle"],
        payload["phone"],
        payload["congratulate"],
        payload["gift_type"],
        payload["greeting_text"],
        payload["comment"],
        payload["source"],
        payload["source_file"],
        payload["source_row"],
        now,
        now,
    ))


def build_updated_payload(existing_row, crm_payload: dict) -> dict:
    """
    CRM — источник истины для:
    - full_name
    - birth_date
    - position
    - department
    - circle
    - phone
    - crm_person_id
    - comment
    - source

    Локальные поля Birthday Reminder сохраняем:
    - congratulate
    - gift_type
    - greeting_text
    - source_file
    - source_row
    """
    return {
        "crm_person_id": crm_payload["crm_person_id"],
        "full_name": crm_payload["full_name"],
        "birth_date": crm_payload["birth_date"],
        "position": crm_payload["position"],
        "department": crm_payload["department"],
        "circle": crm_payload["circle"],
        "phone": crm_payload["phone"],
        "congratulate": normalize_text(existing_row["congratulate"]) or "Да",
        "gift_type": normalize_text(existing_row["gift_type"]),
        "greeting_text": normalize_text(existing_row["greeting_text"]),
        "comment": crm_payload["comment"],
        "source": "crm_sync",
        "source_file": normalize_text(existing_row["source_file"]),
        "source_row": int(existing_row["source_row"] or 0),
    }


def payload_differs(existing_row, new_payload: dict) -> bool:
    comparable_fields = [
        "crm_person_id",
        "full_name",
        "birth_date",
        "position",
        "department",
        "circle",
        "phone",
        "congratulate",
        "gift_type",
        "greeting_text",
        "comment",
        "source",
        "source_file",
        "source_row",
    ]

    for field in comparable_fields:
        old_value = existing_row[field]
        new_value = new_payload[field]

        if field in {"source_row", "crm_person_id"}:
            old_norm = int(old_value) if old_value not in (None, "") else None
            new_norm = int(new_value) if new_value not in (None, "") else None
        else:
            old_norm = normalize_text(old_value)
            new_norm = normalize_text(new_value)

        if old_norm != new_norm:
            return True

    return False


def update_existing_row(cursor: sqlite3.Cursor, row_id: int, payload: dict) -> None:
    now = datetime.now().isoformat(timespec="seconds")

    cursor.execute("""
        UPDATE persons_birthdays
        SET
            crm_person_id = ?,
            full_name = ?,
            birth_date = ?,
            position = ?,
            department = ?,
            circle = ?,
            phone = ?,
            congratulate = ?,
            gift_type = ?,
            greeting_text = ?,
            comment = ?,
            source = ?,
            source_file = ?,
            source_row = ?,
            updated_at = ?
        WHERE id = ?
    """, (
        payload["crm_person_id"],
        payload["full_name"],
        payload["birth_date"],
        payload["position"],
        payload["department"],
        payload["circle"],
        payload["phone"],
        payload["congratulate"],
        payload["gift_type"],
        payload["greeting_text"],
        payload["comment"],
        payload["source"],
        payload["source_file"],
        payload["source_row"],
        now,
        row_id,
    ))


# -------------------------------------------------------------------------
# Синхронизация
# -------------------------------------------------------------------------

def sync_to_birthday_db() -> dict[str, int | str]:
    db_path = get_birthday_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)

    session = get_session()
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    try:
        ensure_birthday_table_exists(conn)

        persons = session.query(Person).all()

        added_count = 0
        updated_count = 0
        unchanged_count = 0
        skipped_empty_birth_date_count = 0
        skipped_empty_name_count = 0

        for person in persons:
            fio_value = normalize_text(getattr(person, "fio", ""))
            birthday_value = getattr(person, "birthday", None)

            if not fio_value:
                skipped_empty_name_count += 1
                continue

            if not normalize_birth_date(birthday_value):
                skipped_empty_birth_date_count += 1
                continue

            crm_payload = extract_person_payload(person)
            if not crm_payload:
                continue

            existing_row = None

            # 1. Самый надёжный путь — по crm_person_id
            existing_row = find_row_by_crm_person_id(cursor, crm_payload["crm_person_id"])

            # 2. Потом — по ФИО + дата рождения
            if existing_row is None:
                existing_row = find_row_by_name_and_birth(
                    cursor,
                    crm_payload["full_name"],
                    crm_payload["birth_date"],
                )

            # 3. Фолбэк — по ФИО, если дата была когда-то внесена неполно/ошибочно
            if existing_row is None:
                existing_row = find_row_by_name_only(
                    cursor,
                    crm_payload["full_name"],
                )

            if existing_row is None:
                insert_payload(cursor, crm_payload)
                added_count += 1
                continue

            updated_payload = build_updated_payload(existing_row, crm_payload)

            if payload_differs(existing_row, updated_payload):
                update_existing_row(cursor, int(existing_row["id"]), updated_payload)
                updated_count += 1
            else:
                unchanged_count += 1

        conn.commit()

        return {
            "db_path": str(db_path),
            "added_count": added_count,
            "updated_count": updated_count,
            "unchanged_count": unchanged_count,
            "skipped_empty_birth_date_count": skipped_empty_birth_date_count,
            "skipped_empty_name_count": skipped_empty_name_count,
        }

    finally:
        conn.close()
        session.close()


# -------------------------------------------------------------------------
# Запуск Birthday Reminder
# -------------------------------------------------------------------------

def launch_birthday_reminder() -> tuple[bool, str]:
    """
    Приоритет:
    1. birthday_reminder.exe
    2. python main.py
    """
    app_dir = get_birthday_app_dir()
    exe_path = get_birthday_exe_path()
    main_path = get_birthday_main_path()

    if exe_path.exists():
        subprocess.Popen([str(exe_path)], cwd=str(app_dir))
        return True, f"Запущен exe:\n{exe_path}"

    if main_path.exists():
        subprocess.Popen([sys.executable, str(main_path)], cwd=str(app_dir))
        return True, f"Запущен main.py:\n{main_path}"

    return False, (
        "Не найден файл запуска.\n\n"
        f"Проверены пути:\n"
        f"{exe_path}\n"
        f"{main_path}"
    )


def sync_and_launch_birthday_reminder() -> dict[str, object]:
    sync_result = sync_to_birthday_db()
    launched, launch_info = launch_birthday_reminder()

    return {
        **sync_result,
        "launched": launched,
        "launch_info": launch_info,
    }