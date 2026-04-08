from __future__ import annotations

import sys
from pathlib import Path

from sqlalchemy import create_engine, text
from sqlalchemy.orm import declarative_base, sessionmaker


def get_runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_data_dir() -> Path:
    data_dir = get_runtime_dir() / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


def get_db_path() -> Path:
    return get_data_dir() / "app.db"


DATABASE_PATH = get_db_path()
DATABASE_URL = f"sqlite:///{DATABASE_PATH.as_posix()}"

engine = create_engine(DATABASE_URL, echo=False, future=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)
Base = declarative_base()


def table_exists(table_name: str) -> bool:
    with engine.connect() as conn:
        result = conn.execute(
            text("SELECT name FROM sqlite_master WHERE type='table' AND name=:table_name"),
            {"table_name": table_name},
        )
        return result.first() is not None


def get_table_columns(table_name: str) -> list[str]:
    if not table_exists(table_name):
        return []
    with engine.connect() as conn:
        result = conn.execute(text(f"PRAGMA table_info({table_name})"))
        return [row[1] for row in result.fetchall()]


def ensure_column_if_missing(table_name: str, column_name: str, sql_definition: str) -> None:
    columns = get_table_columns(table_name)
    if column_name in columns:
        return
    with engine.connect() as conn:
        conn.execute(text(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {sql_definition}"))
        conn.commit()


def ensure_circle_period_column() -> None:
    ensure_column_if_missing("circles", "contact_period_days", "INTEGER NOT NULL DEFAULT 0")


def ensure_interactions_columns() -> None:
    columns = get_table_columns("interactions")
    if not columns:
        return
    ensure_column_if_missing("interactions", "is_active", "INTEGER NOT NULL DEFAULT 1")
    ensure_column_if_missing("interactions", "completed_at", "DATETIME NULL")


def ensure_persons_extended_columns() -> None:
    columns = get_table_columns("persons")
    if not columns:
        return

    extra_columns = {
        "level": "TEXT NOT NULL DEFAULT ''",
        "subcategory": "TEXT NOT NULL DEFAULT ''",
        "legacy_column_1": "TEXT NOT NULL DEFAULT ''",
        "track_calls": "TEXT NOT NULL DEFAULT ''",
        "gender": "TEXT NOT NULL DEFAULT ''",
        "gift_category": "TEXT NOT NULL DEFAULT ''",
        "legacy_column1": "TEXT NOT NULL DEFAULT ''",
        "birthday_today_flag": "TEXT NOT NULL DEFAULT ''",
        "call_periodicity_legacy": "TEXT NOT NULL DEFAULT ''",
        "spouse_birthday": "TEXT NOT NULL DEFAULT ''",
        "children_birthdays": "TEXT NOT NULL DEFAULT ''",
        "served": "TEXT NOT NULL DEFAULT ''",
        "prof_holiday": "TEXT NOT NULL DEFAULT ''",
        "legacy_column2": "TEXT NOT NULL DEFAULT ''",
        "prof_holiday_flag": "TEXT NOT NULL DEFAULT ''",
        "religion": "TEXT NOT NULL DEFAULT ''",
        "hobby": "TEXT NOT NULL DEFAULT ''",
        "phone_contact_person": "TEXT NOT NULL DEFAULT ''",
        "food_preferences": "TEXT NOT NULL DEFAULT ''",
        "avoid_with": "TEXT NOT NULL DEFAULT ''",
        "gifts_already_given": "TEXT NOT NULL DEFAULT ''",
        "last_meeting_legacy": "TEXT NOT NULL DEFAULT ''",
        "next_meeting_legacy": "TEXT NOT NULL DEFAULT ''",
        "meeting_purpose_legacy": "TEXT NOT NULL DEFAULT ''",
        "need_meeting_flag": "TEXT NOT NULL DEFAULT ''",
        "last_call_legacy": "TEXT NOT NULL DEFAULT ''",
        "next_call_legacy": "TEXT NOT NULL DEFAULT ''",
        "call_purpose_legacy": "TEXT NOT NULL DEFAULT ''",
        "need_call_flag": "TEXT NOT NULL DEFAULT ''",
    }
    for column_name, sql_definition in extra_columns.items():
        ensure_column_if_missing("persons", column_name, sql_definition)


def seed_reference_data_once() -> None:
    from models.reference import Responsible, Circle

    session = SessionLocal()

    responsibles_count = session.query(Responsible).count()
    circles_count = session.query(Circle).count()

    if responsibles_count == 0:
        session.add(Responsible(name="Шафоростов"))

    if circles_count == 0:
        for name, period in [("1", 14), ("2", 21), ("3", 45), ("VIP", 14)]:
            session.add(Circle(name=name, contact_period_days=period))

    session.commit()

    default_periods = {"1": 14, "2": 21, "3": 45, "VIP": 14}
    updated = False
    for circle in session.query(Circle).all():
        if circle.contact_period_days == 0 and circle.name in default_periods:
            circle.contact_period_days = default_periods[circle.name]
            updated = True

    if updated:
        session.commit()

    session.close()


def seed_settings_once() -> None:
    from models.app_setting import AppSetting

    session = SessionLocal()
    exists = session.query(AppSetting).filter(AppSetting.key == "meeting_equals_call").first()
    if not exists:
        session.add(AppSetting(key="meeting_equals_call", value="1"))
    session.commit()
    session.close()


def init_db() -> None:
    from models.person import Person  # noqa: F401
    from models.reference import Responsible, Circle  # noqa: F401
    from models.interaction import Interaction  # noqa: F401
    from models.app_setting import AppSetting  # noqa: F401
    from models.task import Task  # noqa: F401
    from models.meeting import Meeting  # noqa: F401
    from models.registry_task import RegistryTask  # noqa: F401

    Base.metadata.create_all(bind=engine)

    ensure_circle_period_column()
    ensure_interactions_columns()
    ensure_persons_extended_columns()

    seed_reference_data_once()
    seed_settings_once()


def get_session():
    return SessionLocal()
