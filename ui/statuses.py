STATUS_ALL = "Все"
STATUS_OVERDUE = "Просрочен"
STATUS_TODAY = "Сегодня"
STATUS_7_DAYS = "7 дней"
STATUS_PLANNED = "Запланирован"
STATUS_NO_DATE = "Без даты"

INTERACTION_STATUS_VALUES = [
    STATUS_ALL,
    STATUS_OVERDUE,
    STATUS_TODAY,
    STATUS_7_DAYS,
    STATUS_PLANNED,
    STATUS_NO_DATE,
]

INTERACTION_STATUS_TRANSLATIONS = {
    "ru": {
        STATUS_ALL: "Все",
        STATUS_OVERDUE: "Просрочен",
        STATUS_TODAY: "Сегодня",
        STATUS_7_DAYS: "7 дней",
        STATUS_PLANNED: "Запланирован",
        STATUS_NO_DATE: "Без даты",
    },
    "en": {
        STATUS_ALL: "All",
        STATUS_OVERDUE: "Overdue",
        STATUS_TODAY: "Today",
        STATUS_7_DAYS: "7 days",
        STATUS_PLANNED: "Planned",
        STATUS_NO_DATE: "No date",
    },
}


def normalize_language(language: str | None) -> str:
    value = (language or "ru").strip().lower()
    return "en" if value == "en" else "ru"


def translate_interaction_status(status: str | None, language: str = "ru") -> str:
    if not status:
        return ""
    language = normalize_language(language)
    return INTERACTION_STATUS_TRANSLATIONS.get(language, {}).get(status, status)



def get_interaction_status_values(language: str = "ru") -> list[str]:
    language = normalize_language(language)
    return [translate_interaction_status(status, language) for status in INTERACTION_STATUS_VALUES]
