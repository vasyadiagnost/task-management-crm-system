from __future__ import annotations

from database import get_session
from models.app_setting import AppSetting


class SettingsService:
    """Простой сервис настроек приложения."""

    def get(self, key: str, default: str | None = None) -> str | None:
        session = get_session()
        try:
            setting = session.query(AppSetting).filter(AppSetting.key == key).first()
            if not setting:
                return default
            return setting.value
        finally:
            session.close()

    def set(self, key: str, value: str) -> None:
        session = get_session()
        try:
            setting = session.query(AppSetting).filter(AppSetting.key == key).first()
            if setting:
                setting.value = value
            else:
                session.add(AppSetting(key=key, value=value))
            session.commit()
        finally:
            session.close()

    def get_bool(self, key: str, default: bool = False) -> bool:
        value = self.get(key)
        if value is None:
            return default
        return str(value).strip().lower() in {"1", "true", "yes", "да"}

    def set_bool(self, key: str, value: bool) -> None:
        self.set(key, "1" if value else "0")
