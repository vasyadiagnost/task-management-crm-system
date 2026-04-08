import traceback
import customtkinter as ctk

from ui.persons_tab import PersonsTab
from ui.interactions_tab import InteractionsTab
from ui.service_tab import ServiceTab
from ui.dashboard_tab import DashboardTab
from ui.reminder_tab import ReminderTab
from ui.tasks_tab import TasksTab
from ui.meetings_tab import MeetingsTab
from ui.registry_tasks_tab import RegistryTasksTab
from ui.database_editor_tab import DatabaseEditorTab
from ui.style_kit import apply_global_style, home_button

from services.person_service import PersonService
from services.interaction_service import InteractionService
from services.settings_service import SettingsService
from services.reminder_service import ReminderService
from services.task_service import TaskService
from services.meeting_service import MeetingService
from services.registry_task_service import RegistryTaskService
from services.database_editor_service import DatabaseEditorService


TRANSLATIONS = {
    "ru": {
        "title": "CRM",
        "home": "Главная",
        "persons": "Люди",
        "interactions": "Контакты",
        "tasks": "Задачи",
        "registry_tasks": "Реестр поручений",
        "reminders": "Напоминания",
        "meetings": "Встречи",
        "db_editor": "Редактор БД",
        "service": "Сервис",
        "home_button": "🏠 На главную",
        "language": "Язык",
        "language_ru": "RU",
        "language_en": "EN",
        "tab_error": "Вкладка «{title}» временно не загрузилась",
        "reminder_service_error": "Не удалось создать ReminderService",
        "dashboard_error": "Не удалось создать DashboardTab: несовместимая сигнатура конструктора",
        "persons_error": "Не удалось создать PersonsTab: несовместимая сигнатура конструктора",
        "interactions_error": "Не удалось создать InteractionsTab: несовместимая сигнатура конструктора",
        "tasks_error": "Не удалось создать TasksTab: несовместимая сигнатура конструктора",
        "registry_tasks_error": "Не удалось создать RegistryTasksTab: несовместимая сигнатура конструктора",
        "reminder_error": "Не удалось создать ReminderTab: несовместимая сигнатура конструктора",
        "db_editor_error": "Не удалось создать DatabaseEditorTab: несовместимая сигнатура конструктора",
        "service_error": "Не удалось создать ServiceTab: несовместимая сигнатура конструктора",
    },
    "en": {
        "title": "CRM",
        "home": "Home",
        "persons": "People",
        "interactions": "Interactions",
        "tasks": "Tasks",
        "registry_tasks": "Task Registry",
        "reminders": "Reminders",
        "meetings": "Meetings",
        "db_editor": "DB Editor",
        "service": "Service",
        "home_button": "🏠 Home",
        "language": "Language",
        "language_ru": "RU",
        "language_en": "EN",
        "tab_error": "The “{title}” tab failed to load temporarily",
        "reminder_service_error": "Failed to create ReminderService",
        "dashboard_error": "Failed to create DashboardTab: incompatible constructor signature",
        "persons_error": "Failed to create PersonsTab: incompatible constructor signature",
        "interactions_error": "Failed to create InteractionsTab: incompatible constructor signature",
        "tasks_error": "Failed to create TasksTab: incompatible constructor signature",
        "registry_tasks_error": "Failed to create RegistryTasksTab: incompatible constructor signature",
        "reminder_error": "Failed to create ReminderTab: incompatible constructor signature",
        "db_editor_error": "Failed to create DatabaseEditorTab: incompatible constructor signature",
        "service_error": "Failed to create ServiceTab: incompatible constructor signature",
    },
}

TAB_DEFINITIONS = [
    ("home", "dashboard_parent", "dashboard_tab", "_create_dashboard_tab", False),
    ("persons", "persons_parent", "persons_tab", "_create_persons_tab", True),
    ("interactions", "interactions_parent", "interactions_tab", "_create_interactions_tab", True),
    ("tasks", "tasks_parent", "tasks_tab", "_create_tasks_tab", True),
    ("registry_tasks", "registry_tasks_parent", "registry_tasks_tab", "_create_registry_tasks_tab", True),
    ("reminders", "reminder_parent", "reminder_tab", "_create_reminder_tab", True),
    ("meetings", "meetings_parent", "meetings_tab", "_create_meetings_tab", True),
    ("db_editor", "db_editor_parent", "database_editor_tab", "_create_database_editor_tab", True),
    ("service", "service_parent", "service_tab", "_create_service_tab", True),
]


class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.settings_service = SettingsService()
        self.current_language = self._normalize_language(self.settings_service.get("ui_language", "ru"))

        self.title(self.tr("title"))
        self.geometry("1500x900")
        self.minsize(1320, 780)

        apply_global_style(self)

        self.person_service = PersonService()
        self.interaction_service = InteractionService(settings_service=self.settings_service)
        self.task_service = TaskService()
        self.meeting_service = MeetingService()
        self.registry_task_service = RegistryTaskService()
        self.database_editor_service = DatabaseEditorService()
        self.reminder_service = self._create_reminder_service()

        self.tabview = None
        self.language_segment = None
        self._tab_key_to_name = {}
        self._build_ui()

    def tr(self, key: str, **kwargs) -> str:
        text = TRANSLATIONS.get(self.current_language, TRANSLATIONS["ru"]).get(key, key)
        if kwargs:
            return text.format(**kwargs)
        return text

    def _normalize_language(self, value: str | None) -> str:
        return "en" if str(value or "ru").strip().lower() == "en" else "ru"

    def _build_ui(self):
        if self.tabview is not None:
            self.tabview.destroy()

        self.title(self.tr("title"))

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self._tab_key_to_name = {}
        for tab_key, *_rest in TAB_DEFINITIONS:
            tab_name = self.tr(tab_key)
            self._tab_key_to_name[tab_key] = tab_name
            self.tabview.add(tab_name)

        self.dashboard_parent = self._create_tab_shell("home", show_home=False)
        self.persons_parent = self._create_tab_shell("persons", show_home=True)
        self.interactions_parent = self._create_tab_shell("interactions", show_home=True)
        self.tasks_parent = self._create_tab_shell("tasks", show_home=True)
        self.registry_tasks_parent = self._create_tab_shell("registry_tasks", show_home=True)
        self.reminder_parent = self._create_tab_shell("reminders", show_home=True)
        self.meetings_parent = self._create_tab_shell("meetings", show_home=True)
        self.db_editor_parent = self._create_tab_shell("db_editor", show_home=True)
        self.service_parent = self._create_tab_shell("service", show_home=True)

        self.dashboard_tab = self._mount_tab(self.dashboard_parent, self._create_dashboard_tab, "home")
        self.persons_tab = self._mount_tab(self.persons_parent, self._create_persons_tab, "persons")
        self.interactions_tab = self._mount_tab(self.interactions_parent, self._create_interactions_tab, "interactions")
        self.tasks_tab = self._mount_tab(self.tasks_parent, self._create_tasks_tab, "tasks")
        self.registry_tasks_tab = self._mount_tab(self.registry_tasks_parent, self._create_registry_tasks_tab, "registry_tasks")
        self.reminder_tab = self._mount_tab(self.reminder_parent, self._create_reminder_tab, "reminders")
        self.meetings_tab = self._mount_tab(self.meetings_parent, self._create_meetings_tab, "meetings")
        self.database_editor_tab = self._mount_tab(self.db_editor_parent, self._create_database_editor_tab, "db_editor")
        self.service_tab = self._mount_tab(self.service_parent, self._create_service_tab, "service")

        self.tabview.set(self.tr("home"))

    def _create_tab_shell(self, tab_key: str, show_home: bool) -> ctk.CTkFrame:
        tab = self.tabview.tab(self._tab_key_to_name[tab_key])
        shell = ctk.CTkFrame(tab, fg_color="transparent")
        shell.pack(fill="both", expand=True)

        if show_home:
            nav_frame = ctk.CTkFrame(shell, fg_color="transparent")
            nav_frame.pack(fill="x", padx=8, pady=(6, 0))

            lang_frame = ctk.CTkFrame(nav_frame, fg_color="transparent")
            lang_frame.pack(side="left", anchor="nw")

            ctk.CTkLabel(
                lang_frame,
                text=self.tr("language"),
                font=ctk.CTkFont(size=12, weight="normal"),
            ).pack(side="left", padx=(0, 8))

            self.language_segment = ctk.CTkSegmentedButton(
                lang_frame,
                values=["RU", "EN"],
                width=110,
                command=self._on_language_change,
            )
            self.language_segment.pack(side="left")
            self.language_segment.set(self.current_language.upper())

            home_btn = home_button(nav_frame, lambda: self.tabview.set(self.tr("home")))
            home_btn.configure(text=self.tr("home_button"))
            home_btn.pack(side="right", anchor="ne")

        content = ctk.CTkFrame(shell, fg_color="transparent")
        content.pack(fill="both", expand=True)
        return content

    def _on_language_change(self, value: str):
        new_language = self._normalize_language(value)
        if new_language == self.current_language:
            return

        self.current_language = new_language
        self.settings_service.set("ui_language", self.current_language)
        self._build_ui()

    def _mount_tab(self, parent, creator, title_key: str):
        try:
            widget = creator(parent)
            widget.pack(fill="both", expand=True)
            return widget
        except Exception as exc:
            self._show_tab_error(parent, title_key, exc)
            return None

    def _show_tab_error(self, parent, title_key: str, exc: Exception):
        box = ctk.CTkFrame(parent, corner_radius=14)
        box.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            box,
            text=self.tr("tab_error", title=self.tr(title_key)),
            font=ctk.CTkFont(size=20, weight="normal"),
        ).pack(anchor="w", padx=16, pady=(16, 8))

        ctk.CTkLabel(
            box,
            text=str(exc),
            justify="left",
            wraplength=1100,
            font=ctk.CTkFont(size=13, weight="normal"),
        ).pack(anchor="w", padx=16, pady=(0, 10))

        details = ctk.CTkTextbox(box, height=340)
        details.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        details.insert("1.0", traceback.format_exc())
        details.configure(state="disabled")

    def _create_reminder_service(self):
        for factory in (
            lambda: ReminderService(person_service=self.person_service, interaction_service=self.interaction_service),
            lambda: ReminderService(interaction_service=self.interaction_service),
            lambda: ReminderService(),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("reminder_service_error"))

    def _create_dashboard_tab(self, parent):
        for factory in (
            lambda: DashboardTab(parent, person_service=self.person_service, interaction_service=self.interaction_service, task_service=self.task_service, meeting_service=self.meeting_service),
            lambda: DashboardTab(parent, self.interaction_service, self.person_service, self.task_service, self.meeting_service),
            lambda: DashboardTab(parent, self.interaction_service, self.person_service),
            lambda: DashboardTab(parent, interaction_service=self.interaction_service),
            lambda: DashboardTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("dashboard_error"))

    def _create_persons_tab(self, parent):
        for factory in (
            lambda: PersonsTab(parent, person_service=self.person_service),
            lambda: PersonsTab(parent, self.person_service),
            lambda: PersonsTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("persons_error"))

    def _create_interactions_tab(self, parent):
        for factory in (
            lambda: InteractionsTab(parent, interaction_service=self.interaction_service, settings_service=self.settings_service),
            lambda: InteractionsTab(parent, self.interaction_service, self.settings_service),
            lambda: InteractionsTab(parent, interaction_service=self.interaction_service),
            lambda: InteractionsTab(parent, self.interaction_service),
            lambda: InteractionsTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("interactions_error"))

    def _create_tasks_tab(self, parent):
        for factory in (
            lambda: TasksTab(parent, task_service=self.task_service),
            lambda: TasksTab(parent, self.task_service),
            lambda: TasksTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("tasks_error"))

    def _create_registry_tasks_tab(self, parent):
        for factory in (
            lambda: RegistryTasksTab(parent, registry_task_service=self.registry_task_service),
            lambda: RegistryTasksTab(parent, self.registry_task_service),
            lambda: RegistryTasksTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("registry_tasks_error"))

    def _create_reminder_tab(self, parent):
        for factory in (
            lambda: ReminderTab(parent, reminder_service=self.reminder_service),
            lambda: ReminderTab(parent, self.reminder_service),
            lambda: ReminderTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("reminder_error"))

    def _create_meetings_tab(self, parent):
        return MeetingsTab(parent, self.meeting_service)

    def _create_database_editor_tab(self, parent):
        for factory in (
            lambda: DatabaseEditorTab(parent, editor_service=self.database_editor_service),
            lambda: DatabaseEditorTab(parent, database_editor_service=self.database_editor_service),
            lambda: DatabaseEditorTab(parent, self.database_editor_service),
            lambda: DatabaseEditorTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("db_editor_error"))

    def _create_service_tab(self, parent):
        for factory in (
            lambda: ServiceTab(parent, settings_service=self.settings_service),
            lambda: ServiceTab(parent, self.settings_service),
            lambda: ServiceTab(parent),
        ):
            try:
                return factory()
            except TypeError:
                continue
        raise TypeError(self.tr("service_error"))
