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


class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CRM")
        self.geometry("1500x900")
        self.minsize(1320, 780)

        apply_global_style(self)

        self.settings_service = SettingsService()
        self.person_service = PersonService()
        self.interaction_service = InteractionService(settings_service=self.settings_service)
        self.task_service = TaskService()
        self.meeting_service = MeetingService()
        self.registry_task_service = RegistryTaskService()
        self.database_editor_service = DatabaseEditorService()
        self.reminder_service = self._create_reminder_service()

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self._tab_names = [
            "Главная",
            "Люди",
            "Контакты",
            "Задачи",
            "Реестр поручений",
            "Напоминания",
            "Встречи",
            "Редактор БД",
            "Сервис",
        ]
        for name in self._tab_names:
            self.tabview.add(name)

        self.dashboard_parent = self._create_tab_shell("Главная", show_home=False)
        self.persons_parent = self._create_tab_shell("Люди", show_home=True)
        self.interactions_parent = self._create_tab_shell("Контакты", show_home=True)
        self.tasks_parent = self._create_tab_shell("Задачи", show_home=True)
        self.registry_tasks_parent = self._create_tab_shell("Реестр поручений", show_home=True)
        self.reminder_parent = self._create_tab_shell("Напоминания", show_home=True)
        self.meetings_parent = self._create_tab_shell("Встречи", show_home=True)
        self.db_editor_parent = self._create_tab_shell("Редактор БД", show_home=True)
        self.service_parent = self._create_tab_shell("Сервис", show_home=True)

        self.dashboard_tab = self._mount_tab(self.dashboard_parent, self._create_dashboard_tab, "Главная")
        self.persons_tab = self._mount_tab(self.persons_parent, self._create_persons_tab, "Люди")
        self.interactions_tab = self._mount_tab(self.interactions_parent, self._create_interactions_tab, "Контакты")
        self.tasks_tab = self._mount_tab(self.tasks_parent, self._create_tasks_tab, "Задачи")
        self.registry_tasks_tab = self._mount_tab(self.registry_tasks_parent, self._create_registry_tasks_tab, "Реестр поручений")
        self.reminder_tab = self._mount_tab(self.reminder_parent, self._create_reminder_tab, "Напоминания")
        self.meetings_tab = self._mount_tab(self.meetings_parent, self._create_meetings_tab, "Встречи")
        self.database_editor_tab = self._mount_tab(self.db_editor_parent, self._create_database_editor_tab, "Редактор БД")
        self.service_tab = self._mount_tab(self.service_parent, self._create_service_tab, "Сервис")

    def _create_tab_shell(self, tab_name: str, show_home: bool) -> ctk.CTkFrame:
        tab = self.tabview.tab(tab_name)
        shell = ctk.CTkFrame(tab, fg_color="transparent")
        shell.pack(fill="both", expand=True)

        if show_home:
            nav_frame = ctk.CTkFrame(shell, fg_color="transparent")
            nav_frame.pack(fill="x", padx=8, pady=(6, 0))
            home_button(nav_frame, lambda: self.tabview.set("Главная")).pack(anchor="ne")

        content = ctk.CTkFrame(shell, fg_color="transparent")
        content.pack(fill="both", expand=True)
        return content

    def _mount_tab(self, parent, creator, title: str):
        try:
            widget = creator(parent)
            widget.pack(fill="both", expand=True)
            return widget
        except Exception as exc:
            self._show_tab_error(parent, title, exc)
            return None

    def _show_tab_error(self, parent, title: str, exc: Exception):
        box = ctk.CTkFrame(parent, corner_radius=14)
        box.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            box,
            text=f"Вкладка «{title}» временно не загрузилась",
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
        raise TypeError("Не удалось создать ReminderService")

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
        raise TypeError("Не удалось создать DashboardTab: несовместимая сигнатура конструктора")

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
        raise TypeError("Не удалось создать PersonsTab: несовместимая сигнатура конструктора")

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
        raise TypeError("Не удалось создать InteractionsTab: несовместимая сигнатура конструктора")

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
        raise TypeError("Не удалось создать TasksTab: несовместимая сигнатура конструктора")

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
        raise TypeError("Не удалось создать RegistryTasksTab: несовместимая сигнатура конструктора")

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
        raise TypeError("Не удалось создать ReminderTab: несовместимая сигнатура конструктора")

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
        raise TypeError("Не удалось создать DatabaseEditorTab: несовместимая сигнатура конструктора")

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
        raise TypeError("Не удалось создать ServiceTab: несовместимая сигнатура конструктора")
