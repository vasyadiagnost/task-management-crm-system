import customtkinter as ctk
from tkinter import ttk
from tkcalendar import Calendar
from datetime import datetime, date, timedelta

from database import get_session
from models.person import Person
from models.reference import Circle
from models.task import (
    TASK_STATUS_VALUES,
    TASK_STATUS_NEW,
    TASK_STATUS_IN_PROGRESS,
    TASK_STATUS_DONE,
    TASK_STATUS_CANCELLED,
    TASK_STATUS_OVERDUE,
)
from services.settings_service import SettingsService
from services.interaction_service import InteractionService
from ui.statuses import normalize_language


DISPLAY_STATUS_ORDER = {
    TASK_STATUS_OVERDUE: 0,
    TASK_STATUS_IN_PROGRESS: 1,
    TASK_STATUS_NEW: 2,
    TASK_STATUS_DONE: 3,
    TASK_STATUS_CANCELLED: 4,
}

TRANSLATIONS = {
    "ru": {
        "btn_new": "➕ Новая задача",
        "btn_edit": "✏️ Редактировать",
        "btn_duplicate": "📄 Дублировать",
        "btn_delete": "🗑 Удалить",
        "btn_refresh": "🔄 Обновить",
        "search": "Поиск",
        "search_placeholder": "Название / ФИО / ответственный...",
        "status": "Статус",
        "responsible": "Ответственный",
        "all": "Все",
        "btn_my_tasks": "Мои задачи",
        "btn_overdue": "Только просроченные",
        "btn_reset": "Сбросить фильтры",
        "counter_new": "🟦 Новые: {count}",
        "counter_progress": "🟨 В работе: {count}",
        "counter_overdue": "🟥 Просроченные: {count}",
        "counter_done": "🟩 Выполненные: {count}",
        "col_id": "ID",
        "col_title": "Задача",
        "col_person": "Контакт",
        "col_main": "Основной",
        "col_co": "Соисполнители",
        "col_controller": "Контроль",
        "col_due": "Срок",
        "col_status": "Статус",
        "msg_select_responsible": "Сначала выбери ответственного в фильтре, и кнопка покажет его задачи",
        "msg_select_task": "Выбери задачу",
        "msg_task_not_found": "Задача не найдена",
        "title_message": "Сообщение",
        "title_error": "Ошибка",
        "status_overdue": "Просрочена",
        "status_done": "Выполнена",
        "status_new": "Новая",
        "status_in_progress": "В работе",
        "status_cancelled": "Отменена",
        "execution_title": "Отработка: {title}",
        "execution_fallback_title": "задача",
        "meta": "Ответственный: {name}",
        "meta_with_due": "Ответственный: {name} | Срок: {due}",
        "execution_info": "Это безопасный режим отработки.\n«Отметить отработанным» подготавливает минимальный контакт автоматически.\n«Открыть контакт» позволяет вручную задать детали. Фактическая дата последнего контакта показывается прямо в строке.",
        "execution_empty": "В описании задачи не найден список людей для отработки.",
        "btn_open_contact": "Открыть контакт",
        "btn_return": "↩ Вернуть в неотработанные",
        "btn_mark_done": "Отметить отработанным",
        "label_done": "✅ Готово",
        "label_not_done": "⌛ Не отработано",
        "last_contact": "Последний зафиксированный контакт: {date}",
        "last_contact_empty": "Последний зафиксированный контакт: —",
        "purpose_auto": "Отработка просроченного контакта по задаче: {title}",
        "msg_person_not_found": "Не удалось найти человека в базе по точному ФИО:\n{fio}",
        "msg_contact_prepare_failed": "Не удалось подготовить карточку контакта для:\n{fio}",
        "msg_auto_prepare_failed": "Не удалось автоматически подготовить контакт для:\n{fio}",
        "progress": "Отработано: {done} из {total}",
        "btn_save_progress": "Сохранить прогресс",
        "btn_finish_task": "Завершить задачу",
        "msg_save_progress_error": "Ошибка сохранения прогресса:\n{error}",
        "msg_finish_error": "Ошибка завершения задачи:\n{error}",
        "msg_progress_saved": "Прогресс по задаче сохранён",
        "msg_task_finished": "Задача переведена в статус «Выполнена»",
        "contact_title": "Контакт: {fio}",
        "field_person": "Человек",
        "field_contact_type": "Тип контакта",
        "contact_type_call": "Звонок",
        "contact_type_meeting": "Встреча",
        "field_contact_date": "Дата контакта",
        "field_next_date": "Дата следующего контакта",
        "field_purpose": "Цель контакта",
        "field_result": "Результат",
        "field_comment": "Комментарий",
        "btn_pick": "📅 Выбрать",
        "btn_save_card": "Сохранить карточку",
        "btn_cancel": "Отмена",
        "msg_contact_date_invalid": "Дата контакта должна быть в формате дд.мм.гггг",
        "msg_next_date_invalid": "Дата следующего контакта должна быть в формате дд.мм.гггг",
        "task_new_title": "Новая задача",
        "task_edit_title": "Редактирование задачи",
        "field_contact": "Контакт",
        "contact_placeholder": "Необязательно. Начните вводить ФИО...",
        "field_task_title": "Название задачи",
        "field_main_responsible": "Основной ответственный",
        "field_co_exec": "Соисполнители",
        "co_exec_placeholder": "Через запятую",
        "field_controller": "Контроль",
        "field_due": "Срок",
        "field_description": "Описание",
        "btn_save": "Сохранить",
        "date_pick_title": "Выбор даты",
        "msg_due_invalid": "Срок должен быть в формате дд.мм.гггг",
    },
    "en": {
        "btn_new": "➕ New task",
        "btn_edit": "✏️ Edit",
        "btn_duplicate": "📄 Duplicate",
        "btn_delete": "🗑 Delete",
        "btn_refresh": "🔄 Refresh",
        "search": "Search",
        "search_placeholder": "Title / full name / responsible...",
        "status": "Status",
        "responsible": "Responsible",
        "all": "All",
        "btn_my_tasks": "My tasks",
        "btn_overdue": "Overdue only",
        "btn_reset": "Reset filters",
        "counter_new": "🟦 New: {count}",
        "counter_progress": "🟨 In progress: {count}",
        "counter_overdue": "🟥 Overdue: {count}",
        "counter_done": "🟩 Done: {count}",
        "col_id": "ID",
        "col_title": "Task",
        "col_person": "Contact",
        "col_main": "Owner",
        "col_co": "Co-executors",
        "col_controller": "Control",
        "col_due": "Due date",
        "col_status": "Status",
        "msg_select_responsible": "First choose a responsible person in the filter, then the button will show their tasks",
        "msg_select_task": "Select a task",
        "msg_task_not_found": "Task not found",
        "title_message": "Message",
        "title_error": "Error",
        "status_overdue": "Overdue",
        "status_done": "Done",
        "status_new": "New",
        "status_in_progress": "In progress",
        "status_cancelled": "Cancelled",
        "execution_title": "Execution: {title}",
        "execution_fallback_title": "task",
        "meta": "Responsible: {name}",
        "meta_with_due": "Responsible: {name} | Due date: {due}",
        "execution_info": "This is a safe execution mode.\n“Mark as completed” prepares a minimal interaction automatically.\n“Open contact” lets you set the details manually. The actual last interaction date is shown right in the row.",
        "execution_empty": "No people list for execution was found in the task description.",
        "btn_open_contact": "Open contact",
        "btn_return": "↩ Return to pending",
        "btn_mark_done": "Mark as completed",
        "label_done": "✅ Done",
        "label_not_done": "⌛ Pending",
        "last_contact": "Last recorded interaction: {date}",
        "last_contact_empty": "Last recorded interaction: —",
        "purpose_auto": "Overdue contact follow-up for task: {title}",
        "msg_person_not_found": "Could not find a person in the database by exact full name:\n{fio}",
        "msg_contact_prepare_failed": "Could not prepare the contact card for:\n{fio}",
        "msg_auto_prepare_failed": "Could not automatically prepare a contact for:\n{fio}",
        "progress": "Completed: {done} of {total}",
        "btn_save_progress": "Save progress",
        "btn_finish_task": "Finish task",
        "msg_save_progress_error": "Progress save error:\n{error}",
        "msg_finish_error": "Task completion error:\n{error}",
        "msg_progress_saved": "Task progress saved",
        "msg_task_finished": "Task moved to the “Done” status",
        "contact_title": "Contact: {fio}",
        "field_person": "Person",
        "field_contact_type": "Interaction type",
        "contact_type_call": "Call",
        "contact_type_meeting": "Meeting",
        "field_contact_date": "Interaction date",
        "field_next_date": "Next interaction date",
        "field_purpose": "Purpose",
        "field_result": "Result",
        "field_comment": "Comment",
        "btn_pick": "📅 Pick",
        "btn_save_card": "Save card",
        "btn_cancel": "Cancel",
        "msg_contact_date_invalid": "Interaction date must be in dd.mm.yyyy format",
        "msg_next_date_invalid": "Next interaction date must be in dd.mm.yyyy format",
        "task_new_title": "New task",
        "task_edit_title": "Edit task",
        "field_contact": "Contact",
        "contact_placeholder": "Optional. Start typing a full name...",
        "field_task_title": "Task title",
        "field_main_responsible": "Main responsible",
        "field_co_exec": "Co-executors",
        "co_exec_placeholder": "Comma-separated",
        "field_controller": "Control",
        "field_due": "Due date",
        "field_description": "Description",
        "btn_save": "Save",
        "date_pick_title": "Pick a date",
        "msg_due_invalid": "Due date must be in dd.mm.yyyy format",
    },
}

TASK_STATUS_TRANSLATIONS = {
    "ru": {
        TASK_STATUS_NEW: "Новая",
        TASK_STATUS_IN_PROGRESS: "В работе",
        TASK_STATUS_DONE: "Выполнена",
        TASK_STATUS_CANCELLED: "Отменена",
        TASK_STATUS_OVERDUE: "Просрочена",
    },
    "en": {
        TASK_STATUS_NEW: "New",
        TASK_STATUS_IN_PROGRESS: "In progress",
        TASK_STATUS_DONE: "Done",
        TASK_STATUS_CANCELLED: "Cancelled",
        TASK_STATUS_OVERDUE: "Overdue",
    },
}

INTERACTION_TYPE_TO_DB = {
    "Звонок": "Звонок",
    "Встреча": "Встреча",
    "Call": "Звонок",
    "Meeting": "Встреча",
}

INTERACTION_TYPE_DISPLAY = {
    "ru": {"Звонок": "Звонок", "Встреча": "Встреча"},
    "en": {"Звонок": "Call", "Встреча": "Meeting"},
}


class LanguageMixin:
    def get_language(self):
        root = self.winfo_toplevel()
        return normalize_language(getattr(root, "current_language", "ru"))

    def tr(self, key: str, **kwargs):
        text = TRANSLATIONS.get(self.get_language(), TRANSLATIONS["ru"]).get(key, key)
        if kwargs:
            return text.format(**kwargs)
        return text

    def display_status(self, status: str | None):
        return TASK_STATUS_TRANSLATIONS.get(self.get_language(), TASK_STATUS_TRANSLATIONS["ru"]).get(status or "", status or "")

    def all_value(self):
        return self.tr("all")

    def status_filter_values(self):
        return [self.all_value()] + [self.display_status(v) for v in TASK_STATUS_VALUES]

    def status_from_display(self, value: str | None):
        value = (value or "").strip()
        if not value or value == self.all_value():
            return "Все"
        for raw_status, display in TASK_STATUS_TRANSLATIONS.get(self.get_language(), TASK_STATUS_TRANSLATIONS["ru"]).items():
            if value == display:
                return raw_status
        return value

    def interaction_type_values(self):
        display_map = INTERACTION_TYPE_DISPLAY.get(self.get_language(), INTERACTION_TYPE_DISPLAY["ru"])
        return [display_map["Звонок"], display_map["Встреча"]]

    def display_interaction_type(self, value: str | None):
        return INTERACTION_TYPE_DISPLAY.get(self.get_language(), INTERACTION_TYPE_DISPLAY["ru"]).get(value or "", value or "")

    def interaction_type_to_db(self, value: str | None):
        return INTERACTION_TYPE_TO_DB.get((value or "").strip(), (value or "").strip())


class TasksTab(ctk.CTkFrame, LanguageMixin):
    def __init__(self, parent, task_service):
        super().__init__(parent)
        self.task_service = task_service
        self.selected_task_id = None

        self.settings_service = SettingsService()
        self.interaction_service = InteractionService(settings_service=self.settings_service)

        self._build_ui()
        self.refresh()

    def _build_ui(self):
        top_frame = ctk.CTkFrame(self, corner_radius=14)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkButton(top_frame, text=self.tr("btn_new"), command=self.open_add_window, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text=self.tr("btn_edit"), command=self.open_edit_window, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text=self.tr("btn_duplicate"), command=self.duplicate_selected, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text=self.tr("btn_delete"), command=self.delete_selected, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text=self.tr("btn_refresh"), command=self.refresh, width=140, height=34, corner_radius=10).pack(side="right", padx=5, pady=8)

        filter_frame = ctk.CTkFrame(self, corner_radius=14)
        filter_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(filter_frame, text=self.tr("search")).pack(side="left", padx=(10, 5), pady=10)
        self.search_entry = ctk.CTkEntry(filter_frame, width=260, placeholder_text=self.tr("search_placeholder"))
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

        ctk.CTkLabel(filter_frame, text=self.tr("status")).pack(side="left", padx=(15, 5), pady=10)
        self.status_filter = ctk.CTkComboBox(filter_frame, values=self.status_filter_values(), width=160, command=lambda _: self.refresh())
        self.status_filter.pack(side="left", padx=5, pady=10)
        self.status_filter.set(self.all_value())

        ctk.CTkLabel(filter_frame, text=self.tr("responsible")).pack(side="left", padx=(15, 5), pady=10)
        self.responsible_filter = ctk.CTkComboBox(
            filter_frame,
            values=self.get_responsible_values_for_ui(),
            width=180,
            command=lambda _: self.refresh(),
        )
        self.responsible_filter.pack(side="left", padx=5, pady=10)
        self.responsible_filter.set(self.all_value())

        ctk.CTkButton(filter_frame, text=self.tr("btn_my_tasks"), width=120, height=34, corner_radius=10, command=self.apply_my_tasks_filter).pack(side="right", padx=(5, 10), pady=10)
        ctk.CTkButton(filter_frame, text=self.tr("btn_overdue"), width=170, height=34, corner_radius=10, command=self.apply_overdue_filter).pack(side="right", padx=5, pady=10)
        ctk.CTkButton(filter_frame, text=self.tr("btn_reset"), width=150, height=34, corner_radius=10, command=self.reset_filters).pack(side="right", padx=5, pady=10)

        self.dashboard_frame = ctk.CTkFrame(self, corner_radius=14)
        self.dashboard_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.new_label = ctk.CTkLabel(self.dashboard_frame, text=self.tr("counter_new", count=0))
        self.new_label.pack(side="left", padx=15, pady=10)
        self.progress_label = ctk.CTkLabel(self.dashboard_frame, text=self.tr("counter_progress", count=0))
        self.progress_label.pack(side="left", padx=15, pady=10)
        self.overdue_label = ctk.CTkLabel(self.dashboard_frame, text=self.tr("counter_overdue", count=0))
        self.overdue_label.pack(side="left", padx=15, pady=10)
        self.done_label = ctk.CTkLabel(self.dashboard_frame, text=self.tr("counter_done", count=0))
        self.done_label.pack(side="left", padx=15, pady=10)

        table_frame = ctk.CTkFrame(self, corner_radius=14)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("id", "title", "person", "main_responsible", "co_executors", "controller", "due_date", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        headings = {
            "id": self.tr("col_id"),
            "title": self.tr("col_title"),
            "person": self.tr("col_person"),
            "main_responsible": self.tr("col_main"),
            "co_executors": self.tr("col_co"),
            "controller": self.tr("col_controller"),
            "due_date": self.tr("col_due"),
            "status": self.tr("col_status"),
        }
        for key, text in headings.items():
            self.tree.heading(key, text=text)

        self.tree.column("id", width=55, anchor="center")
        self.tree.column("title", width=260)
        self.tree.column("person", width=220)
        self.tree.column("main_responsible", width=140)
        self.tree.column("co_executors", width=180)
        self.tree.column("controller", width=140)
        self.tree.column("due_date", width=100, anchor="center")
        self.tree.column("status", width=120, anchor="center")

        self.tree.tag_configure("overdue", background="#ffdddd")
        self.tree.tag_configure("in_progress", background="#fff4cc")
        self.tree.tag_configure("done", background="#ddffdd")

        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 8), pady=8)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_double_click)

    def get_responsible_values_for_ui(self):
        raw_values = self.task_service.get_responsible_values()
        values = [v for v in raw_values if v != "Все"]
        return [self.all_value()] + values

    def responsible_filter_value_to_db(self):
        value = (self.responsible_filter.get() or "").strip()
        return "Все" if value == self.all_value() else value

    def reset_filters(self):
        self.search_entry.delete(0, "end")
        self.status_filter.set(self.all_value())
        self.responsible_filter.configure(values=self.get_responsible_values_for_ui())
        self.responsible_filter.set(self.all_value())
        self.refresh()

    def apply_my_tasks_filter(self):
        current_responsible = self.responsible_filter.get().strip()
        if not current_responsible or current_responsible == self.all_value():
            self.show_message(self.tr("msg_select_responsible"))
            return
        self.status_filter.set(self.all_value())
        self.refresh()

    def apply_overdue_filter(self):
        self.status_filter.set(self.display_status(TASK_STATUS_OVERDUE))
        self.refresh()

    def _row_sort_key(self, row):
        task = row["task"]
        status = row["status"] or TASK_STATUS_NEW
        due = task.due_date or date.max
        return (
            DISPLAY_STATUS_ORDER.get(status, 99),
            due,
            (task.title or "").lower(),
        )

    def refresh(self):
        responsible_values = self.get_responsible_values_for_ui()
        self.responsible_filter.configure(values=responsible_values)
        if self.responsible_filter.get() not in responsible_values:
            self.responsible_filter.set(self.all_value())

        self.status_filter.configure(values=self.status_filter_values())
        if self.status_filter.get() not in self.status_filter_values():
            self.status_filter.set(self.all_value())

        for row in self.tree.get_children():
            self.tree.delete(row)

        rows = self.task_service.list_tasks(
            search=self.search_entry.get().strip(),
            status=self.status_from_display(self.status_filter.get()),
            responsible=self.responsible_filter_value_to_db(),
        )
        rows = sorted(rows, key=self._row_sort_key)

        counts = self.task_service.get_status_counts()
        self.new_label.configure(text=self.tr("counter_new", count=counts.get(TASK_STATUS_NEW, 0)))
        self.progress_label.configure(text=self.tr("counter_progress", count=counts.get(TASK_STATUS_IN_PROGRESS, 0)))
        self.overdue_label.configure(text=self.tr("counter_overdue", count=counts.get(TASK_STATUS_OVERDUE, 0)))
        self.done_label.configure(text=self.tr("counter_done", count=counts.get(TASK_STATUS_DONE, 0)))

        for row in rows:
            task = row["task"]
            person = row["person"]
            status = row["status"]
            due_date = task.due_date.strftime("%d.%m.%Y") if task.due_date else ""

            display_status = self.display_status(status)
            tag = ""
            if status == TASK_STATUS_OVERDUE:
                tag = "overdue"
            elif status == TASK_STATUS_IN_PROGRESS:
                tag = "in_progress"
            elif status == TASK_STATUS_DONE:
                tag = "done"

            self.tree.insert(
                "",
                "end",
                values=(
                    task.id,
                    task.title or "",
                    person.fio if person else "",
                    task.main_responsible or "",
                    task.co_executors or "",
                    task.controller or "",
                    due_date,
                    display_status,
                ),
                tags=(tag,) if tag else (),
            )

        self.selected_task_id = None

    def on_select(self, event=None):
        selected = self.tree.selection()
        if not selected:
            self.selected_task_id = None
            return

        raw_id = self.tree.item(selected[0])["values"][0]
        try:
            self.selected_task_id = int(raw_id)
        except (TypeError, ValueError):
            self.selected_task_id = None

    def on_double_click(self, event):
        if self.selected_task_id is None:
            return

        task = self.task_service.get_task(self.selected_task_id)
        if not task:
            self.show_message(self.tr("msg_task_not_found"))
            return

        if self.task_service.is_execution_task(task):
            TaskExecutionWindow(self, task)
        else:
            self.open_edit_window()

    def open_add_window(self):
        TaskWindow(self, mode="add")

    def open_edit_window(self):
        if self.selected_task_id is None:
            self.show_message(self.tr("msg_select_task"))
            return
        TaskWindow(self, mode="edit", task_id=self.selected_task_id)

    def duplicate_selected(self):
        if self.selected_task_id is None:
            self.show_message(self.tr("msg_select_task"))
            return

        task = self.task_service.get_task(self.selected_task_id)
        if not task:
            self.show_message(self.tr("msg_task_not_found"))
            return

        payload = {
            "person_id": task.person_id,
            "title": f"{(task.title or '').strip()} (копия)",
            "description": task.description or "",
            "main_responsible": task.main_responsible or "",
            "co_executors": task.co_executors or "",
            "controller": task.controller or "",
            "due_date": task.due_date,
            "status": TASK_STATUS_NEW,
        }
        try:
            self.task_service.create_task(payload)
        except ValueError as e:
            self.show_message(str(e))
            return
        self.refresh()

    def delete_selected(self):
        if self.selected_task_id is None:
            self.show_message(self.tr("msg_select_task"))
            return

        deleted = self.task_service.delete_task(self.selected_task_id)
        if not deleted:
            self.show_message(self.tr("msg_task_not_found"))
            return

        self.refresh()

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("460x180")
        win.title(self.tr("title_message"))
        win.grab_set()
        ctk.CTkLabel(win, text=text, wraplength=420, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy, width=120, height=34, corner_radius=10).pack(pady=10)


class TaskExecutionWindow(ctk.CTkToplevel, LanguageMixin):
    def __init__(self, parent_tab, task):
        super().__init__()
        self.parent_tab = parent_tab
        self.task = task

        state = self.parent_tab.task_service.parse_execution_state(task.description or "")
        self.remaining_people = state["remaining"]
        self.done_people = set(state["done"])
        self.persisted_interaction_ids = dict(state["interaction_ids"])
        self.persisted_interaction_dates = dict(state["interaction_dates"])
        self.staged_payloads = {}
        self.row_widgets = {}

        self.title(self.tr("execution_title", title=task.title or self.tr("execution_fallback_title")))
        self.geometry("980x760")
        self.minsize(900, 680)
        self.resizable(True, True)
        self.grab_set()

        self.build()

    def build(self):
        ctk.CTkLabel(self, text=self.task.title or self.tr("execution_fallback_title"), font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(18, 6))

        meta = self.tr("meta", name=self.task.main_responsible or "-")
        if self.task.due_date:
            meta = self.tr("meta_with_due", name=self.task.main_responsible or "-", due=self.task.due_date.strftime("%d.%m.%Y"))
        ctk.CTkLabel(self, text=meta, anchor="w").pack(fill="x", padx=20, pady=(0, 10))

        info_box = ctk.CTkTextbox(self, height=68)
        info_box.pack(fill="x", padx=20, pady=(0, 12))
        info_box.insert("1.0", self.tr("execution_info"))
        info_box.configure(state="disabled")

        list_host = ctk.CTkFrame(self, corner_radius=12)
        list_host.pack(fill="both", expand=True, padx=20, pady=(0, 14))

        self.list_frame = ctk.CTkScrollableFrame(list_host)
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        everyone = self.parent_tab.task_service.merge_execution_people(self.remaining_people, list(self.done_people))
        if not everyone:
            ctk.CTkLabel(self.list_frame, text=self.tr("execution_empty"), anchor="w").pack(fill="x", padx=10, pady=10)
        else:
            for fio in everyone:
                self._create_person_row(fio, is_done=(fio in self.done_people))

        bottom = ctk.CTkFrame(self, corner_radius=12)
        bottom.pack(fill="x", padx=20, pady=(0, 20))

        self.progress_label = ctk.CTkLabel(bottom, text=self._progress_text(), font=ctk.CTkFont(size=14, weight="bold"))
        self.progress_label.pack(side="left", padx=16, pady=14)

        self.save_btn = ctk.CTkButton(bottom, text=self.tr("btn_save_progress"), command=self.save_progress, width=170, height=34, corner_radius=10)
        self.save_btn.pack(side="right", padx=(8, 16), pady=10)

        self.finish_btn = ctk.CTkButton(
            bottom,
            text=self.tr("btn_finish_task"),
            command=self.finish_task,
            width=180,
            height=34,
            corner_radius=10,
            state="normal" if self._all_done() else "disabled",
        )
        self.finish_btn.pack(side="right", padx=8, pady=10)

    def _create_person_row(self, fio: str, is_done: bool):
        row = ctk.CTkFrame(self.list_frame, corner_radius=10)
        row.pack(fill="x", padx=6, pady=6)
        row.grid_columnconfigure(0, weight=1)

        name_label = ctk.CTkLabel(row, text=fio, anchor="w", justify="left")
        name_label.grid(row=0, column=0, sticky="ew", padx=(14, 10), pady=(10, 4))

        open_btn = ctk.CTkButton(row, text=self.tr("btn_open_contact"), width=150, height=32, corner_radius=10, command=lambda current_fio=fio: self.open_contact_editor(current_fio))
        open_btn.grid(row=0, column=1, padx=6, pady=8)

        toggle_btn = ctk.CTkButton(row, text="", width=210, height=32, corner_radius=10, command=lambda current_fio=fio: self.toggle_done(current_fio))
        toggle_btn.grid(row=0, column=2, padx=(6, 12), pady=8)

        status_label = ctk.CTkLabel(row, text="", width=220, anchor="w")
        status_label.grid(row=0, column=3, padx=(0, 12), pady=8)

        last_contact_label = ctk.CTkLabel(row, text="", anchor="w", justify="left")
        last_contact_label.grid(row=1, column=0, columnspan=4, sticky="ew", padx=(14, 12), pady=(0, 10))

        self.row_widgets[fio] = {
            "row": row,
            "open_btn": open_btn,
            "toggle_btn": toggle_btn,
            "status_label": status_label,
            "last_contact_label": last_contact_label,
        }
        self._apply_row_state(fio, is_done)

    def _last_contact_text(self, fio: str) -> str:
        payload = self.staged_payloads.get(fio)
        if payload and payload.get("interaction_date"):
            return self.tr("last_contact", date=payload["interaction_date"].strftime("%d.%m.%Y"))

        saved_date = self.persisted_interaction_dates.get(fio)
        if saved_date:
            return self.tr("last_contact", date=saved_date)

        return self.tr("last_contact_empty")

    def _apply_row_state(self, fio: str, is_done: bool):
        widgets = self.row_widgets[fio]
        widgets["last_contact_label"].configure(text=self._last_contact_text(fio))

        if is_done:
            widgets["row"].configure(fg_color="#ddffdd")
            widgets["toggle_btn"].configure(text=self.tr("btn_return"))
            widgets["status_label"].configure(text=self.tr("label_done"))
        else:
            widgets["row"].configure(fg_color=("gray86", "gray20"))
            widgets["toggle_btn"].configure(text=self.tr("btn_mark_done"))
            widgets["status_label"].configure(text=self.tr("label_not_done"))

    def _find_person_record(self, fio: str):
        target = (fio or "").strip().lower()
        if not target:
            return None
        for item in self.parent_tab.task_service.get_person_records():
            if (item.get("fio") or "").strip().lower() == target:
                return item
        return None

    def _default_payload_for_person(self, fio: str):
        person_record = self._find_person_record(fio)
        if not person_record:
            return None

        interaction_date = date.today()
        next_date = self._calculate_next_date(person_record["id"], interaction_date)

        return {
            "person_id": person_record["id"],
            "interaction_type": "Звонок",
            "interaction_date": interaction_date,
            "next_date": next_date,
            "responsible": self.task.main_responsible or "",
            "purpose": self.tr("purpose_auto", title=self.task.title or "").strip(),
            "result": "",
            "comment": "",
        }

    def _calculate_next_date(self, person_id: int, interaction_date: date):
        session = get_session()
        try:
            person = session.get(Person, person_id)
            if not person or not getattr(person, "circle", None):
                return None
            circle = session.query(Circle).filter(Circle.name == person.circle).first()
            if not circle or not circle.contact_period_days:
                return None
            return interaction_date + timedelta(days=int(circle.contact_period_days))
        finally:
            session.close()

    def open_contact_editor(self, fio: str):
        person_record = self._find_person_record(fio)
        if not person_record:
            self.parent_tab.show_message(self.tr("msg_person_not_found", fio=fio))
            return

        initial_payload = self.staged_payloads.get(fio)
        if not initial_payload:
            initial_payload = self._default_payload_for_person(fio)
        if not initial_payload:
            self.parent_tab.show_message(self.tr("msg_contact_prepare_failed", fio=fio))
            return

        ExecutionContactWindow(self, fio, person_record, initial_payload)

    def apply_contact_payload(self, fio: str, payload: dict):
        self.staged_payloads[fio] = payload
        if fio not in self.done_people:
            self.done_people.add(fio)
        self.remaining_people = [item for item in self.remaining_people if item != fio]
        self._apply_row_state(fio, True)
        self.progress_label.configure(text=self._progress_text())
        self.finish_btn.configure(state="normal" if self._all_done() else "disabled")

    def toggle_done(self, fio: str):
        if fio in self.done_people:
            self.done_people.remove(fio)
            if fio not in self.remaining_people:
                self.remaining_people.append(fio)
            self.remaining_people = self.parent_tab.task_service.merge_execution_people(self.remaining_people, [])
            self.staged_payloads.pop(fio, None)
            self.persisted_interaction_dates.pop(fio, None)
            self._apply_row_state(fio, False)
        else:
            if fio not in self.staged_payloads and fio not in self.persisted_interaction_ids:
                payload = self._default_payload_for_person(fio)
                if not payload:
                    self.parent_tab.show_message(self.tr("msg_auto_prepare_failed", fio=fio))
                    return
                self.staged_payloads[fio] = payload
            self.done_people.add(fio)
            self.remaining_people = [item for item in self.remaining_people if item != fio]
            self._apply_row_state(fio, True)

        self.progress_label.configure(text=self._progress_text())
        self.finish_btn.configure(state="normal" if self._all_done() else "disabled")

    def _all_done(self) -> bool:
        total = len(self.parent_tab.task_service.merge_execution_people(self.remaining_people, list(self.done_people)))
        return total > 0 and len(self.remaining_people) == 0

    def _progress_text(self) -> str:
        total = len(self.parent_tab.task_service.merge_execution_people(self.remaining_people, list(self.done_people)))
        done_count = len(self.done_people)
        return self.tr("progress", done=done_count, total=total)

    def _persist_changes(self, mark_done: bool = False):
        meeting_equals_call = self.parent_tab.settings_service.get_bool("meeting_equals_call", default=True)
        interaction_ids = dict(self.persisted_interaction_ids)
        interaction_dates = dict(self.persisted_interaction_dates)

        for fio, interaction_id in list(interaction_ids.items()):
            if fio not in self.done_people:
                self.parent_tab.interaction_service.delete_interaction(
                    interaction_id,
                    meeting_equals_call=meeting_equals_call,
                )
                interaction_ids.pop(fio, None)
                interaction_dates.pop(fio, None)

        for fio in list(self.done_people):
            payload = self.staged_payloads.get(fio)

            if fio in interaction_ids and payload is None:
                continue

            if fio in interaction_ids and payload is not None:
                self.parent_tab.interaction_service.delete_interaction(
                    interaction_ids[fio],
                    meeting_equals_call=meeting_equals_call,
                )
                interaction_ids.pop(fio, None)
                interaction_dates.pop(fio, None)

            if payload is None:
                payload = self._default_payload_for_person(fio)
                if payload is None:
                    raise ValueError(f"Не удалось подготовить данные контакта для: {fio}")

            interaction = self.parent_tab.interaction_service.create_interaction(payload)
            interaction_ids[fio] = interaction.id
            interaction_dates[fio] = payload["interaction_date"].strftime("%d.%m.%Y")

        updated = self.parent_tab.task_service.save_execution_progress(
            self.task.id,
            remaining_people=self.remaining_people,
            done_people=list(self.done_people),
            interaction_ids=interaction_ids,
            interaction_dates=interaction_dates,
            mark_done=mark_done,
        )
        if not updated:
            raise ValueError("Не удалось сохранить прогресс задачи")
        return updated

    def save_progress(self):
        try:
            updated = self._persist_changes(mark_done=False)
        except Exception as e:
            self.parent_tab.show_message(self.tr("msg_save_progress_error", error=e))
            return

        self.task = updated
        self.parent_tab.refresh()
        self.destroy()
        self.parent_tab.show_message(self.tr("msg_progress_saved"))

    def finish_task(self):
        try:
            updated = self._persist_changes(mark_done=True)
        except Exception as e:
            self.parent_tab.show_message(self.tr("msg_finish_error", error=e))
            return

        self.task = updated
        self.parent_tab.refresh()
        self.destroy()
        self.parent_tab.show_message(self.tr("msg_task_finished"))


class ExecutionContactWindow(ctk.CTkToplevel, LanguageMixin):
    def __init__(self, execution_window, fio: str, person_record: dict, payload: dict):
        super().__init__()
        self.execution_window = execution_window
        self.fio = fio
        self.person_record = person_record
        self.payload = payload

        self.title(self.tr("contact_title", fio=fio))
        self.geometry("860x760")
        self.minsize(820, 680)
        self.resizable(True, True)
        self.grab_set()

        self.build()

    def build(self):
        content = ctk.CTkScrollableFrame(self)
        content.pack(fill="both", expand=True, padx=14, pady=14)

        form_width = 700

        ctk.CTkLabel(content, text=self.tr("field_person")).pack(pady=(8, 2))
        self.person_entry = ctk.CTkEntry(content, width=form_width)
        self.person_entry.pack(pady=(0, 6))
        self.person_entry.insert(0, self.person_record["fio"])
        self.person_entry.configure(state="disabled")

        ctk.CTkLabel(content, text=self.tr("field_contact_type")).pack(pady=(8, 2))
        self.type_combo = ctk.CTkComboBox(content, values=self.interaction_type_values(), width=form_width)
        self.type_combo.pack(pady=(0, 6))
        self.type_combo.set(self.display_interaction_type(self.payload.get("interaction_type") or "Звонок"))

        ctk.CTkLabel(content, text=self.tr("responsible")).pack(pady=(8, 2))
        self.responsible_entry = ctk.CTkEntry(content, width=form_width)
        self.responsible_entry.pack(pady=(0, 6))
        self.responsible_entry.insert(0, self.payload.get("responsible") or "")

        ctk.CTkLabel(content, text=self.tr("field_contact_date")).pack(pady=(8, 2))
        contact_frame = ctk.CTkFrame(content, fg_color="transparent")
        contact_frame.pack(pady=(0, 6))
        self.contact_date_entry = ctk.CTkEntry(contact_frame, width=460, placeholder_text="dd.mm.yyyy")
        self.contact_date_entry.pack(side="left", padx=5)
        if self.payload.get("interaction_date"):
            self.contact_date_entry.insert(0, self.payload["interaction_date"].strftime("%d.%m.%Y"))
        self.contact_date_entry.bind("<FocusOut>", lambda event: self.autofill_next_date())
        ctk.CTkButton(contact_frame, text=self.tr("btn_pick"), width=120, height=34, corner_radius=10, command=lambda: self.open_calendar(self.contact_date_entry, self.autofill_next_date)).pack(side="left", padx=5)

        ctk.CTkLabel(content, text=self.tr("field_next_date")).pack(pady=(8, 2))
        next_frame = ctk.CTkFrame(content, fg_color="transparent")
        next_frame.pack(pady=(0, 6))
        self.next_date_entry = ctk.CTkEntry(next_frame, width=460, placeholder_text="dd.mm.yyyy")
        self.next_date_entry.pack(side="left", padx=5)
        if self.payload.get("next_date"):
            self.next_date_entry.insert(0, self.payload["next_date"].strftime("%d.%m.%Y"))
        ctk.CTkButton(next_frame, text=self.tr("btn_pick"), width=120, height=34, corner_radius=10, command=lambda: self.open_calendar(self.next_date_entry)).pack(side="left", padx=5)

        ctk.CTkLabel(content, text=self.tr("field_purpose")).pack(pady=(8, 2))
        self.purpose_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.purpose_text.pack(pady=(0, 6))
        self.purpose_text.insert("1.0", self.payload.get("purpose") or "")

        ctk.CTkLabel(content, text=self.tr("field_result")).pack(pady=(8, 2))
        self.result_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.result_text.pack(pady=(0, 6))
        self.result_text.insert("1.0", self.payload.get("result") or "")

        ctk.CTkLabel(content, text=self.tr("field_comment")).pack(pady=(8, 2))
        self.comment_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.comment_text.pack(pady=(0, 10))
        self.comment_text.insert("1.0", self.payload.get("comment") or "")

        buttons = ctk.CTkFrame(content, fg_color="transparent")
        buttons.pack(fill="x", pady=(8, 12))

        ctk.CTkButton(buttons, text=self.tr("btn_save_card"), width=180, height=34, corner_radius=10, command=self.save_card).pack(side="left", padx=6)
        ctk.CTkButton(buttons, text=self.tr("btn_cancel"), width=120, height=34, corner_radius=10, command=self.destroy).pack(side="left", padx=6)

    def parse_date(self, raw_value: str):
        raw_value = raw_value.strip()
        if not raw_value:
            return None
        try:
            return datetime.strptime(raw_value, "%d.%m.%Y").date()
        except ValueError:
            return "INVALID"

    def open_calendar(self, target_entry, after_apply=None):
        win = ctk.CTkToplevel(self)
        win.geometry("320x360")
        win.title(self.tr("date_pick_title"))
        win.resizable(False, False)
        win.grab_set()

        cal = Calendar(win, date_pattern="dd.mm.yyyy")
        cal.pack(expand=True, fill="both", padx=10, pady=10)

        def apply():
            target_entry.delete(0, "end")
            target_entry.insert(0, cal.get_date())
            win.destroy()
            if after_apply:
                after_apply()

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(pady=(0, 10))
        ctk.CTkButton(btn_frame, text="OK", width=100, height=34, corner_radius=10, command=apply).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=self.tr("btn_cancel"), width=100, height=34, corner_radius=10, command=win.destroy).pack(side="left", padx=5)

    def autofill_next_date(self):
        interaction_date = self.parse_date(self.contact_date_entry.get())
        if interaction_date in (None, "INVALID"):
            return
        next_date = self.execution_window._calculate_next_date(self.person_record["id"], interaction_date)
        if not next_date:
            return
        self.next_date_entry.delete(0, "end")
        self.next_date_entry.insert(0, next_date.strftime("%d.%m.%Y"))

    def save_card(self):
        interaction_date = self.parse_date(self.contact_date_entry.get())
        if interaction_date == "INVALID":
            self.execution_window.parent_tab.show_message(self.tr("msg_contact_date_invalid"))
            return

        next_date = self.parse_date(self.next_date_entry.get())
        if next_date == "INVALID":
            self.execution_window.parent_tab.show_message(self.tr("msg_next_date_invalid"))
            return

        payload = {
            "person_id": self.person_record["id"],
            "interaction_type": self.interaction_type_to_db(self.type_combo.get().strip()),
            "interaction_date": interaction_date or date.today(),
            "next_date": next_date,
            "responsible": self.responsible_entry.get().strip(),
            "purpose": self.purpose_text.get("1.0", "end").strip(),
            "result": self.result_text.get("1.0", "end").strip(),
            "comment": self.comment_text.get("1.0", "end").strip(),
        }
        self.execution_window.apply_contact_payload(self.fio, payload)
        self.destroy()


class TaskWindow(ctk.CTkToplevel, LanguageMixin):
    def __init__(self, parent_tab, mode="add", task_id=None):
        super().__init__()
        self.parent_tab = parent_tab
        self.mode = mode
        self.task_id = task_id
        self.task = None

        if mode == "edit":
            self.task = self.parent_tab.task_service.get_task(task_id)

        self.title(self.tr("task_new_title") if mode == "add" else self.tr("task_edit_title"))
        self.geometry("920x720")
        self.minsize(900, 680)
        self.resizable(True, True)
        self.grab_set()

        self.person_records = self.parent_tab.task_service.get_person_records()
        self.selected_person_id = None
        self.selected_person_fio = ""
        self.displayed_person_matches = []
        self.person_dropdown_visible = False

        self.build()

    def build(self):
        form_width = 740

        self.person_label = ctk.CTkLabel(self, text=self.tr("field_contact"))
        self.person_label.pack(pady=(10, 3))
        self.person_entry = ctk.CTkEntry(self, width=form_width, placeholder_text=self.tr("contact_placeholder"))
        self.person_entry.pack(pady=(0, 4))
        self.person_entry.bind("<KeyRelease>", self.on_person_entry_keyrelease)
        self.person_entry.bind("<FocusIn>", self.on_person_entry_focus_in)
        self.person_entry.bind("<Return>", self.on_person_entry_return)

        self.person_dropdown_host = ctk.CTkFrame(self, fg_color="transparent", width=form_width, height=1)
        self.person_dropdown_host.pack(pady=(0, 4))
        self.person_dropdown_host.pack_propagate(False)

        self.person_dropdown = ctk.CTkScrollableFrame(self.person_dropdown_host, width=form_width, height=90)
        self.person_dropdown.pack_forget()

        self.title_label = ctk.CTkLabel(self, text=self.tr("field_task_title"))
        self.title_label.pack(pady=(4, 2))
        self.title_entry = ctk.CTkEntry(self, width=form_width)
        self.title_entry.pack(pady=(0, 5))

        self.main_responsible_label = ctk.CTkLabel(self, text=self.tr("field_main_responsible"))
        self.main_responsible_label.pack(pady=(4, 2))
        self.main_responsible_combo = ctk.CTkComboBox(self, values=[v for v in self.parent_tab.task_service.get_responsible_values() if v != "Все"] or [""], width=form_width)
        self.main_responsible_combo.pack(pady=(0, 5))

        self.co_exec_label = ctk.CTkLabel(self, text=self.tr("field_co_exec"))
        self.co_exec_label.pack(pady=(4, 2))
        self.co_exec_entry = ctk.CTkEntry(self, width=form_width, placeholder_text=self.tr("co_exec_placeholder"))
        self.co_exec_entry.pack(pady=(0, 5))

        self.controller_label = ctk.CTkLabel(self, text=self.tr("field_controller"))
        self.controller_label.pack(pady=(4, 2))
        self.controller_entry = ctk.CTkEntry(self, width=form_width)
        self.controller_entry.pack(pady=(0, 5))

        self.due_date_label = ctk.CTkLabel(self, text=self.tr("field_due"))
        self.due_date_label.pack(pady=(4, 2))
        due_frame = ctk.CTkFrame(self, fg_color="transparent")
        due_frame.pack(pady=(0, 5))

        self.due_date_entry = ctk.CTkEntry(due_frame, width=460, placeholder_text="dd.mm.yyyy")
        self.due_date_entry.pack(side="left", padx=5)
        ctk.CTkButton(due_frame, text=self.tr("btn_pick"), width=120, height=34, corner_radius=10, command=lambda: self.open_calendar(self.due_date_entry)).pack(side="left", padx=5)

        self.status_label = ctk.CTkLabel(self, text=self.tr("status"))
        self.status_label.pack(pady=(4, 2))
        self.status_combo = ctk.CTkComboBox(self, values=[self.display_status(v) for v in TASK_STATUS_VALUES], width=form_width)
        self.status_combo.pack(pady=(0, 5))
        self.status_combo.set(self.display_status(TASK_STATUS_NEW))

        self.description_label = ctk.CTkLabel(self, text=self.tr("field_description"))
        self.description_label.pack(pady=(5, 2))
        self.description_text = ctk.CTkTextbox(self, width=form_width, height=55)
        self.description_text.pack(pady=(0, 8))

        ctk.CTkButton(self, text=self.tr("btn_save"), command=self.save, width=230, height=34, corner_radius=10).pack(pady=(8, 12))

        if self.task:
            self.fill()

    def match_person_records(self, query: str):
        query = (query or "").strip().lower()
        if not query:
            return self.person_records[:10]
        return [item for item in self.person_records if query in item["fio"].lower()][:10]

    def show_person_dropdown(self, matches):
        for child in self.person_dropdown.winfo_children():
            child.destroy()

        self.displayed_person_matches = matches
        if not matches:
            self.hide_person_dropdown()
            return

        for item in matches:
            btn = ctk.CTkButton(
                self.person_dropdown,
                text=f'{item["fio"]}   [ID {item["id"]}]',
                anchor="w",
                fg_color="transparent",
                hover_color=("#E5E5E5", "#3A3A3A"),
                text_color=("#111111", "#F0F0F0"),
                command=lambda person=item: self.select_person(person),
            )
            btn.pack(fill="x", padx=2, pady=1)

        if not self.person_dropdown_visible:
            self.person_dropdown_host.configure(height=90)
            self.person_dropdown.pack(fill="both", expand=True)
            self.person_dropdown_visible = True

    def hide_person_dropdown(self):
        if self.person_dropdown_visible:
            self.person_dropdown.pack_forget()
            self.person_dropdown_visible = False
        self.person_dropdown_host.configure(height=1)
        self.displayed_person_matches = []

    def select_person(self, person):
        self.selected_person_id = person["id"]
        self.selected_person_fio = person["fio"]
        self.person_entry.delete(0, "end")
        self.person_entry.insert(0, person["fio"])
        self.hide_person_dropdown()

    def on_person_entry_focus_in(self, event=None):
        self.show_person_dropdown(self.match_person_records(self.person_entry.get()))

    def on_person_entry_keyrelease(self, event=None):
        current_text = self.person_entry.get().strip()
        if current_text != self.selected_person_fio:
            self.selected_person_id = None
            self.selected_person_fio = ""
        self.show_person_dropdown(self.match_person_records(current_text))

    def on_person_entry_return(self, event=None):
        if self.displayed_person_matches:
            self.select_person(self.displayed_person_matches[0])
            return "break"
        return None

    def parse_person_id(self):
        raw = self.person_entry.get().strip()
        if not raw:
            return None
        if self.selected_person_id and raw == self.selected_person_fio:
            return self.selected_person_id
        exact = [item for item in self.person_records if item["fio"].lower() == raw.lower()]
        if len(exact) == 1:
            return exact[0]["id"]
        return None

    def parse_date(self, raw_value: str):
        raw_value = raw_value.strip()
        if not raw_value:
            return None
        try:
            return datetime.strptime(raw_value, "%d.%m.%Y").date()
        except ValueError:
            return "INVALID"

    def open_calendar(self, target_entry):
        win = ctk.CTkToplevel(self)
        win.geometry("320x360")
        win.title(self.tr("date_pick_title"))
        win.resizable(False, False)
        win.grab_set()

        cal = Calendar(win, date_pattern="dd.mm.yyyy")
        cal.pack(expand=True, fill="both", padx=10, pady=10)

        def apply():
            target_entry.delete(0, "end")
            target_entry.insert(0, cal.get_date())
            win.destroy()

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(pady=(0, 10))
        ctk.CTkButton(btn_frame, text="OK", width=100, height=34, corner_radius=10, command=apply).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=self.tr("btn_cancel"), width=100, height=34, corner_radius=10, command=win.destroy).pack(side="left", padx=5)

    def fill(self):
        if not self.task:
            return

        if getattr(self.task, "person", None):
            self.selected_person_id = self.task.person.id
            self.selected_person_fio = self.task.person.fio or ""
            self.person_entry.insert(0, self.selected_person_fio)

        self.title_entry.insert(0, self.task.title or "")
        self.main_responsible_combo.set(self.task.main_responsible or "")
        self.co_exec_entry.insert(0, self.task.co_executors or "")
        self.controller_entry.insert(0, self.task.controller or "")

        if self.task.due_date:
            self.due_date_entry.insert(0, self.task.due_date.strftime("%d.%m.%Y"))

        self.status_combo.set(self.display_status(self.task.status or TASK_STATUS_NEW))
        self.description_text.insert("1.0", self.task.description or "")

    def save(self):
        due_date = self.parse_date(self.due_date_entry.get())
        if due_date == "INVALID":
            self.show_error(self.tr("msg_due_invalid"))
            return

        payload = {
            "person_id": self.parse_person_id(),
            "title": self.title_entry.get().strip(),
            "description": self.description_text.get("1.0", "end").strip(),
            "main_responsible": self.main_responsible_combo.get().strip(),
            "co_executors": self.co_exec_entry.get().strip(),
            "controller": self.controller_entry.get().strip(),
            "due_date": due_date,
            "status": self.status_from_display(self.status_combo.get().strip()),
        }

        try:
            if self.mode == "edit":
                task = self.parent_tab.task_service.update_task(self.task_id, payload)
                if not task:
                    self.show_error(self.tr("msg_task_not_found"))
                    return
            else:
                self.parent_tab.task_service.create_task(payload)
        except ValueError as e:
            self.show_error(str(e))
            return

        self.parent_tab.refresh()
        self.destroy()

    def show_error(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("420x160")
        win.title(self.tr("title_error"))
        win.grab_set()
        ctk.CTkLabel(win, text=text, wraplength=380).pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy, width=120, height=34, corner_radius=10).pack()
