import customtkinter as ctk
from tkinter import ttk
from datetime import date, timedelta

from ui.statuses import STATUS_OVERDUE, STATUS_TODAY, STATUS_7_DAYS, normalize_language

from services.person_service import PersonService
from services.interaction_service import InteractionService
from services.task_service import TaskService

try:
    from services.meeting_service import MeetingService
except Exception:
    MeetingService = None


TRANSLATIONS = {
    "ru": {
        "title": "Главная",
        "subtitle": "Контакты, задачи, встречи и напоминания одним взглядом",
        "refresh": "🔄 Обновить",
        "hint": "Двойной клик по строке открывает нужный модуль",
        "contacts_overdue": "Контакты просрочены",
        "contacts_today": "Контакты сегодня",
        "tasks_overdue": "Задачи просрочены",
        "meetings_soon": "Встречи ≤ 9 дней",
        "birthdays_today": "ДР сегодня",
        "card_overdue_contacts": "🔴 Просроченные контакты",
        "card_today_contacts": "🟡 Контакты на сегодня",
        "card_tasks_focus": "🟥 Задачи: фокус",
        "card_meetings": "🟣 Встречи: планировать",
        "card_birthdays": "🎂 Ближайшие дни рождения",
        "card_week_contacts": "🟢 Контакты на 7 дней",
        "col_object": "Объект",
        "col_date": "Дата",
        "col_details": "Детали",
        "meeting_fallback": "Встреча",
        "tab_interactions": "Контакты",
        "tab_tasks": "Задачи",
        "tab_meetings": "Встречи",
        "tab_reminders": "Напоминания",
    },
    "en": {
        "title": "Home",
        "subtitle": "Contacts, tasks, meetings, and reminders at a glance",
        "refresh": "🔄 Refresh",
        "hint": "Double-click a row to open the corresponding module",
        "contacts_overdue": "Overdue contacts",
        "contacts_today": "Contacts today",
        "tasks_overdue": "Overdue tasks",
        "meetings_soon": "Meetings ≤ 9 days",
        "birthdays_today": "Birthdays today",
        "card_overdue_contacts": "🔴 Overdue contacts",
        "card_today_contacts": "🟡 Contacts for today",
        "card_tasks_focus": "🟥 Tasks: focus",
        "card_meetings": "🟣 Meetings: plan",
        "card_birthdays": "🎂 Upcoming birthdays",
        "card_week_contacts": "🟢 Contacts for 7 days",
        "col_object": "Item",
        "col_date": "Date",
        "col_details": "Details",
        "meeting_fallback": "Meeting",
        "tab_interactions": "Interactions",
        "tab_tasks": "Tasks",
        "tab_meetings": "Meetings",
        "tab_reminders": "Reminders",
    },
}

TYPE_TRANSLATIONS = {
    "ru": {"Звонок": "Звонок", "Встреча": "Встреча"},
    "en": {"Звонок": "Call", "Встреча": "Meeting"},
}


class DashboardTab(ctk.CTkFrame):
    def _get_current_language(self):
        root = self.winfo_toplevel()
        return normalize_language(getattr(root, "current_language", "ru"))

    @property
    def current_language(self):
        return self._get_current_language()

    def tr(self, key: str, **kwargs):
        text = TRANSLATIONS.get(self.current_language, TRANSLATIONS["ru"]).get(key, key)
        if kwargs:
            return text.format(**kwargs)
        return text

    def translate_type(self, value: str | None):
        if not value:
            return ""
        return TYPE_TRANSLATIONS.get(self.current_language, TYPE_TRANSLATIONS["ru"]).get(value, value)

    def _tab_title(self, key: str):
        root = self.winfo_toplevel()
        if hasattr(root, "tr"):
            try:
                return root.tr(key)
            except Exception:
                pass
        return self.tr(key)

    def __init__(
        self,
        parent,
        interaction_service=None,
        person_service=None,
        task_service=None,
        meeting_service=None,
    ):
        super().__init__(parent)

        self.interaction_service = interaction_service or InteractionService()
        self.person_service = person_service or PersonService()
        self.task_service = task_service or TaskService()
        self.meeting_service = meeting_service or (MeetingService() if MeetingService else None)

        self._build_ui()
        self.refresh()

    def _build_ui(self):
        self.top_frame = ctk.CTkFrame(self, corner_radius=14)
        self.top_frame.pack(fill="x", padx=10, pady=(10, 8))

        self.title_label = ctk.CTkLabel(
            self.top_frame,
            text=self.tr("title"),
            font=ctk.CTkFont(size=26, weight="bold"),
        )
        self.title_label.pack(side="left", padx=16, pady=14)

        self.subtitle_label = ctk.CTkLabel(
            self.top_frame,
            text=self.tr("subtitle"),
            font=ctk.CTkFont(size=14),
            text_color=("gray35", "gray70"),
        )
        self.subtitle_label.pack(side="left", padx=(0, 10), pady=14)

        self.refresh_button = ctk.CTkButton(self.top_frame, text=self.tr("refresh"), width=150, command=self.refresh)
        self.refresh_button.pack(side="right", padx=12, pady=12)

        self.hint_label = ctk.CTkLabel(
            self,
            text=self.tr("hint"),
            font=ctk.CTkFont(size=13),
            text_color=("gray35", "gray70"),
        )
        self.hint_label.pack(anchor="w", padx=14, pady=(0, 8))

        self.counters_frame = ctk.CTkFrame(self, corner_radius=14)
        self.counters_frame.pack(fill="x", padx=10, pady=(0, 10))
        for i in range(5):
            self.counters_frame.grid_columnconfigure(i, weight=1)

        self.counter_contacts_overdue = self._create_counter(self.counters_frame, 0, "🔴", self.tr("contacts_overdue"))
        self.counter_contacts_today = self._create_counter(self.counters_frame, 1, "🟡", self.tr("contacts_today"))
        self.counter_tasks_overdue = self._create_counter(self.counters_frame, 2, "🟥", self.tr("tasks_overdue"))
        self.counter_meetings = self._create_counter(self.counters_frame, 3, "🟣", self.tr("meetings_soon"))
        self.counter_birthdays = self._create_counter(self.counters_frame, 4, "🎂", self.tr("birthdays_today"))

        self.cards_host = ctk.CTkFrame(self, fg_color="transparent")
        self.cards_host.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        for col in range(3):
            self.cards_host.grid_columnconfigure(col, weight=1, uniform="cards")
        for row in range(2):
            self.cards_host.grid_rowconfigure(row, weight=1)

        self.card_contacts_overdue = self._create_card(self.cards_host, self.tr("card_overdue_contacts"))
        self.card_contacts_today = self._create_card(self.cards_host, self.tr("card_today_contacts"))
        self.card_tasks_focus = self._create_card(self.cards_host, self.tr("card_tasks_focus"))
        self.card_meetings = self._create_card(self.cards_host, self.tr("card_meetings"))
        self.card_birthdays = self._create_card(self.cards_host, self.tr("card_birthdays"))
        self.card_contacts_week = self._create_card(self.cards_host, self.tr("card_week_contacts"))

        self.card_contacts_overdue["frame"].grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
        self.card_contacts_today["frame"].grid(row=0, column=1, sticky="nsew", padx=6, pady=(0, 6))
        self.card_tasks_focus["frame"].grid(row=0, column=2, sticky="nsew", padx=(6, 0), pady=(0, 6))

        self.card_meetings["frame"].grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(6, 0))
        self.card_birthdays["frame"].grid(row=1, column=1, sticky="nsew", padx=6, pady=(6, 0))
        self.card_contacts_week["frame"].grid(row=1, column=2, sticky="nsew", padx=(6, 0), pady=(6, 0))

        self.card_contacts_overdue["tree"].bind("<Double-1>", lambda e: self._open_tab("tab_interactions"))
        self.card_contacts_today["tree"].bind("<Double-1>", lambda e: self._open_tab("tab_interactions"))
        self.card_contacts_week["tree"].bind("<Double-1>", lambda e: self._open_tab("tab_interactions"))
        self.card_tasks_focus["tree"].bind("<Double-1>", lambda e: self._open_tab("tab_tasks"))
        self.card_meetings["tree"].bind("<Double-1>", lambda e: self._open_tab("tab_meetings"))
        self.card_birthdays["tree"].bind("<Double-1>", lambda e: self._open_tab("tab_reminders"))

    def _create_counter(self, parent, column, emoji, title):
        box = ctk.CTkFrame(parent, corner_radius=12)
        box.grid(row=0, column=column, sticky="nsew", padx=6, pady=8)

        emoji_label = ctk.CTkLabel(box, text=emoji, font=ctk.CTkFont(size=20))
        emoji_label.pack(pady=(12, 2))

        value_label = ctk.CTkLabel(box, text="0", font=ctk.CTkFont(size=28, weight="bold"))
        value_label.pack(pady=(0, 2))

        title_label = ctk.CTkLabel(
            box,
            text=title,
            font=ctk.CTkFont(size=13),
            text_color=("gray35", "gray70"),
        )
        title_label.pack(pady=(0, 12))

        return {"frame": box, "value": value_label, "title": title_label}

    def _create_card(self, parent, title):
        frame = ctk.CTkFrame(parent, corner_radius=14)

        title_label = ctk.CTkLabel(
            frame,
            text=title,
            anchor="w",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        title_label.pack(fill="x", padx=12, pady=(10, 8))

        table_container = ctk.CTkFrame(frame, fg_color="transparent")
        table_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("main", "date", "extra")
        tree = ttk.Treeview(table_container, columns=columns, show="headings", height=9)
        tree.heading("main", text=self.tr("col_object"))
        tree.heading("date", text=self.tr("col_date"))
        tree.heading("extra", text=self.tr("col_details"))

        tree.column("main", width=230)
        tree.column("date", width=95, anchor="center")
        tree.column("extra", width=190)

        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar.set)

        return {"frame": frame, "title_label": title_label, "tree": tree}

    def update_ui_texts(self):
        self.title_label.configure(text=self.tr("title"))
        self.subtitle_label.configure(text=self.tr("subtitle"))
        self.hint_label.configure(text=self.tr("hint"))
        self.refresh_button.configure(text=self.tr("refresh"))

        self.counter_contacts_overdue["title"].configure(text=self.tr("contacts_overdue"))
        self.counter_contacts_today["title"].configure(text=self.tr("contacts_today"))
        self.counter_tasks_overdue["title"].configure(text=self.tr("tasks_overdue"))
        self.counter_meetings["title"].configure(text=self.tr("meetings_soon"))
        self.counter_birthdays["title"].configure(text=self.tr("birthdays_today"))

        self.card_contacts_overdue["title_label"].configure(text=self.tr("card_overdue_contacts"))
        self.card_contacts_today["title_label"].configure(text=self.tr("card_today_contacts"))
        self.card_tasks_focus["title_label"].configure(text=self.tr("card_tasks_focus"))
        self.card_meetings["title_label"].configure(text=self.tr("card_meetings"))
        self.card_birthdays["title_label"].configure(text=self.tr("card_birthdays"))
        self.card_contacts_week["title_label"].configure(text=self.tr("card_week_contacts"))

        for card in (
            self.card_contacts_overdue,
            self.card_contacts_today,
            self.card_tasks_focus,
            self.card_meetings,
            self.card_birthdays,
            self.card_contacts_week,
        ):
            tree = card["tree"]
            tree.heading("main", text=self.tr("col_object"))
            tree.heading("date", text=self.tr("col_date"))
            tree.heading("extra", text=self.tr("col_details"))

        self.refresh()

    def _open_tab(self, tab_key: str):
        root = self.winfo_toplevel()
        if hasattr(root, "tabview"):
            try:
                root.tabview.set(self._tab_title(tab_key))
            except Exception:
                pass

    def refresh(self):
        overdue_contacts = self._get_overdue_contacts()
        today_contacts = self._get_today_contacts()
        week_contacts = self._get_week_contacts()

        overdue_tasks = self._get_overdue_tasks()
        today_tasks = self._get_today_tasks()
        week_tasks = self._get_week_tasks()

        focus_tasks = overdue_tasks + [task for task in today_tasks if task not in overdue_tasks]
        if len(focus_tasks) < 12:
            for task in week_tasks:
                if task not in focus_tasks:
                    focus_tasks.append(task)

        today_birthdays = self._get_birthdays_today()
        upcoming_birthdays = self._get_birthdays_upcoming()

        near_meetings = self._get_near_meetings(days=9)

        self.counter_contacts_overdue["value"].configure(text=str(len(overdue_contacts)))
        self.counter_contacts_today["value"].configure(text=str(len(today_contacts)))
        self.counter_tasks_overdue["value"].configure(text=str(len(overdue_tasks)))
        self.counter_meetings["value"].configure(text=str(len(near_meetings)))
        self.counter_birthdays["value"].configure(text=str(len(today_birthdays)))

        self._fill_contacts_tree(self.card_contacts_overdue["tree"], overdue_contacts)
        self._fill_contacts_tree(self.card_contacts_today["tree"], today_contacts)
        self._fill_tasks_tree(self.card_tasks_focus["tree"], focus_tasks[:12])
        self._fill_meetings_tree(self.card_meetings["tree"], near_meetings[:12])
        self._fill_birthdays_tree(self.card_birthdays["tree"], upcoming_birthdays[:12])
        self._fill_contacts_tree(self.card_contacts_week["tree"], week_contacts[:12])

    def _get_all_active_contacts(self):
        try:
            return self.interaction_service.list_active_interactions()
        except Exception:
            return []

    def _get_overdue_contacts(self):
        rows = self._get_all_active_contacts()
        return [row for row in rows if row["status"] == STATUS_OVERDUE]

    def _get_today_contacts(self):
        rows = self._get_all_active_contacts()
        return [row for row in rows if row["status"] == STATUS_TODAY]

    def _get_week_contacts(self):
        rows = self._get_all_active_contacts()
        return [row for row in rows if row["status"] in (STATUS_TODAY, STATUS_7_DAYS)]

    def _get_overdue_tasks(self):
        try:
            return self.task_service.get_overdue_tasks()
        except Exception:
            try:
                return self.task_service.list_tasks(status=STATUS_OVERDUE)
            except Exception:
                return []

    def _get_today_tasks(self):
        try:
            return self.task_service.get_today_tasks()
        except Exception:
            return []

    def _get_week_tasks(self):
        try:
            return self.task_service.get_upcoming_tasks(days=7)
        except Exception:
            return []

    def _get_birthdays_today(self):
        try:
            return self.person_service.get_birthdays_today()
        except Exception:
            return []

    def _get_birthdays_upcoming(self):
        try:
            return self.person_service.get_upcoming_birthdays(days=7)
        except Exception:
            return []

    def _get_near_meetings(self, days=9):
        if not self.meeting_service:
            return []

        today = date.today()
        limit = today + timedelta(days=days)

        try:
            rows = self.meeting_service.list_meetings(search="", status="Все")
        except TypeError:
            try:
                rows = self.meeting_service.list_meetings()
            except Exception:
                return []
        except Exception:
            return []

        result = []
        for row in rows:
            meeting = row["meeting"] if isinstance(row, dict) else getattr(row, "meeting", row)
            deadline = getattr(meeting, "start_datetime", None)
            if not deadline:
                continue

            deadline_date = deadline.date() if hasattr(deadline, "date") else deadline
            if today <= deadline_date <= limit:
                result.append(meeting)

        result.sort(key=lambda x: x.start_datetime or date.max)
        return result

    def _fill_contacts_tree(self, tree, rows):
        for item in tree.get_children():
            tree.delete(item)

        for row in rows:
            interaction = row["interaction"]
            person = row["person"]
            responsible = row["responsible"]

            person_name = person.fio if person else f"[ID {interaction.person_id}]"
            next_date = interaction.next_date.strftime("%d.%m.%Y") if interaction.next_date else ""
            extra = f"{self.translate_type(interaction.interaction_type) or ''} | {responsible or '-'}"

            tree.insert("", "end", values=(person_name, next_date, extra))

    def _fill_tasks_tree(self, tree, rows):
        for item in tree.get_children():
            tree.delete(item)

        for row in rows:
            task = row["task"] if isinstance(row, dict) else row
            person = row["person"] if isinstance(row, dict) else getattr(task, "person", None)
            status = row["status"] if isinstance(row, dict) else getattr(task, "status", "")

            title = getattr(task, "title", "") or ""
            due_date = task.due_date.strftime("%d.%m.%Y") if getattr(task, "due_date", None) else ""
            extra_parts = []
            if person and getattr(person, "fio", ""):
                extra_parts.append(person.fio)
            if getattr(task, "main_responsible", ""):
                extra_parts.append(task.main_responsible)
            if status:
                extra_parts.append(status)

            tree.insert("", "end", values=(title, due_date, " | ".join(extra_parts)))

    def _fill_birthdays_tree(self, tree, people):
        for item in tree.get_children():
            tree.delete(item)

        for person in people:
            birthday_str = person.birthday.strftime("%d.%m") if getattr(person, "birthday", None) else ""
            extra = f"{getattr(person, 'circle', '') or '-'} | {getattr(person, 'responsible', '') or '-'}"
            tree.insert("", "end", values=(person.fio, birthday_str, extra))

    def _fill_meetings_tree(self, tree, meetings):
        for item in tree.get_children():
            tree.delete(item)

        for meeting in meetings:
            person_name = ""
            if getattr(meeting, "person", None) and getattr(meeting.person, "fio", ""):
                person_name = meeting.person.fio
            elif getattr(meeting, "subject", ""):
                person_name = meeting.subject
            else:
                person_name = self.tr("meeting_fallback")

            deadline = ""
            if getattr(meeting, "start_datetime", None):
                deadline = meeting.start_datetime.strftime("%d.%m.%Y")

            extra_parts = []
            if getattr(meeting, "recurrence_rule", ""):
                extra_parts.append(meeting.recurrence_rule)
            if getattr(meeting, "status", ""):
                extra_parts.append(meeting.status)

            tree.insert("", "end", values=(person_name, deadline, " | ".join(extra_parts)))
