import customtkinter as ctk
from tkinter import ttk
from tkcalendar import Calendar
from datetime import datetime

from services.meeting_service import (
    MEETING_STATUS_VALUES,
    MEETING_STATUS_ACTIVE,
    MEETING_STATUS_PAUSED,
    MEETING_STATUS_DONE,
    MEETING_STATUS_CANCELLED,
    MEETING_STATUS_OVERDUE,
)


class MeetingsTab(ctk.CTkFrame):
    def __init__(self, parent, meeting_service):
        super().__init__(parent)
        self.meeting_service = meeting_service
        self.selected_meeting_id = None

        top_frame = ctk.CTkFrame(self, corner_radius=14)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkButton(top_frame, text="➕ Новая встреча", command=self.open_add_window, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="✏️ Редактировать", command=self.open_edit_window, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="🗑 Удалить", command=self.delete_selected, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="📋 Сообщение", command=self.copy_summary, width=150, height=34, corner_radius=10).pack(side="right", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="🔄 Обновить", command=self.refresh, width=140, height=34, corner_radius=10).pack(side="right", padx=5, pady=8)

        filter_frame = ctk.CTkFrame(self, corner_radius=14)
        filter_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(filter_frame, text="Поиск").pack(side="left", padx=(10, 5), pady=10)
        self.search_entry = ctk.CTkEntry(filter_frame, width=260, placeholder_text="Контакт / периодичность / заметки...")
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

        ctk.CTkLabel(filter_frame, text="Статус").pack(side="left", padx=(15, 5), pady=10)
        self.status_filter = ctk.CTkComboBox(
            filter_frame,
            values=["Все"] + MEETING_STATUS_VALUES,
            width=180,
            command=lambda _: self.refresh(),
        )
        self.status_filter.pack(side="left", padx=5, pady=10)
        self.status_filter.set("Все")

        ctk.CTkButton(filter_frame, text="Сбросить фильтры", width=150, height=34, corner_radius=10, command=self.reset_filters).pack(side="right", padx=10, pady=10)

        self.dashboard_frame = ctk.CTkFrame(self, corner_radius=14)
        self.dashboard_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.overdue_label = ctk.CTkLabel(self.dashboard_frame, text="🟥 Просрочены: 0")
        self.overdue_label.pack(side="left", padx=15, pady=10)
        self.active_label = ctk.CTkLabel(self.dashboard_frame, text="🟦 Активны: 0")
        self.active_label.pack(side="left", padx=15, pady=10)
        self.done_label = ctk.CTkLabel(self.dashboard_frame, text="🟩 Проведены: 0")
        self.done_label.pack(side="left", padx=15, pady=10)
        self.week_label = ctk.CTkLabel(self.dashboard_frame, text="🟣 Ближайшие 9 дней: 0")
        self.week_label.pack(side="left", padx=15, pady=10)

        body_frame = ctk.CTkFrame(self, fg_color="transparent")
        body_frame.pack(fill="both", expand=True, padx=10, pady=10)
        body_frame.grid_columnconfigure(0, weight=3)
        body_frame.grid_columnconfigure(1, weight=2)
        body_frame.grid_rowconfigure(0, weight=1)
        body_frame.grid_rowconfigure(1, weight=1)

        table_frame = ctk.CTkFrame(body_frame, corner_radius=14)
        table_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 8), pady=0)

        ctk.CTkLabel(table_frame, text="Реестр регулярных встреч", font=ctk.CTkFont(size=18, weight="normal")).pack(anchor="w", padx=12, pady=(10, 8))

        columns = ("id", "contact", "recurrence", "deadline", "days_left", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.tree.heading("id", text="ID")
        self.tree.heading("contact", text="Контакт")
        self.tree.heading("recurrence", text="Периодичность")
        self.tree.heading("deadline", text="Дедлайн")
        self.tree.heading("days_left", text="Через")
        self.tree.heading("status", text="Статус")

        self.tree.column("id", width=55, anchor="center")
        self.tree.column("contact", width=220)
        self.tree.column("recurrence", width=180)
        self.tree.column("deadline", width=110, anchor="center")
        self.tree.column("days_left", width=90, anchor="center")
        self.tree.column("status", width=120, anchor="center")

        self.tree.tag_configure("overdue", background="#ffdddd")
        self.tree.tag_configure("soon", background="#fff4cc")
        self.tree.tag_configure("active", background="#ddffdd")
        self.tree.tag_configure("paused", background="#eeeeee")

        self.tree.pack(side="left", fill="both", expand=True, padx=(8,0), pady=8)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y", padx=(0,8), pady=8)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_double_click)

        upcoming_frame = ctk.CTkFrame(body_frame, corner_radius=14)
        upcoming_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 8))
        ctk.CTkLabel(upcoming_frame, text="Встречи ближайшей недели", font=ctk.CTkFont(size=18, weight="normal")).pack(anchor="w", padx=12, pady=(10, 8))

        columns2 = ("contact", "recurrence", "deadline", "days_left")
        self.upcoming_tree = ttk.Treeview(upcoming_frame, columns=columns2, show="headings", height=8)
        self.upcoming_tree.heading("contact", text="Контакт")
        self.upcoming_tree.heading("recurrence", text="Периодичность")
        self.upcoming_tree.heading("deadline", text="Дедлайн")
        self.upcoming_tree.heading("days_left", text="Через")
        self.upcoming_tree.column("contact", width=180)
        self.upcoming_tree.column("recurrence", width=140)
        self.upcoming_tree.column("deadline", width=100, anchor="center")
        self.upcoming_tree.column("days_left", width=80, anchor="center")
        self.upcoming_tree.pack(fill="both", expand=True, padx=8, pady=(0,8))

        note_frame = ctk.CTkFrame(body_frame, corner_radius=14)
        note_frame.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(8, 0))
        ctk.CTkLabel(note_frame, text="Смысл модуля", font=ctk.CTkFont(size=18, weight="normal")).pack(anchor="w", padx=12, pady=(10, 8))
        text = (
            "Здесь фиксируется не само календарное событие, а регулярность встречи и дедлайн её планирования.\n\n"
            "Секретари уже создают конкретное событие сами, когда приходит время."
        )
        box = ctk.CTkTextbox(note_frame, height=120)
        box.pack(fill="both", expand=True, padx=8, pady=(0,8))
        box.insert("1.0", text)
        box.configure(state="disabled")

        self.refresh()

    def reset_filters(self):
        self.search_entry.delete(0, "end")
        self.status_filter.set("Все")
        self.refresh()

    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for row in self.upcoming_tree.get_children():
            self.upcoming_tree.delete(row)

        rows = self.meeting_service.list_meetings(
            search=self.search_entry.get().strip(),
            status=self.status_filter.get().strip(),
        )
        upcoming_rows = self.meeting_service.get_upcoming_meetings(days=9)
        counts = self.meeting_service.get_status_counts()

        self.overdue_label.configure(text=f"🟥 Просрочены: {counts.get(MEETING_STATUS_OVERDUE, 0)}")
        self.active_label.configure(text=f"🟦 Активны: {counts.get(MEETING_STATUS_ACTIVE, 0)}")
        self.done_label.configure(text=f"🟩 Проведены: {counts.get(MEETING_STATUS_DONE, 0)}")
        self.week_label.configure(text=f"🟣 Ближайшие 9 дней: {len(upcoming_rows)}")

        for row in rows:
            meeting = row["meeting"]
            contact_name = row["contact_name"] or "Без контакта"
            recurrence = meeting.recurrence_rule or ""
            status = row["status"]
            days_left = row["days_left"]
            deadline = meeting.start_datetime.strftime("%d.%m.%Y") if meeting.start_datetime else ""
            days_text = "—"
            if days_left is not None:
                if days_left < 0:
                    days_text = f"{abs(days_left)} дн. назад"
                elif days_left == 0:
                    days_text = "сегодня"
                else:
                    days_text = f"{days_left} дн."

            tag = "active"
            if status == MEETING_STATUS_OVERDUE:
                tag = "overdue"
            elif status == MEETING_STATUS_PAUSED:
                tag = "paused"
            elif days_left is not None and days_left <= 3:
                tag = "soon"

            self.tree.insert(
                "",
                "end",
                values=(meeting.id, contact_name, recurrence, deadline, days_text, status),
                tags=(tag,),
            )

        for row in upcoming_rows[:12]:
            meeting = row["meeting"]
            deadline = meeting.start_datetime.strftime("%d.%m.%Y") if meeting.start_datetime else ""
            days_left = row["days_left"]
            if days_left is None:
                days_text = "—"
            elif days_left == 0:
                days_text = "сегодня"
            elif days_left < 0:
                days_text = f"{abs(days_left)} назад"
            else:
                days_text = f"{days_left} дн."

            self.upcoming_tree.insert(
                "",
                "end",
                values=(row["contact_name"] or "Без контакта", meeting.recurrence_rule or "", deadline, days_text),
            )

        self.selected_meeting_id = None

    def on_select(self, event=None):
        selected = self.tree.selection()
        if not selected:
            self.selected_meeting_id = None
            return
        raw_id = self.tree.item(selected[0])["values"][0]
        try:
            self.selected_meeting_id = int(raw_id)
        except (TypeError, ValueError):
            self.selected_meeting_id = None

    def on_double_click(self, event):
        self.open_edit_window()

    def open_add_window(self):
        MeetingWindow(self, mode="add")

    def open_edit_window(self):
        if self.selected_meeting_id is None:
            self.show_message("Выбери встречу")
            return
        MeetingWindow(self, mode="edit", meeting_id=self.selected_meeting_id)

    def delete_selected(self):
        if self.selected_meeting_id is None:
            self.show_message("Выбери встречу")
            return
        deleted = self.meeting_service.delete_meeting(self.selected_meeting_id)
        if not deleted:
            self.show_message("Встреча не найдена")
            return
        self.refresh()

    def copy_summary(self):
        text = self.meeting_service.generate_summary(days=9)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.show_message("Сообщение по встречам скопировано в буфер обмена")

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("460x180")
        win.title("Сообщение")
        win.grab_set()
        ctk.CTkLabel(win, text=text, wraplength=420, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy, width=120, height=34, corner_radius=10).pack(pady=10)


class MeetingWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab, mode="add", meeting_id=None):
        super().__init__()
        self.parent_tab = parent_tab
        self.mode = mode
        self.meeting_id = meeting_id
        self.meeting = None
        if mode == "edit":
            self.meeting = self.parent_tab.meeting_service.get_meeting(meeting_id)

        self.title("Новая встреча" if mode == "add" else "Редактирование встречи")
        self.geometry("920x700")
        self.minsize(880, 660)
        self.resizable(True, True)
        self.grab_set()

        self.person_records = self.parent_tab.meeting_service.get_person_records()
        self.selected_person_id = None
        self.selected_person_fio = ""
        self.displayed_person_matches = []
        self.person_dropdown_visible = False

        self.build()

    def build(self):
        form_width = 740

        self.person_label = ctk.CTkLabel(self, text="Контакт")
        self.person_label.pack(pady=(10, 3))
        self.person_entry = ctk.CTkEntry(self, width=form_width, placeholder_text="Начните вводить ФИО...")
        self.person_entry.pack(pady=(0, 4))
        self.person_entry.bind("<KeyRelease>", self.on_person_entry_keyrelease)
        self.person_entry.bind("<FocusIn>", self.on_person_entry_focus_in)
        self.person_entry.bind("<Return>", self.on_person_entry_return)

        self.person_dropdown_host = ctk.CTkFrame(self, fg_color="transparent", width=form_width, height=90)
        self.person_dropdown_host.pack(pady=(0, 8))
        self.person_dropdown_host.pack_propagate(False)
        self.person_dropdown = ctk.CTkScrollableFrame(self.person_dropdown_host, width=form_width, height=90)
        self.person_dropdown.pack_forget()

        self.repeat_label = ctk.CTkLabel(self, text="Периодичность встречи")
        self.repeat_label.pack(pady=(4, 2))
        self.repeat_entry = ctk.CTkEntry(self, width=form_width, placeholder_text="Например: еженедельно / раз в месяц / раз в квартал")
        self.repeat_entry.pack(pady=(0, 8))

        self.deadline_label = ctk.CTkLabel(self, text="Дедлайн планирования")
        self.deadline_label.pack(pady=(4, 2))
        deadline_frame = ctk.CTkFrame(self, fg_color="transparent")
        deadline_frame.pack(pady=(0, 8))
        self.deadline_entry = ctk.CTkEntry(deadline_frame, width=460, placeholder_text="дд.мм.гггг")
        self.deadline_entry.pack(side="left", padx=5)
        ctk.CTkButton(deadline_frame, text="📅 Выбрать", width=120, height=34, corner_radius=10, command=lambda: self.open_calendar(self.deadline_entry)).pack(side="left", padx=5)

        self.status_label = ctk.CTkLabel(self, text="Статус")
        self.status_label.pack(pady=(4, 2))
        self.status_combo = ctk.CTkComboBox(self, values=MEETING_STATUS_VALUES, width=form_width)
        self.status_combo.pack(pady=(0, 8))
        self.status_combo.set(MEETING_STATUS_ACTIVE)

        self.notes_label = ctk.CTkLabel(self, text="Заметки")
        self.notes_label.pack(pady=(4, 2))
        self.notes_text = ctk.CTkTextbox(self, width=form_width, height=140)
        self.notes_text.pack(pady=(0, 10))

        ctk.CTkButton(self, text="Сохранить", command=self.save, width=220, height=34, corner_radius=10).pack(pady=(8, 12))

        if self.meeting:
            self.fill()

    def match_person_records(self, query: str):
        query = (query or "").strip().lower()
        if not query:
            return self.person_records[:12]
        return [item for item in self.person_records if query in item["fio"].lower()][:12]

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
            self.person_dropdown.pack(fill="both", expand=True)
            self.person_dropdown_visible = True

    def hide_person_dropdown(self):
        if self.person_dropdown_visible:
            self.person_dropdown.pack_forget()
            self.person_dropdown_visible = False
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
            return datetime.strptime(raw_value, "%d.%m.%Y")
        except ValueError:
            return "INVALID"

    def open_calendar(self, target_entry):
        win = ctk.CTkToplevel(self)
        win.geometry("320x360")
        win.title("Выбор даты")
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
        ctk.CTkButton(btn_frame, text="Отмена", width=100, height=34, corner_radius=10, command=win.destroy).pack(side="left", padx=5)

    def fill(self):
        if not self.meeting:
            return
        if getattr(self.meeting, "person", None):
            self.selected_person_id = self.meeting.person.id
            self.selected_person_fio = self.meeting.person.fio or ""
            self.person_entry.insert(0, self.selected_person_fio)
        elif getattr(self.meeting, "subject", ""):
            self.person_entry.insert(0, self.meeting.subject)

        if self.meeting.start_datetime:
            self.deadline_entry.insert(0, self.meeting.start_datetime.strftime("%d.%m.%Y"))
        self.repeat_entry.insert(0, self.meeting.recurrence_rule or "")
        self.status_combo.set(self.meeting.status or MEETING_STATUS_ACTIVE)
        self.notes_text.insert("1.0", self.meeting.notes or "")

    def save(self):
        planning_deadline = self.parse_date(self.deadline_entry.get())
        if planning_deadline == "INVALID":
            self.show_error("Дедлайн планирования должен быть в формате дд.мм.гггг")
            return

        payload = {
            "person_id": self.parse_person_id(),
            "subject": self.person_entry.get().strip(),
            "planning_deadline": planning_deadline,
            "recurrence_rule": self.repeat_entry.get().strip(),
            "status": self.status_combo.get().strip(),
            "notes": self.notes_text.get("1.0", "end").strip(),
        }

        try:
            if self.mode == "edit":
                meeting = self.parent_tab.meeting_service.update_meeting(self.meeting_id, payload)
                if not meeting:
                    self.show_error("Встреча не найдена")
                    return
            else:
                self.parent_tab.meeting_service.create_meeting(payload)
        except ValueError as e:
            self.show_error(str(e))
            return

        self.parent_tab.refresh()
        self.destroy()

    def show_error(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("420x160")
        win.title("Ошибка")
        win.grab_set()
        ctk.CTkLabel(win, text=text, wraplength=380).pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy, width=120, height=34, corner_radius=10).pack()
