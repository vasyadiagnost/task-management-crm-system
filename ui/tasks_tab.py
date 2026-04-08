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


DISPLAY_STATUS_ORDER = {
    TASK_STATUS_OVERDUE: 0,
    TASK_STATUS_IN_PROGRESS: 1,
    TASK_STATUS_NEW: 2,
    TASK_STATUS_DONE: 3,
    TASK_STATUS_CANCELLED: 4,
}


class TasksTab(ctk.CTkFrame):
    def __init__(self, parent, task_service):
        super().__init__(parent)
        self.task_service = task_service
        self.selected_task_id = None

        self.settings_service = SettingsService()
        self.interaction_service = InteractionService(settings_service=self.settings_service)

        top_frame = ctk.CTkFrame(self, corner_radius=14)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkButton(top_frame, text="➕ Новая задача", command=self.open_add_window, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="✏️ Редактировать", command=self.open_edit_window, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="📄 Дублировать", command=self.duplicate_selected, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="🗑 Удалить", command=self.delete_selected, width=150, height=34, corner_radius=10).pack(side="left", padx=5, pady=8)
        ctk.CTkButton(top_frame, text="🔄 Обновить", command=self.refresh, width=140, height=34, corner_radius=10).pack(side="right", padx=5, pady=8)

        filter_frame = ctk.CTkFrame(self, corner_radius=14)
        filter_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(filter_frame, text="Поиск").pack(side="left", padx=(10, 5), pady=10)
        self.search_entry = ctk.CTkEntry(filter_frame, width=260, placeholder_text="Название / ФИО / ответственный...")
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

        ctk.CTkLabel(filter_frame, text="Статус").pack(side="left", padx=(15, 5), pady=10)
        self.status_filter = ctk.CTkComboBox(filter_frame, values=["Все"] + TASK_STATUS_VALUES, width=160, command=lambda _: self.refresh())
        self.status_filter.pack(side="left", padx=5, pady=10)
        self.status_filter.set("Все")

        ctk.CTkLabel(filter_frame, text="Ответственный").pack(side="left", padx=(15, 5), pady=10)
        self.responsible_filter = ctk.CTkComboBox(
            filter_frame,
            values=self.task_service.get_responsible_values(),
            width=180,
            command=lambda _: self.refresh(),
        )
        self.responsible_filter.pack(side="left", padx=5, pady=10)
        self.responsible_filter.set("Все")

        ctk.CTkButton(filter_frame, text="Мои задачи", width=120, height=34, corner_radius=10, command=self.apply_my_tasks_filter).pack(side="right", padx=(5, 10), pady=10)
        ctk.CTkButton(filter_frame, text="Только просроченные", width=170, height=34, corner_radius=10, command=self.apply_overdue_filter).pack(side="right", padx=5, pady=10)
        ctk.CTkButton(filter_frame, text="Сбросить фильтры", width=150, height=34, corner_radius=10, command=self.reset_filters).pack(side="right", padx=5, pady=10)

        self.dashboard_frame = ctk.CTkFrame(self, corner_radius=14)
        self.dashboard_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.new_label = ctk.CTkLabel(self.dashboard_frame, text="🟦 Новые: 0")
        self.new_label.pack(side="left", padx=15, pady=10)
        self.progress_label = ctk.CTkLabel(self.dashboard_frame, text="🟨 В работе: 0")
        self.progress_label.pack(side="left", padx=15, pady=10)
        self.overdue_label = ctk.CTkLabel(self.dashboard_frame, text="🟥 Просроченные: 0")
        self.overdue_label.pack(side="left", padx=15, pady=10)
        self.done_label = ctk.CTkLabel(self.dashboard_frame, text="🟩 Выполненные: 0")
        self.done_label.pack(side="left", padx=15, pady=10)

        table_frame = ctk.CTkFrame(self, corner_radius=14)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("id", "title", "person", "main_responsible", "co_executors", "controller", "due_date", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        headings = {
            "id": "ID",
            "title": "Задача",
            "person": "Контакт",
            "main_responsible": "Основной",
            "co_executors": "Соисполнители",
            "controller": "Контроль",
            "due_date": "Срок",
            "status": "Статус",
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

        self.refresh()

    def reset_filters(self):
        self.search_entry.delete(0, "end")
        self.status_filter.set("Все")
        self.responsible_filter.configure(values=self.task_service.get_responsible_values())
        self.responsible_filter.set("Все")
        self.refresh()

    def apply_my_tasks_filter(self):
        current_responsible = self.responsible_filter.get().strip()
        if not current_responsible or current_responsible == "Все":
            self.show_message("Сначала выбери ответственного в фильтре, и кнопка покажет его задачи")
            return
        self.status_filter.set("Все")
        self.refresh()

    def apply_overdue_filter(self):
        self.status_filter.set(TASK_STATUS_OVERDUE)
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
        responsible_values = self.task_service.get_responsible_values()
        self.responsible_filter.configure(values=responsible_values)
        if self.responsible_filter.get() not in responsible_values:
            self.responsible_filter.set("Все")

        for row in self.tree.get_children():
            self.tree.delete(row)

        rows = self.task_service.list_tasks(
            search=self.search_entry.get().strip(),
            status=self.status_filter.get().strip(),
            responsible=self.responsible_filter.get().strip(),
        )
        rows = sorted(rows, key=self._row_sort_key)

        counts = self.task_service.get_status_counts()
        self.new_label.configure(text=f"🟦 Новые: {counts.get(TASK_STATUS_NEW, 0)}")
        self.progress_label.configure(text=f"🟨 В работе: {counts.get(TASK_STATUS_IN_PROGRESS, 0)}")
        self.overdue_label.configure(text=f"🟥 Просроченные: {counts.get(TASK_STATUS_OVERDUE, 0)}")
        self.done_label.configure(text=f"🟩 Выполненные: {counts.get(TASK_STATUS_DONE, 0)}")

        for row in rows:
            task = row["task"]
            person = row["person"]
            status = row["status"]
            due_date = task.due_date.strftime("%d.%m.%Y") if task.due_date else ""

            display_status = status
            tag = ""
            if status == TASK_STATUS_OVERDUE:
                tag = "overdue"
                display_status = "Просрочена"
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
            self.show_message("Задача не найдена")
            return

        if self.task_service.is_execution_task(task):
            TaskExecutionWindow(self, task)
            return

        self.open_edit_window()

    def open_add_window(self):
        TaskWindow(self, mode="add")

    def open_edit_window(self):
        if self.selected_task_id is None:
            self.show_message("Выбери задачу")
            return
        TaskWindow(self, mode="edit", task_id=self.selected_task_id)

    def duplicate_selected(self):
        if self.selected_task_id is None:
            self.show_message("Выбери задачу")
            return

        task = self.task_service.get_task(self.selected_task_id)
        if not task:
            self.show_message("Задача не найдена")
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
            self.show_message("Выбери задачу")
            return

        deleted = self.task_service.delete_task(self.selected_task_id)
        if not deleted:
            self.show_message("Задача не найдена")
            return

        self.refresh()

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("460x180")
        win.title("Сообщение")
        win.grab_set()
        ctk.CTkLabel(win, text=text, wraplength=420, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy, width=120, height=34, corner_radius=10).pack(pady=10)


class TaskExecutionWindow(ctk.CTkToplevel):
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

        self.title(f"Отработка: {task.title or 'задача'}")
        self.geometry("980x760")
        self.minsize(900, 680)
        self.resizable(True, True)
        self.grab_set()

        self.build()

    def build(self):
        ctk.CTkLabel(self, text=self.task.title or "Отработка задачи", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(fill="x", padx=20, pady=(18, 6))

        meta = f"Ответственный: {self.task.main_responsible or '-'}"
        if self.task.due_date:
            meta += f" | Срок: {self.task.due_date.strftime('%d.%m.%Y')}"
        ctk.CTkLabel(self, text=meta, anchor="w").pack(fill="x", padx=20, pady=(0, 10))

        info_box = ctk.CTkTextbox(self, height=68)
        info_box.pack(fill="x", padx=20, pady=(0, 12))
        info_box.insert(
            "1.0",
            "Это безопасный режим отработки.\n"
            "«Отметить отработанным» подготавливает минимальный контакт автоматически.\n"
            "«Открыть контакт» позволяет вручную задать детали. Фактическая дата последнего контакта показывается прямо в строке.",
        )
        info_box.configure(state="disabled")

        list_host = ctk.CTkFrame(self, corner_radius=12)
        list_host.pack(fill="both", expand=True, padx=20, pady=(0, 14))

        self.list_frame = ctk.CTkScrollableFrame(list_host)
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        everyone = self.parent_tab.task_service.merge_execution_people(self.remaining_people, list(self.done_people))
        if not everyone:
            ctk.CTkLabel(self.list_frame, text="В описании задачи не найден список людей для отработки.", anchor="w").pack(fill="x", padx=10, pady=10)
        else:
            for fio in everyone:
                self._create_person_row(fio, is_done=(fio in self.done_people))

        bottom = ctk.CTkFrame(self, corner_radius=12)
        bottom.pack(fill="x", padx=20, pady=(0, 20))

        self.progress_label = ctk.CTkLabel(bottom, text=self._progress_text(), font=ctk.CTkFont(size=14, weight="bold"))
        self.progress_label.pack(side="left", padx=16, pady=14)

        self.save_btn = ctk.CTkButton(bottom, text="Сохранить прогресс", command=self.save_progress, width=170, height=34, corner_radius=10)
        self.save_btn.pack(side="right", padx=(8, 16), pady=10)

        self.finish_btn = ctk.CTkButton(
            bottom,
            text="Завершить задачу",
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

        open_btn = ctk.CTkButton(row, text="Открыть контакт", width=150, height=32, corner_radius=10, command=lambda current_fio=fio: self.open_contact_editor(current_fio))
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
            return f"Последний зафиксированный контакт: {payload['interaction_date'].strftime('%d.%m.%Y')}"

        saved_date = self.persisted_interaction_dates.get(fio)
        if saved_date:
            return f"Последний зафиксированный контакт: {saved_date}"

        return "Последний зафиксированный контакт: —"

    def _apply_row_state(self, fio: str, is_done: bool):
        widgets = self.row_widgets[fio]
        widgets["last_contact_label"].configure(text=self._last_contact_text(fio))

        if is_done:
            widgets["row"].configure(fg_color="#ddffdd")
            widgets["toggle_btn"].configure(text="↩ Вернуть в неотработанные")
            widgets["status_label"].configure(text="✅ Готово")
        else:
            widgets["row"].configure(fg_color=("gray86", "gray20"))
            widgets["toggle_btn"].configure(text="Отметить отработанным")
            widgets["status_label"].configure(text="⌛ Не отработано")

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
            "purpose": f"Отработка просроченного контакта по задаче: {self.task.title or ''}".strip(),
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
            self.parent_tab.show_message(f"Не удалось найти человека в базе по точному ФИО:\n{fio}")
            return

        initial_payload = self.staged_payloads.get(fio)
        if not initial_payload:
            initial_payload = self._default_payload_for_person(fio)
        if not initial_payload:
            self.parent_tab.show_message(f"Не удалось подготовить карточку контакта для:\n{fio}")
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
                    self.parent_tab.show_message(f"Не удалось автоматически подготовить контакт для:\n{fio}")
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
        return f"Отработано: {done_count} из {total}"

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
            self.parent_tab.show_message(f"Ошибка сохранения прогресса:\n{e}")
            return

        self.task = updated
        self.parent_tab.refresh()
        self.destroy()
        self.parent_tab.show_message("Прогресс по задаче сохранён")

    def finish_task(self):
        try:
            updated = self._persist_changes(mark_done=True)
        except Exception as e:
            self.parent_tab.show_message(f"Ошибка завершения задачи:\n{e}")
            return

        self.task = updated
        self.parent_tab.refresh()
        self.destroy()
        self.parent_tab.show_message("Задача переведена в статус «Выполнена»")


class ExecutionContactWindow(ctk.CTkToplevel):
    def __init__(self, execution_window, fio: str, person_record: dict, payload: dict):
        super().__init__()
        self.execution_window = execution_window
        self.fio = fio
        self.person_record = person_record
        self.payload = payload

        self.title(f"Контакт: {fio}")
        self.geometry("860x760")
        self.minsize(820, 680)
        self.resizable(True, True)
        self.grab_set()

        self.build()

    def build(self):
        content = ctk.CTkScrollableFrame(self)
        content.pack(fill="both", expand=True, padx=14, pady=14)

        form_width = 700

        ctk.CTkLabel(content, text="Человек").pack(pady=(8, 2))
        self.person_entry = ctk.CTkEntry(content, width=form_width)
        self.person_entry.pack(pady=(0, 6))
        self.person_entry.insert(0, self.person_record["fio"])
        self.person_entry.configure(state="disabled")

        ctk.CTkLabel(content, text="Тип контакта").pack(pady=(8, 2))
        self.type_combo = ctk.CTkComboBox(content, values=["Звонок", "Встреча"], width=form_width)
        self.type_combo.pack(pady=(0, 6))
        self.type_combo.set(self.payload.get("interaction_type") or "Звонок")

        ctk.CTkLabel(content, text="Ответственный").pack(pady=(8, 2))
        self.responsible_entry = ctk.CTkEntry(content, width=form_width)
        self.responsible_entry.pack(pady=(0, 6))
        self.responsible_entry.insert(0, self.payload.get("responsible") or "")

        ctk.CTkLabel(content, text="Дата контакта").pack(pady=(8, 2))
        contact_frame = ctk.CTkFrame(content, fg_color="transparent")
        contact_frame.pack(pady=(0, 6))
        self.contact_date_entry = ctk.CTkEntry(contact_frame, width=460, placeholder_text="дд.мм.гггг")
        self.contact_date_entry.pack(side="left", padx=5)
        if self.payload.get("interaction_date"):
            self.contact_date_entry.insert(0, self.payload["interaction_date"].strftime("%d.%m.%Y"))
        self.contact_date_entry.bind("<FocusOut>", lambda event: self.autofill_next_date())
        ctk.CTkButton(contact_frame, text="📅 Выбрать", width=120, height=34, corner_radius=10, command=lambda: self.open_calendar(self.contact_date_entry, self.autofill_next_date)).pack(side="left", padx=5)

        ctk.CTkLabel(content, text="Дата следующего контакта").pack(pady=(8, 2))
        next_frame = ctk.CTkFrame(content, fg_color="transparent")
        next_frame.pack(pady=(0, 6))
        self.next_date_entry = ctk.CTkEntry(next_frame, width=460, placeholder_text="дд.мм.гггг")
        self.next_date_entry.pack(side="left", padx=5)
        if self.payload.get("next_date"):
            self.next_date_entry.insert(0, self.payload["next_date"].strftime("%d.%m.%Y"))
        ctk.CTkButton(next_frame, text="📅 Выбрать", width=120, height=34, corner_radius=10, command=lambda: self.open_calendar(self.next_date_entry)).pack(side="left", padx=5)

        ctk.CTkLabel(content, text="Цель контакта").pack(pady=(8, 2))
        self.purpose_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.purpose_text.pack(pady=(0, 6))
        self.purpose_text.insert("1.0", self.payload.get("purpose") or "")

        ctk.CTkLabel(content, text="Результат").pack(pady=(8, 2))
        self.result_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.result_text.pack(pady=(0, 6))
        self.result_text.insert("1.0", self.payload.get("result") or "")

        ctk.CTkLabel(content, text="Комментарий").pack(pady=(8, 2))
        self.comment_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.comment_text.pack(pady=(0, 10))
        self.comment_text.insert("1.0", self.payload.get("comment") or "")

        buttons = ctk.CTkFrame(content, fg_color="transparent")
        buttons.pack(fill="x", pady=(8, 12))

        ctk.CTkButton(buttons, text="Сохранить карточку", width=180, height=34, corner_radius=10, command=self.save_card).pack(side="left", padx=6)
        ctk.CTkButton(buttons, text="Отмена", width=120, height=34, corner_radius=10, command=self.destroy).pack(side="left", padx=6)

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
        win.title("Выбор даты")
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
        ctk.CTkButton(btn_frame, text="Отмена", width=100, height=34, corner_radius=10, command=win.destroy).pack(side="left", padx=5)

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
            self.execution_window.parent_tab.show_message("Дата контакта должна быть в формате дд.мм.гггг")
            return

        next_date = self.parse_date(self.next_date_entry.get())
        if next_date == "INVALID":
            self.execution_window.parent_tab.show_message("Дата следующего контакта должна быть в формате дд.мм.гггг")
            return

        payload = {
            "person_id": self.person_record["id"],
            "interaction_type": self.type_combo.get().strip(),
            "interaction_date": interaction_date or date.today(),
            "next_date": next_date,
            "responsible": self.responsible_entry.get().strip(),
            "purpose": self.purpose_text.get("1.0", "end").strip(),
            "result": self.result_text.get("1.0", "end").strip(),
            "comment": self.comment_text.get("1.0", "end").strip(),
        }
        self.execution_window.apply_contact_payload(self.fio, payload)
        self.destroy()


class TaskWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab, mode="add", task_id=None):
        super().__init__()
        self.parent_tab = parent_tab
        self.mode = mode
        self.task_id = task_id
        self.task = None

        if mode == "edit":
            self.task = self.parent_tab.task_service.get_task(task_id)

        self.title("Новая задача" if mode == "add" else "Редактирование задачи")
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

        self.person_label = ctk.CTkLabel(self, text="Контакт")
        self.person_label.pack(pady=(10, 3))
        self.person_entry = ctk.CTkEntry(self, width=form_width, placeholder_text="Необязательно. Начните вводить ФИО...")
        self.person_entry.pack(pady=(0, 4))
        self.person_entry.bind("<KeyRelease>", self.on_person_entry_keyrelease)
        self.person_entry.bind("<FocusIn>", self.on_person_entry_focus_in)
        self.person_entry.bind("<Return>", self.on_person_entry_return)

        self.person_dropdown_host = ctk.CTkFrame(self, fg_color="transparent", width=form_width, height=1)
        self.person_dropdown_host.pack(pady=(0, 4))
        self.person_dropdown_host.pack_propagate(False)

        self.person_dropdown = ctk.CTkScrollableFrame(self.person_dropdown_host, width=form_width, height=90)
        self.person_dropdown.pack_forget()

        self.title_label = ctk.CTkLabel(self, text="Название задачи")
        self.title_label.pack(pady=(4, 2))
        self.title_entry = ctk.CTkEntry(self, width=form_width)
        self.title_entry.pack(pady=(0, 5))

        self.main_responsible_label = ctk.CTkLabel(self, text="Основной ответственный")
        self.main_responsible_label.pack(pady=(4, 2))
        self.main_responsible_combo = ctk.CTkComboBox(self, values=[v for v in self.parent_tab.task_service.get_responsible_values() if v != "Все"] or [""], width=form_width)
        self.main_responsible_combo.pack(pady=(0, 5))

        self.co_exec_label = ctk.CTkLabel(self, text="Соисполнители")
        self.co_exec_label.pack(pady=(4, 2))
        self.co_exec_entry = ctk.CTkEntry(self, width=form_width, placeholder_text="Через запятую")
        self.co_exec_entry.pack(pady=(0, 5))

        self.controller_label = ctk.CTkLabel(self, text="Контроль")
        self.controller_label.pack(pady=(4, 2))
        self.controller_entry = ctk.CTkEntry(self, width=form_width)
        self.controller_entry.pack(pady=(0, 5))

        self.due_date_label = ctk.CTkLabel(self, text="Срок")
        self.due_date_label.pack(pady=(4, 2))
        due_frame = ctk.CTkFrame(self, fg_color="transparent")
        due_frame.pack(pady=(0, 5))

        self.due_date_entry = ctk.CTkEntry(due_frame, width=460, placeholder_text="дд.мм.гггг")
        self.due_date_entry.pack(side="left", padx=5)
        ctk.CTkButton(due_frame, text="📅 Выбрать", width=120, height=34, corner_radius=10, command=lambda: self.open_calendar(self.due_date_entry)).pack(side="left", padx=5)

        self.status_label = ctk.CTkLabel(self, text="Статус")
        self.status_label.pack(pady=(4, 2))
        self.status_combo = ctk.CTkComboBox(self, values=TASK_STATUS_VALUES, width=form_width)
        self.status_combo.pack(pady=(0, 5))
        self.status_combo.set(TASK_STATUS_NEW)

        self.description_label = ctk.CTkLabel(self, text="Описание")
        self.description_label.pack(pady=(5, 2))
        self.description_text = ctk.CTkTextbox(self, width=form_width, height=55)
        self.description_text.pack(pady=(0, 8))

        ctk.CTkButton(self, text="Сохранить", command=self.save, width=230, height=34, corner_radius=10).pack(pady=(8, 12))

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

        self.status_combo.set(self.task.status or TASK_STATUS_NEW)
        self.description_text.insert("1.0", self.task.description or "")

    def save(self):
        due_date = self.parse_date(self.due_date_entry.get())
        if due_date == "INVALID":
            self.show_error("Срок должен быть в формате дд.мм.гггг")
            return

        payload = {
            "person_id": self.parse_person_id(),
            "title": self.title_entry.get().strip(),
            "description": self.description_text.get("1.0", "end").strip(),
            "main_responsible": self.main_responsible_combo.get().strip(),
            "co_executors": self.co_exec_entry.get().strip(),
            "controller": self.controller_entry.get().strip(),
            "due_date": due_date,
            "status": self.status_combo.get().strip(),
        }

        try:
            if self.mode == "edit":
                task = self.parent_tab.task_service.update_task(self.task_id, payload)
                if not task:
                    self.show_error("Задача не найдена")
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
        win.title("Ошибка")
        win.grab_set()
        ctk.CTkLabel(win, text=text, wraplength=380).pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy, width=120, height=34, corner_radius=10).pack()
