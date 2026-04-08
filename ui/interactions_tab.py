import customtkinter as ctk
from collections import defaultdict
from tkinter import ttk
from tkcalendar import Calendar
from datetime import datetime, timedelta

from database import get_session
from models.interaction import Interaction
from models.person import Person
from models.reference import Responsible, Circle
from services.ics_export import export_interactions_to_ics
from services.overdue_tasks_generator import OverdueTasksGenerator
from ui.statuses import (
    STATUS_ALL,
    STATUS_OVERDUE,
    STATUS_TODAY,
    STATUS_7_DAYS,
    STATUS_PLANNED,
    STATUS_NO_DATE,
    INTERACTION_STATUS_VALUES,
)


INTERACTION_TYPES = ["Звонок", "Встреча"]
TYPE_FILTER_VALUES = ["Все", "Звонок", "Встреча"]
RESPONSIBLE_FILTER_DEFAULT = "Все"


class InteractionsTab(ctk.CTkFrame):
    def __init__(self, parent, interaction_service, settings_service):
        super().__init__(parent)

        self.interaction_service = interaction_service
        self.settings_service = settings_service
        self.selected_interaction_id = None
        self.overdue_tasks_generator = OverdueTasksGenerator()

        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkButton(top_frame, text="➕ Новый контакт", command=self.open_add_window).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="✏️ Редактировать", command=self.open_edit_window).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="🗑 Удалить", command=self.delete_selected).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="📝 Создать задачи из просрочки", command=self.create_overdue_tasks, width=220).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="📋 Сводка", command=self.generate_summary).pack(side="right", padx=5)
        ctk.CTkButton(top_frame, text="📅 Экспорт в ICS", command=self.export_to_ics).pack(side="right", padx=5)
        ctk.CTkButton(top_frame, text="🔄 Обновить", command=self.refresh).pack(side="right", padx=5)

        filter_frame = ctk.CTkFrame(self)
        filter_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.search_label = ctk.CTkLabel(filter_frame, text="Поиск по ФИО")
        self.search_label.pack(side="left", padx=(10, 5), pady=10)

        self.search_entry = ctk.CTkEntry(filter_frame, width=200, placeholder_text="Введите ФИО...")
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

        self.type_filter_label = ctk.CTkLabel(filter_frame, text="Тип")
        self.type_filter_label.pack(side="left", padx=(15, 5), pady=10)

        self.type_filter = ctk.CTkComboBox(
            filter_frame,
            values=TYPE_FILTER_VALUES,
            width=120,
            command=lambda _: self.refresh()
        )
        self.type_filter.pack(side="left", padx=5, pady=10)
        self.type_filter.set("Все")

        self.status_filter_label = ctk.CTkLabel(filter_frame, text="Статус")
        self.status_filter_label.pack(side="left", padx=(15, 5), pady=10)

        self.status_filter = ctk.CTkComboBox(
            filter_frame,
            values=INTERACTION_STATUS_VALUES,
            width=150,
            command=lambda _: self.refresh()
        )
        self.status_filter.pack(side="left", padx=5, pady=10)
        self.status_filter.set(STATUS_ALL)

        self.responsible_filter_label = ctk.CTkLabel(filter_frame, text="Ответственный")
        self.responsible_filter_label.pack(side="left", padx=(15, 5), pady=10)

        self.responsible_filter = ctk.CTkComboBox(
            filter_frame,
            values=self.get_responsible_filter_values(),
            width=170,
            command=lambda _: self.refresh()
        )
        self.responsible_filter.pack(side="left", padx=5, pady=10)
        self.responsible_filter.set(RESPONSIBLE_FILTER_DEFAULT)

        ctk.CTkButton(
            filter_frame,
            text="Сбросить фильтры",
            width=150,
            command=self.reset_filters
        ).pack(side="right", padx=10, pady=10)

        self.dashboard_frame = ctk.CTkFrame(self)
        self.dashboard_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.overdue_label = ctk.CTkLabel(self.dashboard_frame, text="🔴 Просрочено: 0")
        self.overdue_label.pack(side="left", padx=15, pady=10)

        self.today_label = ctk.CTkLabel(self.dashboard_frame, text="🟡 Сегодня: 0")
        self.today_label.pack(side="left", padx=15, pady=10)

        self.week_label = ctk.CTkLabel(self.dashboard_frame, text="🟢 7 дней: 0")
        self.week_label.pack(side="left", padx=15, pady=10)

        self.planned_label = ctk.CTkLabel(self.dashboard_frame, text="⚪ Запланировано: 0")
        self.planned_label.pack(side="left", padx=15, pady=10)

        self.no_date_label = ctk.CTkLabel(self.dashboard_frame, text="⚫ Без даты: 0")
        self.no_date_label.pack(side="left", padx=15, pady=10)

        table_frame = ctk.CTkFrame(self)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = (
            "id",
            "person",
            "interaction_type",
            "interaction_date",
            "next_date",
            "status",
            "responsible",
            "purpose",
        )

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        self.tree.heading("id", text="ID")
        self.tree.heading("person", text="Человек")
        self.tree.heading("interaction_type", text="Тип")
        self.tree.heading("interaction_date", text="Дата")
        self.tree.heading("next_date", text="Следующий контакт")
        self.tree.heading("status", text="Статус")
        self.tree.heading("responsible", text="Ответственный")
        self.tree.heading("purpose", text="Цель")

        self.tree.column("id", width=60, anchor="center")
        self.tree.column("person", width=220)
        self.tree.column("interaction_type", width=100, anchor="center")
        self.tree.column("interaction_date", width=100, anchor="center")
        self.tree.column("next_date", width=140, anchor="center")
        self.tree.column("status", width=130, anchor="center")
        self.tree.column("responsible", width=160)
        self.tree.column("purpose", width=320)

        self.tree.tag_configure("overdue", background="#ffdddd")
        self.tree.tag_configure("today", background="#fff4cc")
        self.tree.tag_configure("week", background="#ddffdd")
        self.tree.tag_configure("planned", background="#f7f7f7")
        self.tree.tag_configure("nodate", background="#eeeeee")

        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_double_click)

        self.refresh()

    def get_responsible_filter_values(self):
        session = get_session()
        values = {RESPONSIBLE_FILTER_DEFAULT}

        for item in session.query(Responsible).order_by(Responsible.name).all():
            if (item.name or "").strip():
                values.add(item.name.strip())

        for item in session.query(Person).all():
            if (item.responsible or "").strip():
                values.add(item.responsible.strip())

        session.close()
        ordered = [RESPONSIBLE_FILTER_DEFAULT] + sorted(v for v in values if v != RESPONSIBLE_FILTER_DEFAULT)
        return ordered

    def reset_filters(self):
        self.search_entry.delete(0, "end")
        self.type_filter.set("Все")
        self.status_filter.set(STATUS_ALL)
        self.responsible_filter.configure(values=self.get_responsible_filter_values())
        self.responsible_filter.set("Все")
        self.refresh()

    def get_row_tag_by_status(self, status: str) -> str:
        if status == STATUS_OVERDUE:
            return "overdue"
        if status == STATUS_TODAY:
            return "today"
        if status == STATUS_7_DAYS:
            return "week"
        if status == STATUS_NO_DATE:
            return "nodate"
        return "planned"

    def get_active_interactions(self):
        return self.interaction_service.list_active_interactions(
            search=self.search_entry.get().strip(),
            status=self.status_filter.get().strip(),
            type_=self.type_filter.get().strip(),
            responsible=self.responsible_filter.get().strip(),
        )

    def refresh(self):
        responsible_values = self.get_responsible_filter_values()
        self.responsible_filter.configure(values=responsible_values)

        current_responsible = self.responsible_filter.get().strip()
        if current_responsible not in responsible_values:
            self.responsible_filter.set("Все")

        for row in self.tree.get_children():
            self.tree.delete(row)

        all_rows = self.interaction_service.list_active_interactions()
        overdue_count = len([row for row in all_rows if row['status'] == STATUS_OVERDUE])
        today_count = len([row for row in all_rows if row['status'] == STATUS_TODAY])
        week_count = len([row for row in all_rows if row['status'] == STATUS_7_DAYS])
        planned_count = len([row for row in all_rows if row['status'] == STATUS_PLANNED])
        no_date_count = len([row for row in all_rows if row['status'] == STATUS_NO_DATE])

        interactions = self.get_active_interactions()

        for row in interactions:
            interaction = row['interaction']
            person = row['person']
            status = row['status']
            tag = self.get_row_tag_by_status(status)
            interaction_date = interaction.interaction_date.strftime("%d.%m.%Y") if interaction.interaction_date else ""
            next_date = interaction.next_date.strftime("%d.%m.%Y") if interaction.next_date else ""

            self.tree.insert(
                "",
                "end",
                values=(
                    interaction.id,
                    (person.fio if person else f"[ID {interaction.person_id}]") or f"[ID {interaction.person_id}]",
                    interaction.interaction_type or "",
                    interaction_date,
                    next_date,
                    status,
                    row['responsible'],
                    (interaction.purpose or "")[:80],
                ),
                tags=(tag,)
            )

        self.overdue_label.configure(text=f"🔴 Просрочено: {overdue_count}")
        self.today_label.configure(text=f"🟡 Сегодня: {today_count}")
        self.week_label.configure(text=f"🟢 7 дней: {week_count}")
        self.planned_label.configure(text=f"⚪ Запланировано: {planned_count}")
        self.no_date_label.configure(text=f"⚫ Без даты: {no_date_count}")

        self.selected_interaction_id = None

    def on_select(self, event):
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0])["values"]
            self.selected_interaction_id = values[0]

    def on_double_click(self, event):
        self.open_edit_window()

    def open_add_window(self):
        InteractionWindow(self, mode="add")

    def open_edit_window(self):
        if not self.selected_interaction_id:
            self.show_message("Выбери контакт")
            return
        InteractionWindow(self, mode="edit", interaction_id=self.selected_interaction_id)

    def delete_selected(self):
        if not self.selected_interaction_id:
            self.show_message("Выбери контакт")
            return

        deleted = self.interaction_service.delete_interaction(
            self.selected_interaction_id,
            meeting_equals_call=self.settings_service.get_bool("meeting_equals_call", default=True),
        )
        if not deleted:
            self.show_message("Контакт не найден")
            return

        self.refresh()

    def create_overdue_tasks(self):
        interactions = self.interaction_service.list_active_interactions()
        overdue_rows = [row for row in interactions if row["status"] == STATUS_OVERDUE]

        if not overdue_rows:
            self.show_message("Просроченных контактов нет")
            return

        try:
            groups = self.overdue_tasks_generator.build_groups(overdue_rows)
            result = self.overdue_tasks_generator.create_tasks_from_groups(groups)
        except Exception as e:
            self.show_message(f"Ошибка создания задач из просрочки:\n{e}")
            return

        text_lines = [
            "Задачи из просроченных контактов обработаны.",
            "",
            f"Создано задач: {result['created']}",
            f"Пропущено дублей: {result['skipped']}",
        ]
        if result.get("details"):
            text_lines.append("")
            text_lines.extend(result["details"])

        self.show_message("\n".join(text_lines))

    def generate_summary(self):
        interactions = self.interaction_service.list_active_interactions()

        grouped_overdue = defaultdict(list)
        total_overdue = 0

        for row in interactions:
            if row["status"] != STATUS_OVERDUE:
                continue

            interaction = row["interaction"]
            person = row["person"]
            person_name = person.fio if person else f"[ID {interaction.person_id}]"
            responsible = (row["responsible"] or "Без ответственного").strip()
            grouped_overdue[responsible].append(person_name)
            total_overdue += 1

        summary_blocks = []
        summary_blocks.append("Добрый день! Сводка по просроченным контактам.\n")
        summary_blocks.append(f"Всего просроченных контактов: {total_overdue}\n")

        if not grouped_overdue:
            summary_blocks.append("Просроченные контакты отсутствуют.")
        else:
            for responsible in sorted(grouped_overdue.keys()):
                names = grouped_overdue[responsible]
                summary_blocks.append(f"{responsible}: {len(names)}")
                for fio in names:
                    summary_blocks.append(f"— {fio}")
                summary_blocks.append("")

        full_text = "\n".join(summary_blocks).strip()

        responsible_texts = {}
        for responsible in sorted(grouped_overdue.keys()):
            names = grouped_overdue[responsible]
            lines = [
                f"Здравствуйте! Просьба обратить внимание на просроченные контакты.",
                "",
                f"По вам просроченных контактов: {len(names)}",
                "",
            ]
            for fio in names:
                lines.append(f"— {fio}")
            responsible_texts[responsible] = "\n".join(lines).strip()

        self.show_summary_window(full_text, responsible_texts)

    def show_summary_window(self, text, responsible_texts=None):
        responsible_texts = responsible_texts or {}

        win = ctk.CTkToplevel(self)
        win.geometry("860x620")
        win.title("Сводка")
        win.grab_set()

        textbox = ctk.CTkTextbox(win, width=800, height=460)
        textbox.pack(padx=20, pady=(20, 12))
        textbox.insert("1.0", text)

        buttons_frame = ctk.CTkFrame(win, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=20, pady=(0, 14))

        def copy_text(copy_value, success_text):
            win.clipboard_clear()
            win.clipboard_append(copy_value)
            self.show_message(success_text)

        ctk.CTkButton(
            buttons_frame,
            text="📋 Копировать всю сводку",
            command=lambda: copy_text(text, "Сводка скопирована в буфер обмена"),
            width=200
        ).pack(side="left", padx=(0, 10))

        if responsible_texts:
            responsible_frame = ctk.CTkScrollableFrame(win, height=90)
            responsible_frame.pack(fill="x", padx=20, pady=(0, 16))

            ctk.CTkLabel(
                responsible_frame,
                text="Точечное копирование по исполнителям",
                font=ctk.CTkFont(size=14, weight="bold")
            ).pack(anchor="w", padx=4, pady=(2, 8))

            for responsible in sorted(responsible_texts.keys()):
                ctk.CTkButton(
                    responsible_frame,
                    text=f"📋 {responsible}",
                    command=lambda r=responsible: copy_text(
                        responsible_texts[r],
                        f"Список для {r} скопирован в буфер обмена"
                    ),
                    width=220
                ).pack(side="left", padx=4, pady=(0, 8))

    def export_to_ics(self):
        rows = self.interaction_service.list_active_interactions()
        prepared = []
        for row in rows:
            interaction = row["interaction"]
            person = row["person"]
            if not interaction.next_date:
                continue
            prepared.append((interaction, person, row['status']))

        if not prepared:
            self.show_message("Нет активных контактов с датой следующего контакта для экспорта в ICS")
            return

        try:
            output_path = export_interactions_to_ics(prepared)
            self.show_message(f"Экспорт в ICS завершён:\n{output_path}")
        except Exception as e:
            self.show_message(f"Ошибка экспорта в ICS:\n{e}")

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("460x180")
        win.title("Сообщение")
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=420, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack(pady=10)




class InteractionWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab, mode="add", interaction_id=None):
        super().__init__()

        self.parent_tab = parent_tab
        self.mode = mode
        self.interaction_id = interaction_id

        self.title("Новый контакт" if mode == "add" else "Редактирование контакта")
        self.geometry("980x980")
        self.minsize(920, 760)
        self.resizable(True, True)
        self.grab_set()

        self.interaction = None
        if mode == "edit":
            self.interaction = self.parent_tab.interaction_service.get_interaction(interaction_id)

        self.person_records = self.get_person_records()
        self.selected_person_id = None
        self.selected_person_fio = ""
        self.displayed_person_matches = []
        self.person_dropdown_visible = False

        self.scrollable_form = None
        self.build()

    def get_person_records(self):
        session = get_session()
        items = session.query(Person).order_by(Person.fio).all()
        records = []
        for item in items:
            if not self.parent_tab.interaction_service.is_track_calls_enabled(item.track_calls):
                continue
            fio = (item.fio or "").strip()
            if not fio:
                continue
            records.append({
                "id": item.id,
                "fio": fio,
                "search": f"{item.id} {fio}".lower(),
            })
        session.close()
        return records

    def get_responsibles(self):
        session = get_session()
        values = set()

        for item in session.query(Responsible).order_by(Responsible.name).all():
            if (item.name or "").strip():
                values.add(item.name.strip())

        for person in session.query(Person).all():
            if (person.responsible or "").strip():
                values.add(person.responsible.strip())

        session.close()
        ordered = sorted(values)
        return ordered if ordered else [""]

    def _enable_mousewheel_scroll(self, widget):
        canvas = getattr(widget, "_parent_canvas", None)
        if canvas is None:
            return

        def _on_mousewheel(event):
            try:
                if getattr(event, "delta", 0):
                    raw = int(event.delta / 120) if event.delta else 0
                    if raw == 0:
                        raw = 1 if event.delta > 0 else -1
                    steps = max(1, abs(raw)) * 5
                    canvas.yview_scroll(-steps if raw > 0 else steps, "units")
                    return "break"
                if getattr(event, "num", None) == 4:
                    canvas.yview_scroll(-4, "units")
                    return "break"
                if getattr(event, "num", None) == 5:
                    canvas.yview_scroll(4, "units")
                    return "break"
            except Exception:
                return None
            return "break"

        def _bind_recursive(target):
            try:
                target.bind("<MouseWheel>", _on_mousewheel, add="+")
                target.bind("<Button-4>", _on_mousewheel, add="+")
                target.bind("<Button-5>", _on_mousewheel, add="+")
            except Exception:
                pass
            for child in target.winfo_children():
                _bind_recursive(child)

        def _install():
            _bind_recursive(widget)
            try:
                _bind_recursive(canvas)
            except Exception:
                pass

        self.after(50, _install)

    def build(self):
        form_width = 760
        self.form_width = form_width

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.scrollable_form = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scrollable_form.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        self._enable_mousewheel_scroll(self.scrollable_form)
        self.after(150, lambda: self._enable_mousewheel_scroll(self.scrollable_form))
        content = self.scrollable_form

        self.person_label = ctk.CTkLabel(content, text="Человек")
        self.person_label.pack(pady=(15, 4))

        self.person_entry = ctk.CTkEntry(
            content,
            width=form_width,
            placeholder_text="Начните вводить фамилию..."
        )
        self.person_entry.pack(pady=(0, 6))
        self.person_entry.bind("<KeyRelease>", self.on_person_entry_keyrelease)
        self.person_entry.bind("<FocusIn>", self.on_person_entry_focus_in)
        self.person_entry.bind("<Return>", self.on_person_entry_return)

        self.person_dropdown_host = ctk.CTkFrame(content, fg_color="transparent", width=form_width, height=140)
        self.person_dropdown_host.pack(pady=(0, 10))
        self.person_dropdown_host.pack_propagate(False)

        self.person_dropdown = ctk.CTkScrollableFrame(
            self.person_dropdown_host,
            width=form_width,
            height=140,
            fg_color=("#F2F2F2", "#2B2B2B")
        )
        self.person_dropdown.pack_forget()
        self._enable_mousewheel_scroll(self.person_dropdown)

        self.type_label = ctk.CTkLabel(content, text="Тип контакта")
        self.type_label.pack(pady=(6, 2))
        self.type_combo = ctk.CTkComboBox(content, values=INTERACTION_TYPES, width=form_width)
        self.type_combo.pack(pady=(0, 8))
        self.type_combo.set(INTERACTION_TYPES[0])

        self.responsible_label = ctk.CTkLabel(content, text="Ответственный")
        self.responsible_label.pack(pady=(6, 2))
        self.responsible_combo = ctk.CTkComboBox(content, values=self.get_responsibles(), width=form_width)
        self.responsible_combo.pack(pady=(0, 8))

        self.interaction_date_label = ctk.CTkLabel(content, text="Дата контакта")
        self.interaction_date_label.pack(pady=(6, 2))
        self.interaction_date_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.interaction_date_frame.pack(pady=(0, 8))

        self.interaction_date_entry = ctk.CTkEntry(
            self.interaction_date_frame,
            width=430,
            placeholder_text="дд.мм.гггг"
        )
        self.interaction_date_entry.pack(side="left", padx=5)
        self.interaction_date_entry.bind("<FocusOut>", lambda event: self.autofill_next_date())

        ctk.CTkButton(
            self.interaction_date_frame,
            text="📅 Выбрать",
            width=120,
            command=lambda: self.open_calendar(self.interaction_date_entry, after_apply=self.autofill_next_date)
        ).pack(side="left", padx=5)

        self.next_date_label = ctk.CTkLabel(content, text="Дата следующего контакта")
        self.next_date_label.pack(pady=(6, 2))
        self.next_date_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.next_date_frame.pack(pady=(0, 8))

        self.next_date_entry = ctk.CTkEntry(
            self.next_date_frame,
            width=430,
            placeholder_text="дд.мм.гггг"
        )
        self.next_date_entry.pack(side="left", padx=5)

        ctk.CTkButton(
            self.next_date_frame,
            text="📅 Выбрать",
            width=120,
            command=lambda: self.open_calendar(self.next_date_entry)
        ).pack(side="left", padx=5)

        self.purpose_label = ctk.CTkLabel(content, text="Цель контакта")
        self.purpose_label.pack(pady=(8, 2))
        self.purpose_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.purpose_text.pack(pady=(0, 8))

        self.result_label = ctk.CTkLabel(content, text="Результат")
        self.result_label.pack(pady=(8, 2))
        self.result_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.result_text.pack(pady=(0, 8))

        self.comment_label = ctk.CTkLabel(content, text="Комментарий")
        self.comment_label.pack(pady=(8, 2))
        self.comment_text = ctk.CTkTextbox(content, width=form_width, height=90)
        self.comment_text.pack(pady=(0, 8))

        self.save_button = ctk.CTkButton(
            content,
            text="Сохранить",
            command=self.save,
            width=220
        )
        self.save_button.pack(pady=(15, 20))

        if self.interaction:
            self.fill()

    def match_person_records(self, query: str):
        query = (query or "").strip().lower()
        if not query:
            return self.person_records[:12]
        return [item for item in self.person_records if query in item["search"]][:12]

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
        self.on_person_changed()

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

    def on_person_changed(self):
        self.prefill_responsible_from_person()
        self.autofill_next_date()

    def prefill_responsible_from_person(self):
        person_id = self.parse_person_id(self.person_entry.get())
        if not person_id:
            return

        session = get_session()
        person = session.get(Person, person_id)
        session.close()

        if person and (person.responsible or "").strip():
            self.responsible_combo.set(person.responsible.strip())

    def fill(self):
        session = get_session()
        person = session.get(Person, self.interaction.person_id)
        session.close()

        if person and self.parent_tab.interaction_service.is_track_calls_enabled(person.track_calls):
            self.selected_person_id = person.id
            self.selected_person_fio = person.fio or ""
            self.person_entry.delete(0, "end")
            self.person_entry.insert(0, self.selected_person_fio)

        effective_responsible = self.interaction.responsible or (person.responsible if person else "") or ""
        self.type_combo.set(self.interaction.interaction_type or INTERACTION_TYPES[0])
        self.responsible_combo.set(effective_responsible)

        if self.interaction.interaction_date:
            self.interaction_date_entry.insert(0, self.interaction.interaction_date.strftime("%d.%m.%Y"))

        if self.interaction.next_date:
            self.next_date_entry.insert(0, self.interaction.next_date.strftime("%d.%m.%Y"))

        self.purpose_text.insert("1.0", self.interaction.purpose or "")
        self.result_text.insert("1.0", self.interaction.result or "")
        self.comment_text.insert("1.0", self.interaction.comment or "")

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

        ctk.CTkButton(btn_frame, text="OK", width=100, command=apply).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Отмена", width=100, command=win.destroy).pack(side="left", padx=5)

    def parse_person_id(self, current_text: str):
        if self.selected_person_id and (current_text or "").strip() == self.selected_person_fio:
            return self.selected_person_id

        raw = (current_text or "").strip()
        if not raw:
            return None

        if "|" in raw:
            try:
                return int(raw.split("|", 1)[0].strip())
            except (ValueError, IndexError):
                pass

        exact = [item for item in self.person_records if item["fio"].strip().lower() == raw.lower()]
        if len(exact) == 1:
            self.selected_person_id = exact[0]["id"]
            self.selected_person_fio = exact[0]["fio"]
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

    def get_circle_period_days_for_selected_person(self):
        person_id = self.parse_person_id(self.person_entry.get())
        if not person_id:
            return None

        session = get_session()
        person = session.get(Person, person_id)

        if not person or not person.circle:
            session.close()
            return None

        circle = session.query(Circle).filter(Circle.name == person.circle).first()
        session.close()

        if not circle:
            return None

        return circle.contact_period_days or None

    def autofill_next_date(self):
        interaction_date = self.parse_date(self.interaction_date_entry.get())
        if interaction_date in (None, "INVALID"):
            return

        period_days = self.get_circle_period_days_for_selected_person()
        if not period_days:
            return

        next_date = interaction_date + timedelta(days=period_days)

        self.next_date_entry.delete(0, "end")
        self.next_date_entry.insert(0, next_date.strftime("%d.%m.%Y"))

    def save(self):
        person_id = self.parse_person_id(self.person_entry.get())
        if not person_id:
            self.show_error("Нужно выбрать человека из списка")
            return

        interaction_date = self.parse_date(self.interaction_date_entry.get())
        if interaction_date == "INVALID":
            self.show_error("Дата контакта должна быть в формате дд.мм.гггг")
            return

        next_date = self.parse_date(self.next_date_entry.get())
        if next_date == "INVALID":
            self.show_error("Дата следующего контакта должна быть в формате дд.мм.гггг")
            return

        payload = {
            'person_id': person_id,
            'interaction_type': self.type_combo.get().strip(),
            'interaction_date': interaction_date,
            'next_date': next_date,
            'responsible': self.responsible_combo.get().strip(),
            'purpose': self.purpose_text.get("1.0", "end").strip(),
            'result': self.result_text.get("1.0", "end").strip(),
            'comment': self.comment_text.get("1.0", "end").strip(),
        }

        try:
            if self.mode == "edit":
                interaction = self.parent_tab.interaction_service.update_interaction(self.interaction_id, payload)
                if not interaction:
                    self.show_error("Контакт не найден")
                    return
            else:
                self.parent_tab.interaction_service.create_interaction(payload)
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
        ctk.CTkButton(win, text="OK", command=win.destroy).pack()
