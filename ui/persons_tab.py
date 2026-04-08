import customtkinter as ctk
from tkinter import ttk, filedialog
from tkcalendar import Calendar
from datetime import datetime
import tkinter as tk

from database import get_session
from models.person import Person
from models.reference import Responsible, Circle
from services.people_import import (
    create_people_import_template,
    import_people_from_excel,
)
from services.sync_to_birthday import sync_and_launch_birthday_reminder


YES_NO_VALUES = ["Да", "Нет"]


class PersonsTab(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)

        self.selected_person_id = None
        self.only_without_birthday_var = tk.BooleanVar(value=False)

        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=(10, 6))

        ctk.CTkButton(top_frame, text="➕ Новый", command=self.open_add_window).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="✏️ Редактировать", command=self.open_edit_window).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="🗑 Удалить", command=self.delete_selected).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="📥 Импорт из Excel", command=self.import_people).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="📄 Шаблон импорта", command=self.download_import_template).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="🎂 Экспорт в ДР", command=self.export_birthdays).pack(side="left", padx=5)
        ctk.CTkButton(top_frame, text="🔄 Обновить", command=self.refresh).pack(side="right", padx=5)

        filter_frame = ctk.CTkFrame(self)
        filter_frame.pack(fill="x", padx=10, pady=(0, 8))

        self.search_label = ctk.CTkLabel(filter_frame, text="Поиск по ФИО")
        self.search_label.pack(side="left", padx=(10, 5), pady=10)

        self.search_entry = ctk.CTkEntry(filter_frame, width=280, placeholder_text="Введите ФИО...")
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

        self.only_without_birthday_checkbox = ctk.CTkCheckBox(
            filter_frame,
            text="Отображать только без даты рождения",
            variable=self.only_without_birthday_var,
            command=self.refresh
        )
        self.only_without_birthday_checkbox.pack(side="left", padx=(15, 5), pady=10)

        ctk.CTkButton(
            filter_frame,
            text="Сбросить поиск",
            width=150,
            command=self.reset_search
        ).pack(side="right", padx=10, pady=10)

        table_frame = ctk.CTkFrame(self)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("id", "fio", "position", "circle", "level", "phone", "track_calls", "responsible", "birthday")

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        self.tree.heading("id", text="ID")
        self.tree.heading("fio", text="ФИО")
        self.tree.heading("position", text="Должность")
        self.tree.heading("circle", text="Круг")
        self.tree.heading("level", text="Уровень")
        self.tree.heading("phone", text="Телефон/Контакт")
        self.tree.heading("track_calls", text="Отслеживать")
        self.tree.heading("responsible", text="Ответственный")
        self.tree.heading("birthday", text="ДР")

        self.tree.column("id", width=60, anchor="center")
        self.tree.column("fio", width=240)
        self.tree.column("position", width=220)
        self.tree.column("circle", width=80, anchor="center")
        self.tree.column("level", width=120)
        self.tree.column("phone", width=180)
        self.tree.column("track_calls", width=110, anchor="center")
        self.tree.column("responsible", width=160)
        self.tree.column("birthday", width=100, anchor="center")

        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_double_click)

        self.refresh()

    def reset_search(self):
        self.search_entry.delete(0, "end")
        self.only_without_birthday_var.set(False)
        self.refresh()

    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        search_text = self.search_entry.get().strip().lower()
        only_without_birthday = self.only_without_birthday_var.get()

        session = get_session()
        persons = session.query(Person).order_by(Person.id).all()

        for p in persons:
            fio_value = p.fio or ""

            if search_text and search_text not in fio_value.lower():
                continue

            if only_without_birthday and p.birthday:
                continue

            bday = p.birthday.strftime("%d.%m.%Y") if p.birthday else ""
            phone_value = p.phone_contact_person or p.phone or ""

            self.tree.insert(
                "",
                "end",
                values=(
                    p.id,
                    fio_value,
                    p.position or "",
                    p.circle or "",
                    p.level or "",
                    phone_value,
                    p.track_calls or "",
                    p.responsible or "",
                    bday,
                )
            )

        session.close()
        self.selected_person_id = None

    def on_select(self, event):
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0])["values"]
            self.selected_person_id = values[0]

    def on_double_click(self, event):
        self.open_edit_window()

    def open_add_window(self):
        PersonWindow(self, mode="add")

    def open_edit_window(self):
        if not self.selected_person_id:
            self.show_message("Выбери человека")
            return
        PersonWindow(self, mode="edit", person_id=self.selected_person_id)

    def delete_selected(self):
        if not self.selected_person_id:
            self.show_message("Выбери человека")
            return

        session = get_session()
        person = session.query(Person).filter(Person.id == self.selected_person_id).first()
        if person:
            session.delete(person)
            session.commit()
        session.close()

        self.refresh()

    def download_import_template(self):
        try:
            output_path = create_people_import_template()
            self.show_message(f"Шаблон импорта создан:\n{output_path}")
        except Exception as e:
            self.show_message(f"Ошибка создания шаблона:\n{e}")

    def import_people(self):
        file_path = filedialog.askopenfilename(
            title="Выбери Excel-файл для импорта людей",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            result = import_people_from_excel(file_path)
            self.refresh()
            self.show_message(
                "Импорт завершён.\n\n"
                f"Создано людей: {result['created']}\n"
                f"Обновлено людей: {result['updated']}\n"
                f"Пропущено пустых строк: {result['skipped']}\n"
                f"Создано контактов из legacy-полей: {result['interactions_created']}"
            )
        except Exception as e:
            self.show_message(f"Ошибка импорта:\n{e}")

    def export_birthdays(self):
        try:
            result = sync_and_launch_birthday_reminder()

            message = (
                f"Синхронизация с Birthday Reminder завершена.\n\n"
                f"Добавлено новых: {result['added_count']}\n"
                f"Обновлено существующих: {result['updated_count']}\n"
                f"Без изменений: {result['unchanged_count']}\n"
                f"Пропущено без даты рождения: {result['skipped_empty_birth_date_count']}\n"
                f"Пропущено без ФИО: {result['skipped_empty_name_count']}\n\n"
                f"База Birthday Reminder:\n{result['db_path']}\n\n"
            )

            if result["launched"]:
                message += f"Birthday Reminder запущен:\n{result['launch_info']}"
            else:
                message += f"Приложение не запущено:\n{result['launch_info']}"

            self.show_message(message)

        except Exception as e:
            self.show_message(f"Ошибка экспорта в Birthday Reminder:\n{e}")

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("620x320")
        win.title("Сообщение")
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=570, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack(pady=10)


class PersonWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab, mode="add", person_id=None):
        super().__init__()

        self.parent_tab = parent_tab
        self.mode = mode
        self.person_id = person_id

        self.geometry("780x920")
        self.minsize(780, 920)
        self.title("Новый человек" if mode == "add" else "Редактирование человека")
        self.resizable(True, False)
        self.grab_set()

        self.person = None
        if mode == "edit":
            session = get_session()
            self.person = session.query(Person).filter(Person.id == person_id).first()
            session.close()

        self.build()

    def get_responsibles(self):
        session = get_session()
        data = [x.name for x in session.query(Responsible).order_by(Responsible.name).all()]
        session.close()
        return data if data else [""]

    def get_circles(self):
        session = get_session()
        data = [x.name for x in session.query(Circle).order_by(Circle.name).all()]
        session.close()
        return data if data else [""]

    def add_labeled_entry(self, parent, label_text, attr_name, width=320):
        label = ctk.CTkLabel(parent, text=label_text)
        label.pack(anchor="w", pady=(8, 2))
        entry = ctk.CTkEntry(parent, width=width)
        entry.pack(anchor="w")
        setattr(self, attr_name, entry)

    def add_labeled_textbox(self, parent, label_text, attr_name, width=640, height=80):
        label = ctk.CTkLabel(parent, text=label_text)
        label.pack(anchor="w", pady=(8, 2))
        textbox = ctk.CTkTextbox(parent, width=width, height=height)
        textbox.pack(anchor="w")
        setattr(self, attr_name, textbox)

    def open_calendar(self):
        win = ctk.CTkToplevel(self)
        win.geometry("320x360")
        win.title("Выбор даты")
        win.resizable(False, False)
        win.grab_set()

        cal = Calendar(win, date_pattern="dd.mm.yyyy")
        cal.pack(expand=True, fill="both", padx=10, pady=10)

        def apply_date():
            self.birthday_entry.delete(0, "end")
            self.birthday_entry.insert(0, cal.get_date())
            win.destroy()

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(pady=(0, 10))

        ctk.CTkButton(btn_frame, text="OK", width=100, command=apply_date).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Отмена", width=100, command=win.destroy).pack(side="left", padx=5)

    def build(self):
        container = ctk.CTkScrollableFrame(self, width=720, height=840)
        container.pack(fill="both", expand=True, padx=15, pady=15)

        self.add_labeled_entry(container, "ФИО", "fio_entry", width=640)
        self.add_labeled_entry(container, "Должность", "position_entry", width=640)
        self.add_labeled_entry(container, "Подразделение", "department_entry", width=640)
        self.add_labeled_entry(container, "Уровень", "level_entry", width=320)

        circle_label = ctk.CTkLabel(container, text="Круг общения")
        circle_label.pack(anchor="w", pady=(8, 2))
        self.circle_combo = ctk.CTkComboBox(container, values=self.get_circles(), width=320)
        self.circle_combo.pack(anchor="w")

        responsible_label = ctk.CTkLabel(container, text="Ответственный")
        responsible_label.pack(anchor="w", pady=(8, 2))
        self.responsible_combo = ctk.CTkComboBox(container, values=self.get_responsibles(), width=320)
        self.responsible_combo.pack(anchor="w")

        self.add_labeled_entry(container, "Телефон", "phone_entry", width=320)
        self.add_labeled_entry(container, "Контактный телефон", "phone_contact_person_entry", width=320)

        track_calls_label = ctk.CTkLabel(container, text="Отслеживать звонки?")
        track_calls_label.pack(anchor="w", pady=(8, 2))
        self.track_calls_combo = ctk.CTkComboBox(container, values=YES_NO_VALUES, width=160)
        self.track_calls_combo.pack(anchor="w")
        self.track_calls_combo.set("Да")

        birthday_label = ctk.CTkLabel(container, text="Дата рождения")
        birthday_label.pack(anchor="w", pady=(8, 2))

        birthday_frame = ctk.CTkFrame(container, fg_color="transparent")
        birthday_frame.pack(anchor="w")

        self.birthday_entry = ctk.CTkEntry(birthday_frame, width=200, placeholder_text="дд.мм.гггг")
        self.birthday_entry.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            birthday_frame,
            text="📅 Выбрать",
            width=120,
            command=self.open_calendar
        ).pack(side="left")

        self.add_labeled_textbox(container, "Комментарий", "comment_text", width=640, height=100)

        ctk.CTkButton(container, text="Сохранить", command=self.save, width=220).pack(pady=20)

        if self.person:
            self.fill_form()

    def fill_form(self):
        self.fio_entry.insert(0, self.person.fio or "")
        self.position_entry.insert(0, self.person.position or "")
        self.department_entry.insert(0, self.person.department or "")
        self.level_entry.insert(0, self.person.level or "")
        self.circle_combo.set(self.person.circle or "")
        self.responsible_combo.set(self.person.responsible or "")
        self.phone_entry.insert(0, self.person.phone or "")
        self.phone_contact_person_entry.insert(0, self.person.phone_contact_person or "")
        self.track_calls_combo.set(self.person.track_calls or "Да")

        if self.person.birthday:
            self.birthday_entry.insert(0, self.person.birthday.strftime("%d.%m.%Y"))

        self.comment_text.insert("1.0", self.person.comment or "")

    def parse_date(self, value: str):
        value = (value or "").strip()
        if not value:
            return None

        date_formats = [
            "%d.%m.%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%d-%m-%Y",
        ]

        for fmt in date_formats:
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue

        return "INVALID"

    def save(self):
        fio = self.fio_entry.get().strip()
        if not fio:
            self.show_error("ФИО обязательно")
            return

        birthday = self.parse_date(self.birthday_entry.get())
        if birthday == "INVALID":
            self.show_error("Дата рождения должна быть в корректном формате")
            return

        session = get_session()

        if self.mode == "edit":
            person = session.query(Person).filter(Person.id == self.person_id).first()
            if not person:
                session.close()
                self.show_error("Человек не найден")
                return
        else:
            person = Person()

        person.fio = fio
        person.position = self.position_entry.get().strip()
        person.department = self.department_entry.get().strip()
        person.level = self.level_entry.get().strip()
        person.circle = self.circle_combo.get().strip()
        person.responsible = self.responsible_combo.get().strip()
        person.phone = self.phone_entry.get().strip()
        person.phone_contact_person = self.phone_contact_person_entry.get().strip()
        person.track_calls = self.track_calls_combo.get().strip()
        person.birthday = birthday
        person.comment = self.comment_text.get("1.0", "end").strip()

        if self.mode == "add":
            session.add(person)

        session.commit()
        session.close()

        self.parent_tab.refresh()
        self.destroy()

    def show_error(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("420x160")
        win.title("Ошибка")
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=380).pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack()