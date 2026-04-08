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


YES_NO_VALUES = ["Да", "Нет"]

TRANSLATIONS = {
    "ru": {
        "yes": "Да",
        "no": "Нет",
        "new": "➕ Новый",
        "edit": "✏️ Редактировать",
        "delete": "🗑 Удалить",
        "import_excel": "📥 Импорт из Excel",
        "import_template": "📄 Шаблон импорта",
        "refresh": "🔄 Обновить",
        "search_fio": "Поиск по ФИО",
        "search_placeholder": "Введите ФИО...",
        "only_without_birthday": "Отображать только без даты рождения",
        "reset_search": "Сбросить поиск",
        "col_id": "ID",
        "col_fio": "ФИО",
        "col_position": "Должность",
        "col_circle": "Круг",
        "col_level": "Уровень",
        "col_phone": "Телефон/Контакт",
        "col_track_calls": "Отслеживать",
        "col_responsible": "Ответственный",
        "col_birthday": "ДР",
        "pick_person": "Выбери человека",
        "message_title": "Сообщение",
        "error_title": "Ошибка",
        "import_template_created": "Шаблон импорта создан:\n{path}",
        "import_template_error": "Ошибка создания шаблона:\n{error}",
        "import_pick_file": "Выбери Excel-файл для импорта людей",
        "import_done": "Импорт завершён.\n\nСоздано людей: {created}\nОбновлено людей: {updated}\nПропущено пустых строк: {skipped}\nСоздано контактов из legacy-полей: {interactions_created}",
        "import_error": "Ошибка импорта:\n{error}",
        "person_new_title": "Новый человек",
        "person_edit_title": "Редактирование человека",
        "fio": "ФИО",
        "position": "Должность",
        "department": "Подразделение",
        "level": "Уровень",
        "circle": "Круг общения",
        "responsible": "Ответственный",
        "phone": "Телефон",
        "phone_contact": "Контактный телефон",
        "track_calls": "Отслеживать звонки?",
        "birthday": "Дата рождения",
        "birthday_placeholder": "дд.мм.гггг",
        "pick_date": "📅 Выбрать",
        "comment": "Комментарий",
        "save": "Сохранить",
        "calendar_title": "Выбор даты",
        "cancel": "Отмена",
        "fio_required": "ФИО обязательно",
        "birthday_invalid": "Дата рождения должна быть в корректном формате",
        "person_not_found": "Человек не найден",
        "ok": "OK",
    },
    "en": {
        "yes": "Yes",
        "no": "No",
        "new": "➕ New",
        "edit": "✏️ Edit",
        "delete": "🗑 Delete",
        "import_excel": "📥 Import from Excel",
        "import_template": "📄 Import template",
        "refresh": "🔄 Refresh",
        "search_fio": "Search by name",
        "search_placeholder": "Enter full name...",
        "only_without_birthday": "Show only people without birthday",
        "reset_search": "Reset search",
        "col_id": "ID",
        "col_fio": "Full name",
        "col_position": "Position",
        "col_circle": "Circle",
        "col_level": "Level",
        "col_phone": "Phone/Contact",
        "col_track_calls": "Track",
        "col_responsible": "Responsible",
        "col_birthday": "Birthday",
        "pick_person": "Select a person",
        "message_title": "Message",
        "error_title": "Error",
        "import_template_created": "Import template created:\n{path}",
        "import_template_error": "Failed to create import template:\n{error}",
        "import_pick_file": "Select an Excel file to import people",
        "import_done": "Import completed.\n\nPeople created: {created}\nPeople updated: {updated}\nEmpty rows skipped: {skipped}\nInteractions created from legacy fields: {interactions_created}",
        "import_error": "Import error:\n{error}",
        "person_new_title": "New person",
        "person_edit_title": "Edit person",
        "fio": "Full name",
        "position": "Position",
        "department": "Department",
        "level": "Level",
        "circle": "Circle",
        "responsible": "Responsible",
        "phone": "Phone",
        "phone_contact": "Contact phone",
        "track_calls": "Track calls?",
        "birthday": "Birthday",
        "birthday_placeholder": "dd.mm.yyyy",
        "pick_date": "📅 Pick",
        "comment": "Comment",
        "save": "Save",
        "calendar_title": "Pick a date",
        "cancel": "Cancel",
        "fio_required": "Full name is required",
        "birthday_invalid": "Birthday must be in a valid format",
        "person_not_found": "Person not found",
        "ok": "OK",
    },
}


def normalize_language(value):
    return "en" if str(value or "ru").strip().lower() == "en" else "ru"


class PersonsTab(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)

        self.selected_person_id = None
        self.only_without_birthday_var = tk.BooleanVar(value=False)

        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=10, pady=(10, 6))

        self.new_button = ctk.CTkButton(top_frame, text=self.tr("new"), command=self.open_add_window)
        self.new_button.pack(side="left", padx=5)
        self.edit_button = ctk.CTkButton(top_frame, text=self.tr("edit"), command=self.open_edit_window)
        self.edit_button.pack(side="left", padx=5)
        self.delete_button = ctk.CTkButton(top_frame, text=self.tr("delete"), command=self.delete_selected)
        self.delete_button.pack(side="left", padx=5)
        self.import_button = ctk.CTkButton(top_frame, text=self.tr("import_excel"), command=self.import_people)
        self.import_button.pack(side="left", padx=5)
        self.template_button = ctk.CTkButton(top_frame, text=self.tr("import_template"), command=self.download_import_template)
        self.template_button.pack(side="left", padx=5)
        self.refresh_button = ctk.CTkButton(top_frame, text=self.tr("refresh"), command=self.refresh)
        self.refresh_button.pack(side="right", padx=5)

        filter_frame = ctk.CTkFrame(self)
        filter_frame.pack(fill="x", padx=10, pady=(0, 8))

        self.search_label = ctk.CTkLabel(filter_frame, text=self.tr("search_fio"))
        self.search_label.pack(side="left", padx=(10, 5), pady=10)

        self.search_entry = ctk.CTkEntry(filter_frame, width=280, placeholder_text=self.tr("search_placeholder"))
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

        self.only_without_birthday_checkbox = ctk.CTkCheckBox(
            filter_frame,
            text=self.tr("only_without_birthday"),
            variable=self.only_without_birthday_var,
            command=self.refresh,
        )
        self.only_without_birthday_checkbox.pack(side="left", padx=(15, 5), pady=10)

        self.reset_button = ctk.CTkButton(
            filter_frame,
            text=self.tr("reset_search"),
            width=150,
            command=self.reset_search,
        )
        self.reset_button.pack(side="right", padx=10, pady=10)

        table_frame = ctk.CTkFrame(self)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("id", "fio", "position", "circle", "level", "phone", "track_calls", "responsible", "birthday")

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.tree.heading("id", text=self.tr("col_id"))
        self.tree.heading("fio", text=self.tr("col_fio"))
        self.tree.heading("position", text=self.tr("col_position"))
        self.tree.heading("circle", text=self.tr("col_circle"))
        self.tree.heading("level", text=self.tr("col_level"))
        self.tree.heading("phone", text=self.tr("col_phone"))
        self.tree.heading("track_calls", text=self.tr("col_track_calls"))
        self.tree.heading("responsible", text=self.tr("col_responsible"))
        self.tree.heading("birthday", text=self.tr("col_birthday"))

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

    def get_language(self):
        root = self.winfo_toplevel()
        return normalize_language(getattr(root, "current_language", "ru"))

    def tr(self, key, **kwargs):
        text = TRANSLATIONS[self.get_language()].get(key, key)
        return text.format(**kwargs) if kwargs else text

    def yes_no_values(self):
        return [self.tr("yes"), self.tr("no")]

    def to_display_yes_no(self, value):
        if str(value).strip().lower() == "нет":
            return self.tr("no")
        return self.tr("yes") if str(value).strip() else ""

    def from_display_yes_no(self, value):
        return "Нет" if value == self.tr("no") else "Да"

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
                    self.to_display_yes_no(p.track_calls or ""),
                    p.responsible or "",
                    bday,
                ),
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
            self.show_message(self.tr("pick_person"))
            return
        PersonWindow(self, mode="edit", person_id=self.selected_person_id)

    def delete_selected(self):
        if not self.selected_person_id:
            self.show_message(self.tr("pick_person"))
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
            try:
                output_path = create_people_import_template(language=self.get_language())
            except TypeError:
                output_path = create_people_import_template(self.get_language())
            self.show_message(self.tr("import_template_created", path=output_path))
        except Exception as e:
            self.show_message(self.tr("import_template_error", error=e))

    def import_people(self):
        file_path = filedialog.askopenfilename(
            title=self.tr("import_pick_file"),
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not file_path:
            return

        try:
            result = import_people_from_excel(file_path)
            self.refresh()
            self.show_message(
                self.tr(
                    "import_done",
                    created=result["created"],
                    updated=result["updated"],
                    skipped=result["skipped"],
                    interactions_created=result["interactions_created"],
                )
            )
        except Exception as e:
            self.show_message(self.tr("import_error", error=e))

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("620x320")
        win.title(self.tr("message_title"))
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=570, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text=self.tr("ok"), command=win.destroy).pack(pady=10)


class PersonWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab, mode="add", person_id=None):
        super().__init__()

        self.parent_tab = parent_tab
        self.mode = mode
        self.person_id = person_id

        self.geometry("780x920")
        self.minsize(780, 920)
        self.title(self.tr("person_new_title") if mode == "add" else self.tr("person_edit_title"))
        self.resizable(True, False)
        self.grab_set()

        self.person = None
        if mode == "edit":
            session = get_session()
            self.person = session.query(Person).filter(Person.id == person_id).first()
            session.close()

        self.build()

    def get_language(self):
        return self.parent_tab.get_language()

    def tr(self, key, **kwargs):
        text = TRANSLATIONS[self.get_language()].get(key, key)
        return text.format(**kwargs) if kwargs else text

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
        win.title(self.tr("calendar_title"))
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

        ctk.CTkButton(btn_frame, text=self.tr("ok"), width=100, command=apply_date).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text=self.tr("cancel"), width=100, command=win.destroy).pack(side="left", padx=5)

    def build(self):
        container = ctk.CTkScrollableFrame(self, width=720, height=840)
        container.pack(fill="both", expand=True, padx=15, pady=15)

        self.add_labeled_entry(container, self.tr("fio"), "fio_entry", width=640)
        self.add_labeled_entry(container, self.tr("position"), "position_entry", width=640)
        self.add_labeled_entry(container, self.tr("department"), "department_entry", width=640)
        self.add_labeled_entry(container, self.tr("level"), "level_entry", width=320)

        circle_label = ctk.CTkLabel(container, text=self.tr("circle"))
        circle_label.pack(anchor="w", pady=(8, 2))
        self.circle_combo = ctk.CTkComboBox(container, values=self.get_circles(), width=320)
        self.circle_combo.pack(anchor="w")

        responsible_label = ctk.CTkLabel(container, text=self.tr("responsible"))
        responsible_label.pack(anchor="w", pady=(8, 2))
        self.responsible_combo = ctk.CTkComboBox(container, values=self.get_responsibles(), width=320)
        self.responsible_combo.pack(anchor="w")

        self.add_labeled_entry(container, self.tr("phone"), "phone_entry", width=320)
        self.add_labeled_entry(container, self.tr("phone_contact"), "phone_contact_person_entry", width=320)

        track_calls_label = ctk.CTkLabel(container, text=self.tr("track_calls"))
        track_calls_label.pack(anchor="w", pady=(8, 2))
        self.track_calls_combo = ctk.CTkComboBox(container, values=self.parent_tab.yes_no_values(), width=160)
        self.track_calls_combo.pack(anchor="w")
        self.track_calls_combo.set(self.tr("yes"))

        birthday_label = ctk.CTkLabel(container, text=self.tr("birthday"))
        birthday_label.pack(anchor="w", pady=(8, 2))

        birthday_frame = ctk.CTkFrame(container, fg_color="transparent")
        birthday_frame.pack(anchor="w")

        self.birthday_entry = ctk.CTkEntry(birthday_frame, width=200, placeholder_text=self.tr("birthday_placeholder"))
        self.birthday_entry.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            birthday_frame,
            text=self.tr("pick_date"),
            width=120,
            command=self.open_calendar,
        ).pack(side="left")

        self.add_labeled_textbox(container, self.tr("comment"), "comment_text", width=640, height=100)

        ctk.CTkButton(container, text=self.tr("save"), command=self.save, width=220).pack(pady=20)

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
        self.track_calls_combo.set(self.parent_tab.to_display_yes_no(self.person.track_calls or "Да"))

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
            self.show_error(self.tr("fio_required"))
            return

        birthday = self.parse_date(self.birthday_entry.get())
        if birthday == "INVALID":
            self.show_error(self.tr("birthday_invalid"))
            return

        session = get_session()

        if self.mode == "edit":
            person = session.query(Person).filter(Person.id == self.person_id).first()
            if not person:
                session.close()
                self.show_error(self.tr("person_not_found"))
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
        person.track_calls = self.parent_tab.from_display_yes_no(self.track_calls_combo.get().strip())
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
        win.title(self.tr("error_title"))
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=380).pack(pady=20, padx=20)
        ctk.CTkButton(win, text=self.tr("ok"), command=win.destroy).pack()
