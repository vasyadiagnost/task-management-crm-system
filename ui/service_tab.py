import customtkinter as ctk
from tkinter import ttk

from database import get_session
from models.reference import Responsible, Circle
from models.app_setting import AppSetting


class ServiceTab(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)

        self.selected_responsible_id = None
        self.selected_circle_id = None

        container = ctk.CTkFrame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        left_frame = ctk.CTkFrame(container)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5), pady=5)

        right_frame = ctk.CTkFrame(container)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)

        self.build_responsibles_block(left_frame)
        self.build_circles_block(right_frame)
        self.build_settings_block(right_frame)

        self.refresh_all()

    # =========================================================
    # RESPONSIBLES
    # =========================================================

    def build_responsibles_block(self, parent):
        title = ctk.CTkLabel(parent, text="Ответственные", font=ctk.CTkFont(size=18, weight="normal"))
        title.pack(pady=(10, 8))

        top_btn_frame = ctk.CTkFrame(parent)
        top_btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkButton(top_btn_frame, text="➕ Новый", width=120, command=self.add_responsible).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(top_btn_frame, text="✏️ Редактировать", width=140, command=self.edit_responsible).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(top_btn_frame, text="🗑 Удалить", width=120, command=self.delete_responsible).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(top_btn_frame, text="🔄 Обновить", width=120, command=self.refresh_responsibles).pack(side="right", padx=5, pady=5)

        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.responsible_tree = ttk.Treeview(
            table_frame,
            columns=("id", "name"),
            show="headings"
        )
        self.responsible_tree.heading("id", text="ID")
        self.responsible_tree.heading("name", text="ФИО / Наименование")

        self.responsible_tree.column("id", width=70, anchor="center")
        self.responsible_tree.column("name", width=320)

        self.responsible_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.responsible_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.responsible_tree.configure(yscrollcommand=scrollbar.set)

        self.responsible_tree.bind("<<TreeviewSelect>>", self.on_responsible_select)
        self.responsible_tree.bind("<Double-1>", self.on_responsible_double_click)

    def refresh_responsibles(self):
        for row in self.responsible_tree.get_children():
            self.responsible_tree.delete(row)

        session = get_session()
        items = session.query(Responsible).order_by(Responsible.name).all()

        for item in items:
            self.responsible_tree.insert("", "end", values=(item.id, item.name))

        session.close()
        self.selected_responsible_id = None

    def on_responsible_select(self, event):
        selected = self.responsible_tree.selection()
        if selected:
            values = self.responsible_tree.item(selected[0])["values"]
            self.selected_responsible_id = values[0]

    def on_responsible_double_click(self, event):
        self.edit_responsible()

    def add_responsible(self):
        ReferenceEditWindow(
            parent=self,
            title="Новый ответственный",
            entity_type="responsible",
            on_saved=self.refresh_responsibles
        )

    def edit_responsible(self):
        if not self.selected_responsible_id:
            self.show_message("Выбери ответственного")
            return

        session = get_session()
        item = session.get(Responsible, self.selected_responsible_id)
        session.close()

        if not item:
            self.show_message("Запись не найдена")
            return

        ReferenceEditWindow(
            parent=self,
            title="Редактирование ответственного",
            entity_type="responsible",
            item_id=item.id,
            initial_value=item.name,
            on_saved=self.refresh_responsibles
        )

    def delete_responsible(self):
        if not self.selected_responsible_id:
            self.show_message("Выбери ответственного")
            return

        session = get_session()
        item = session.get(Responsible, self.selected_responsible_id)
        if not item:
            session.close()
            self.show_message("Запись не найдена")
            return

        from models.person import Person
        linked_person = session.query(Person).filter(Person.responsible == item.name).first()
        if linked_person:
            session.close()
            self.show_message("Нельзя удалить: значение уже используется в карточках людей")
            return

        session.delete(item)
        session.commit()
        session.close()

        self.refresh_responsibles()

    # =========================================================
    # CIRCLES
    # =========================================================

    def build_circles_block(self, parent):
        title = ctk.CTkLabel(parent, text="Круги общения", font=ctk.CTkFont(size=18, weight="normal"))
        title.pack(pady=(10, 8))

        top_btn_frame = ctk.CTkFrame(parent)
        top_btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkButton(top_btn_frame, text="➕ Новый", width=120, command=self.add_circle).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(top_btn_frame, text="✏️ Редактировать", width=140, command=self.edit_circle).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(top_btn_frame, text="🗑 Удалить", width=120, command=self.delete_circle).pack(side="left", padx=5, pady=5)
        ctk.CTkButton(top_btn_frame, text="🔄 Обновить", width=120, command=self.refresh_circles).pack(side="right", padx=5, pady=5)

        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.circle_tree = ttk.Treeview(
            table_frame,
            columns=("id", "name", "period"),
            show="headings"
        )
        self.circle_tree.heading("id", text="ID")
        self.circle_tree.heading("name", text="Наименование круга")
        self.circle_tree.heading("period", text="Периодичность, дней")

        self.circle_tree.column("id", width=70, anchor="center")
        self.circle_tree.column("name", width=180)
        self.circle_tree.column("period", width=160, anchor="center")

        self.circle_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.circle_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.circle_tree.configure(yscrollcommand=scrollbar.set)

        self.circle_tree.bind("<<TreeviewSelect>>", self.on_circle_select)
        self.circle_tree.bind("<Double-1>", self.on_circle_double_click)

    def refresh_circles(self):
        for row in self.circle_tree.get_children():
            self.circle_tree.delete(row)

        session = get_session()
        items = session.query(Circle).order_by(Circle.name).all()

        for item in items:
            self.circle_tree.insert("", "end", values=(item.id, item.name, item.contact_period_days))

        session.close()
        self.selected_circle_id = None

    def on_circle_select(self, event):
        selected = self.circle_tree.selection()
        if selected:
            values = self.circle_tree.item(selected[0])["values"]
            self.selected_circle_id = values[0]

    def on_circle_double_click(self, event):
        self.edit_circle()

    def add_circle(self):
        CircleEditWindow(
            parent=self,
            title="Новый круг общения",
            on_saved=self.refresh_circles
        )

    def edit_circle(self):
        if not self.selected_circle_id:
            self.show_message("Выбери круг общения")
            return

        session = get_session()
        item = session.get(Circle, self.selected_circle_id)
        session.close()

        if not item:
            self.show_message("Запись не найдена")
            return

        CircleEditWindow(
            parent=self,
            title="Редактирование круга общения",
            item_id=item.id,
            initial_name=item.name,
            initial_period=item.contact_period_days,
            on_saved=self.refresh_circles
        )

    def delete_circle(self):
        if not self.selected_circle_id:
            self.show_message("Выбери круг общения")
            return

        session = get_session()
        item = session.get(Circle, self.selected_circle_id)
        if not item:
            session.close()
            self.show_message("Запись не найдена")
            return

        from models.person import Person
        linked_person = session.query(Person).filter(Person.circle == item.name).first()
        if linked_person:
            session.close()
            self.show_message("Нельзя удалить: значение уже используется в карточках людей")
            return

        session.delete(item)
        session.commit()
        session.close()

        self.refresh_circles()

    # =========================================================
    # SETTINGS
    # =========================================================

    def build_settings_block(self, parent):
        self.settings_frame = ctk.CTkFrame(parent)
        self.settings_frame.pack(fill="x", padx=10, pady=(0, 10))

        title = ctk.CTkLabel(self.settings_frame, text="Настройки контактов", font=ctk.CTkFont(size=16, weight="normal"))
        title.pack(anchor="w", padx=10, pady=(10, 5))

        self.meeting_equals_call_var = ctk.BooleanVar(value=True)

        self.meeting_equals_call_checkbox = ctk.CTkCheckBox(
            self.settings_frame,
            text="Встреча = звонок",
            variable=self.meeting_equals_call_var,
            command=self.save_contact_settings
        )
        self.meeting_equals_call_checkbox.pack(anchor="w", padx=10, pady=(0, 10))

    def load_settings(self):
        session = get_session()
        setting = session.query(AppSetting).filter(AppSetting.key == "meeting_equals_call").first()
        session.close()

        if setting:
            self.meeting_equals_call_var.set(setting.value == "1")
        else:
            self.meeting_equals_call_var.set(True)

    def save_contact_settings(self):
        value = "1" if self.meeting_equals_call_var.get() else "0"

        session = get_session()
        setting = session.query(AppSetting).filter(AppSetting.key == "meeting_equals_call").first()
        if not setting:
            setting = AppSetting(key="meeting_equals_call", value=value)
            session.add(setting)
        else:
            setting.value = value

        session.commit()
        session.close()

    # =========================================================

    def refresh_all(self):
        self.refresh_responsibles()
        self.refresh_circles()
        self.load_settings()

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("420x160")
        win.title("Сообщение")
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=380, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack(pady=10)


class ReferenceEditWindow(ctk.CTkToplevel):
    def __init__(self, parent, title, entity_type, on_saved, item_id=None, initial_value=""):
        super().__init__(parent)

        self.entity_type = entity_type
        self.item_id = item_id
        self.on_saved = on_saved

        self.title(title)
        self.geometry("420x180")
        self.resizable(True, False)
        self.grab_set()

        self.label = ctk.CTkLabel(self, text="ФИО / Наименование")
        self.label.pack(pady=(20, 5))

        self.entry = ctk.CTkEntry(self, width=320)
        self.entry.pack(pady=5)
        self.entry.insert(0, initial_value)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(btn_frame, text="Сохранить", width=120, command=self.save).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Отмена", width=120, command=self.destroy).pack(side="left", padx=5)

    def save(self):
        value = self.entry.get().strip()
        if not value:
            self.show_error("Поле не должно быть пустым")
            return

        session = get_session()

        duplicate = session.query(Responsible).filter(Responsible.name == value).first()
        if duplicate and duplicate.id != self.item_id:
            session.close()
            self.show_error("Такое значение уже существует")
            return

        if self.item_id:
            item = session.get(Responsible, self.item_id)
            if not item:
                session.close()
                self.show_error("Запись не найдена")
                return

            old_value = item.name
            item.name = value

            from models.person import Person
            linked_persons = session.query(Person).filter(Person.responsible == old_value).all()
            for person in linked_persons:
                person.responsible = value
        else:
            item = Responsible(name=value)
            session.add(item)

        session.commit()
        session.close()

        if self.on_saved:
            self.on_saved()

        self.destroy()

    def show_error(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("340x140")
        win.title("Ошибка")
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=300).pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack()


class CircleEditWindow(ctk.CTkToplevel):
    def __init__(self, parent, title, on_saved, item_id=None, initial_name="", initial_period=0):
        super().__init__(parent)

        self.item_id = item_id
        self.on_saved = on_saved

        self.title(title)
        self.geometry("420x260")
        self.resizable(True, False)
        self.grab_set()

        self.name_label = ctk.CTkLabel(self, text="Наименование круга")
        self.name_label.pack(pady=(20, 5))

        self.name_entry = ctk.CTkEntry(self, width=320)
        self.name_entry.pack(pady=5)
        self.name_entry.insert(0, initial_name)

        self.period_label = ctk.CTkLabel(self, text="Периодичность контакта, дней")
        self.period_label.pack(pady=(15, 5))

        self.period_entry = ctk.CTkEntry(self, width=320, placeholder_text="например: 14")
        self.period_entry.pack(pady=5)
        if initial_period:
            self.period_entry.insert(0, str(initial_period))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(btn_frame, text="Сохранить", width=120, command=self.save).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Отмена", width=120, command=self.destroy).pack(side="left", padx=5)

    def save(self):
        name_value = self.name_entry.get().strip()
        if not name_value:
            self.show_error("Наименование круга не должно быть пустым")
            return

        period_raw = self.period_entry.get().strip()
        if not period_raw:
            self.show_error("Периодичность должна быть заполнена")
            return

        try:
            period_value = int(period_raw)
            if period_value < 0:
                raise ValueError
        except ValueError:
            self.show_error("Периодичность должна быть целым числом 0 или больше")
            return

        session = get_session()

        duplicate = session.query(Circle).filter(Circle.name == name_value).first()
        if duplicate and duplicate.id != self.item_id:
            session.close()
            self.show_error("Такое значение уже существует")
            return

        if self.item_id:
            item = session.get(Circle, self.item_id)
            if not item:
                session.close()
                self.show_error("Запись не найдена")
                return

            old_value = item.name
            item.name = name_value
            item.contact_period_days = period_value

            from models.person import Person
            linked_persons = session.query(Person).filter(Person.circle == old_value).all()
            for person in linked_persons:
                person.circle = name_value
        else:
            item = Circle(name=name_value, contact_period_days=period_value)
            session.add(item)

        session.commit()
        session.close()

        if self.on_saved:
            self.on_saved()

        self.destroy()

    def show_error(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("360x140")
        win.title("Ошибка")
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=320).pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack()