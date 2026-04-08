from __future__ import annotations

from datetime import date

import customtkinter as ctk
from tkinter import ttk

from services.reminder_service import ReminderService, ReminderItem
from services.ics_export import export_interactions_to_ics
from ui.statuses import normalize_language, translate_interaction_status


TRANSLATIONS = {
    "ru": {
        "responsible": "Ответственный",
        "circle": "Круг",
        "all": "Все",
        "btn_refresh": "🔄 Обновить",
        "btn_export_ics": "📅 Экспорт в ICS",
        "btn_copy_message": "📋 Скопировать сообщение",
        "counter_today": "🟡 Сегодня: {count}",
        "counter_week": "🟢 На 7 дней: {count}",
        "card_birthdays_today": "🎂 Дни рождения сегодня",
        "card_contacts_today": "📞 Контакты на сегодня",
        "card_birthdays_week": "🎉 Ближайшие дни рождения",
        "card_contacts_week": "🗓 Контакты на ближайшие 7 дней",
        "col_fio": "ФИО",
        "col_date": "Дата",
        "col_type": "Тип",
        "col_responsible": "Ответственный",
        "message_title": "Сообщение",
        "msg_copied": "Сообщение скопировано в буфер обмена",
        "msg_no_contacts_ics": "Нет контактов на ближайшие дни для экспорта в ICS",
        "msg_ics_done": "Экспорт в ICS завершён:\n{path}",
        "msg_ics_error": "Ошибка экспорта в ICS:\n{error}",
        "hello": "Всем здравствуйте!",
        "birthdays_today": "Дни рождения сегодня:",
        "birthdays_none": "На сегодня дни рождения отсутствуют",
        "contacts_today": "Сегодня напоминаем написать/позвонить следующим людям для поддержания контакта с ними:",
        "contacts_none": "На сегодня напоминаний по поддержанию контактов не запланировано.",
        "birthday_title": "День рождения",
        "birthday_in_days": "Через {days} дн.",
        "contact_default": "Контакт",
        "type_call": "Звонок",
        "type_meeting": "Встреча",
    },
    "en": {
        "responsible": "Responsible",
        "circle": "Circle",
        "all": "All",
        "btn_refresh": "🔄 Refresh",
        "btn_export_ics": "📅 Export to ICS",
        "btn_copy_message": "📋 Copy message",
        "counter_today": "🟡 Today: {count}",
        "counter_week": "🟢 Next 7 days: {count}",
        "card_birthdays_today": "🎂 Birthdays today",
        "card_contacts_today": "📞 Contacts for today",
        "card_birthdays_week": "🎉 Upcoming birthdays",
        "card_contacts_week": "🗓 Contacts for the next 7 days",
        "col_fio": "Full name",
        "col_date": "Date",
        "col_type": "Type",
        "col_responsible": "Responsible",
        "message_title": "Message",
        "msg_copied": "Message copied to clipboard",
        "msg_no_contacts_ics": "No upcoming contacts to export to ICS",
        "msg_ics_done": "ICS export completed:\n{path}",
        "msg_ics_error": "ICS export error:\n{error}",
        "hello": "Hello everyone!",
        "birthdays_today": "Birthdays today:",
        "birthdays_none": "There are no birthdays today",
        "contacts_today": "Today, please message/call the following people to keep in touch:",
        "contacts_none": "No contact reminders are scheduled for today.",
        "birthday_title": "Birthday",
        "birthday_in_days": "In {days} days",
        "contact_default": "Contact",
        "type_call": "Call",
        "type_meeting": "Meeting",
    },
}


class ReminderTab(ctk.CTkFrame):
    def __init__(self, parent, reminder_service: ReminderService):
        super().__init__(parent)
        self.reminder_service = reminder_service
        self.configure(fg_color="transparent")

        self._build_topbar()
        self._build_counters()
        self._build_cards()
        self.refresh()

    def get_language(self) -> str:
        root = self.winfo_toplevel()
        return normalize_language(getattr(root, "current_language", "ru"))

    def tr(self, key: str, **kwargs) -> str:
        text = TRANSLATIONS.get(self.get_language(), TRANSLATIONS["ru"]).get(key, key)
        return text.format(**kwargs) if kwargs else text

    def all_value(self) -> str:
        return self.tr("all")

    def to_service_filter(self, value: str) -> str:
        return "Все" if (value or "").strip() == self.all_value() else (value or "").strip()

    def from_service_filter(self, value: str) -> str:
        return self.all_value() if (value or "").strip() == "Все" else (value or "").strip()

    def display_subtitle(self, item: ReminderItem) -> str:
        raw = (item.subtitle or item.title or "").strip()
        if not raw:
            return self.tr("contact_default") if item.kind == "contact" else self.tr("birthday_title")

        lowered = raw.lower()
        if lowered == "звонок":
            return self.tr("type_call")
        if lowered == "встреча":
            return self.tr("type_meeting")
        if lowered == "контакт":
            return self.tr("contact_default")
        if lowered == "день рождения":
            return self.tr("birthday_title")
        return raw

    def _build_topbar(self):
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=10, pady=(10, 6))

        self.responsible_label = ctk.CTkLabel(top, text=self.tr("responsible"))
        self.responsible_label.pack(side="left", padx=(10, 5), pady=10)

        self.responsible_filter = ctk.CTkComboBox(top, values=[self.all_value()], width=180, command=lambda _=None: self.refresh())
        self.responsible_filter.pack(side="left", padx=5, pady=10)
        self.responsible_filter.set(self.all_value())

        self.circle_label = ctk.CTkLabel(top, text=self.tr("circle"))
        self.circle_label.pack(side="left", padx=(15, 5), pady=10)

        self.circle_filter = ctk.CTkComboBox(top, values=[self.all_value()], width=150, command=lambda _=None: self.refresh())
        self.circle_filter.pack(side="left", padx=5, pady=10)
        self.circle_filter.set(self.all_value())

        self.refresh_button = ctk.CTkButton(top, text=self.tr("btn_refresh"), width=140, command=self.refresh)
        self.refresh_button.pack(side="right", padx=10, pady=10)
        self.export_button = ctk.CTkButton(top, text=self.tr("btn_export_ics"), width=160, command=self.export_to_ics)
        self.export_button.pack(side="right", padx=5, pady=10)
        self.copy_button = ctk.CTkButton(top, text=self.tr("btn_copy_message"), width=210, command=self.copy_message)
        self.copy_button.pack(side="right", padx=5, pady=10)

    def _build_counters(self):
        counter = ctk.CTkFrame(self)
        counter.pack(fill="x", padx=10, pady=(0, 10))

        self.today_total_label = ctk.CTkLabel(counter, text=self.tr("counter_today", count=0), font=ctk.CTkFont(size=16, weight="bold"))
        self.today_total_label.pack(side="left", padx=18, pady=12)

        self.week_total_label = ctk.CTkLabel(counter, text=self.tr("counter_week", count=0), font=ctk.CTkFont(size=16, weight="bold"))
        self.week_total_label.pack(side="left", padx=18, pady=12)

    def _build_cards(self):
        self.cards_host = ctk.CTkFrame(self)
        self.cards_host.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.cards_host.grid_columnconfigure(0, weight=1)
        self.cards_host.grid_columnconfigure(1, weight=1)
        self.cards_host.grid_rowconfigure(0, weight=1)
        self.cards_host.grid_rowconfigure(1, weight=1)

        self.birthdays_today_card = self._create_card(self.cards_host, self.tr("card_birthdays_today"))
        self.contacts_today_card = self._create_card(self.cards_host, self.tr("card_contacts_today"))
        self.birthdays_week_card = self._create_card(self.cards_host, self.tr("card_birthdays_week"))
        self.contacts_week_card = self._create_card(self.cards_host, self.tr("card_contacts_week"))

        self.birthdays_today_card["frame"].grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        self.contacts_today_card["frame"].grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 8))
        self.birthdays_week_card["frame"].grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(8, 0))
        self.contacts_week_card["frame"].grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(8, 0))

    def _create_card(self, parent, title: str):
        frame = ctk.CTkFrame(parent, corner_radius=12)
        title_label = ctk.CTkLabel(frame, text=title, anchor="w", font=ctk.CTkFont(size=22, weight="bold"))
        title_label.pack(fill="x", padx=14, pady=(12, 8))

        table_container = ctk.CTkFrame(frame, fg_color="transparent")
        table_container.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("fio", "date", "subtitle", "responsible")
        tree = ttk.Treeview(table_container, columns=columns, show="headings", height=10)
        tree.heading("fio", text=self.tr("col_fio"))
        tree.heading("date", text=self.tr("col_date"))
        tree.heading("subtitle", text=self.tr("col_type"))
        tree.heading("responsible", text=self.tr("col_responsible"))
        tree.column("fio", width=260)
        tree.column("date", width=90, anchor="center")
        tree.column("subtitle", width=130, anchor="center")
        tree.column("responsible", width=150)
        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=tree.yview)
        scrollbar.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar.set)

        return {"frame": frame, "title": title_label, "tree": tree}

    def refresh(self):
        filter_values = self.reminder_service.get_filter_values()

        current_responsible = self.responsible_filter.get().strip()
        current_circle = self.circle_filter.get().strip()

        responsible_values = [self.from_service_filter(v) for v in filter_values["responsibles"]]
        circle_values = [self.from_service_filter(v) for v in filter_values["circles"]]

        self.responsible_filter.configure(values=responsible_values)
        if current_responsible not in responsible_values:
            self.responsible_filter.set(self.all_value())
        else:
            self.responsible_filter.set(current_responsible)

        self.circle_filter.configure(values=circle_values)
        if current_circle not in circle_values:
            self.circle_filter.set(self.all_value())
        else:
            self.circle_filter.set(current_circle)

        payload = self.reminder_service.get_dashboard_payload(
            days=7,
            responsible=self.to_service_filter(self.responsible_filter.get()),
            circle=self.to_service_filter(self.circle_filter.get()),
        )

        self.today_total_label.configure(text=self.tr("counter_today", count=payload["today_total"]))
        self.week_total_label.configure(text=self.tr("counter_week", count=payload["week_total"]))

        self._fill_tree(self.birthdays_today_card["tree"], payload["birthdays_today"])
        self._fill_tree(self.contacts_today_card["tree"], payload["contacts_today"])
        self._fill_tree(self.birthdays_week_card["tree"], payload["birthdays_week"])
        self._fill_tree(self.contacts_week_card["tree"], payload["contacts_week"])

    def _fill_tree(self, tree: ttk.Treeview, items: list[ReminderItem]):
        for row in tree.get_children():
            tree.delete(row)
        for item in items:
            tree.insert(
                "",
                "end",
                values=(
                    item.fio,
                    item.event_date.strftime("%d.%m.%Y"),
                    self.display_subtitle(item),
                    item.responsible or "",
                )
            )

    def _build_compact_message(self) -> str:
        payload = self.reminder_service.get_dashboard_payload(
            days=7,
            responsible=self.to_service_filter(self.responsible_filter.get()),
            circle=self.to_service_filter(self.circle_filter.get()),
        )

        birthdays_today = payload["birthdays_today"]
        contacts_today = payload["contacts_today"]

        lines = [self.tr("hello"), ""]

        if birthdays_today:
            lines.append(self.tr("birthdays_today"))
            for item in birthdays_today:
                lines.append(f"— {item.fio}")
        else:
            lines.append(self.tr("birthdays_none"))

        lines.append("")

        if contacts_today:
            lines.append(self.tr("contacts_today"))
            for item in contacts_today:
                lines.append(f"— {item.fio}")
        else:
            lines.append(self.tr("contacts_none"))

        return "\n".join(lines).strip()

    def copy_message(self):
        text = self._build_compact_message()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.show_message(self.tr("msg_copied"))

    def export_to_ics(self):
        payload = self.reminder_service.get_dashboard_payload(
            days=7,
            responsible=self.to_service_filter(self.responsible_filter.get()),
            circle=self.to_service_filter(self.circle_filter.get()),
        )

        today = date.today()
        contacts = [item for item in payload["contacts_week"] if item.event_date > today]

        rows = []
        if self.reminder_service.interaction_service:
            all_rows = self.reminder_service.interaction_service.list_active_interactions(
                responsible=self.to_service_filter(self.responsible_filter.get()),
            )
            interaction_ids = {item.interaction_id for item in contacts if item.interaction_id}
            for row in all_rows:
                interaction = row["interaction"]
                if interaction.id in interaction_ids and interaction.next_date:
                    rows.append((interaction, row["person"], row["status"]))

        if not rows:
            self.show_message(self.tr("msg_no_contacts_ics"))
            return

        try:
            output_path = export_interactions_to_ics(rows, include_status=False)
            self.show_message(self.tr("msg_ics_done", path=output_path))
        except Exception as exc:
            self.show_message(self.tr("msg_ics_error", error=exc))

    def show_message(self, text: str):
        win = ctk.CTkToplevel(self)
        win.geometry("460x180")
        win.title(self.tr("message_title"))
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=420, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack(pady=10)
