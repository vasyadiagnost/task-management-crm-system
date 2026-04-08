from __future__ import annotations

from datetime import date

import customtkinter as ctk
from tkinter import ttk

from services.reminder_service import ReminderService, ReminderItem
from services.ics_export import export_interactions_to_ics


class ReminderTab(ctk.CTkFrame):
    def __init__(self, parent, reminder_service: ReminderService):
        super().__init__(parent)
        self.reminder_service = reminder_service

        self.configure(fg_color="transparent")

        self._build_topbar()
        self._build_counters()
        self._build_cards()
        self.refresh()

    def _build_topbar(self):
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=10, pady=(10, 6))

        self.responsible_label = ctk.CTkLabel(top, text="Ответственный")
        self.responsible_label.pack(side="left", padx=(10, 5), pady=10)

        self.responsible_filter = ctk.CTkComboBox(top, values=["Все"], width=180, command=lambda _=None: self.refresh())
        self.responsible_filter.pack(side="left", padx=5, pady=10)
        self.responsible_filter.set("Все")

        self.circle_label = ctk.CTkLabel(top, text="Круг")
        self.circle_label.pack(side="left", padx=(15, 5), pady=10)

        self.circle_filter = ctk.CTkComboBox(top, values=["Все"], width=150, command=lambda _=None: self.refresh())
        self.circle_filter.pack(side="left", padx=5, pady=10)
        self.circle_filter.set("Все")

        ctk.CTkButton(top, text="🔄 Обновить", width=140, command=self.refresh).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(top, text="📅 Экспорт в ICS", width=160, command=self.export_to_ics).pack(side="right", padx=5, pady=10)
        ctk.CTkButton(top, text="📋 Скопировать сообщение", width=210, command=self.copy_message).pack(side="right", padx=5, pady=10)

    def _build_counters(self):
        counter = ctk.CTkFrame(self)
        counter.pack(fill="x", padx=10, pady=(0, 10))

        self.today_total_label = ctk.CTkLabel(counter, text="🟡 Сегодня: 0", font=ctk.CTkFont(size=16, weight="bold"))
        self.today_total_label.pack(side="left", padx=18, pady=12)

        self.week_total_label = ctk.CTkLabel(counter, text="🟢 На 7 дней: 0", font=ctk.CTkFont(size=16, weight="bold"))
        self.week_total_label.pack(side="left", padx=18, pady=12)

    def _build_cards(self):
        self.cards_host = ctk.CTkFrame(self)
        self.cards_host.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.cards_host.grid_columnconfigure(0, weight=1)
        self.cards_host.grid_columnconfigure(1, weight=1)
        self.cards_host.grid_rowconfigure(0, weight=1)
        self.cards_host.grid_rowconfigure(1, weight=1)

        self.birthdays_today_card = self._create_card(self.cards_host, "🎂 Дни рождения сегодня")
        self.contacts_today_card = self._create_card(self.cards_host, "📞 Контакты на сегодня")
        self.birthdays_week_card = self._create_card(self.cards_host, "🎉 Ближайшие дни рождения")
        self.contacts_week_card = self._create_card(self.cards_host, "🗓 Контакты на ближайшие 7 дней")

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
        tree.heading("fio", text="ФИО")
        tree.heading("date", text="Дата")
        tree.heading("subtitle", text="Тип")
        tree.heading("responsible", text="Ответственный")
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
        self.responsible_filter.configure(values=filter_values["responsibles"])
        if self.responsible_filter.get() not in filter_values["responsibles"]:
            self.responsible_filter.set("Все")

        self.circle_filter.configure(values=filter_values["circles"])
        if self.circle_filter.get() not in filter_values["circles"]:
            self.circle_filter.set("Все")

        payload = self.reminder_service.get_dashboard_payload(
            days=7,
            responsible=self.responsible_filter.get(),
            circle=self.circle_filter.get(),
        )

        self.today_total_label.configure(text=f"🟡 Сегодня: {payload['today_total']}")
        self.week_total_label.configure(text=f"🟢 На 7 дней: {payload['week_total']}")

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
                    item.subtitle or item.title,
                    item.responsible or "",
                )
            )

    def _build_compact_message(self) -> str:
        payload = self.reminder_service.get_dashboard_payload(
            days=7,
            responsible=self.responsible_filter.get(),
            circle=self.circle_filter.get(),
        )

        birthdays_today = payload["birthdays_today"]
        contacts_today = payload["contacts_today"]

        lines = ["Всем здравствуйте!", ""]

        if birthdays_today:
            lines.append("Дни рождения сегодня:")
            for item in birthdays_today:
                lines.append(f"— {item.fio}")
        else:
            lines.append("На сегодня дни рождения отсутствуют")

        lines.append("")

        if contacts_today:
            lines.append("Сегодня напоминаем написать/позвонить следующим людям для поддержания контакта с ними:")
            for item in contacts_today:
                lines.append(f"— {item.fio}")
        else:
            lines.append("На сегодня напоминаний по поддержанию контактов не запланировано.")

        return "\n".join(lines).strip()

    def copy_message(self):
        text = self._build_compact_message()
        self.clipboard_clear()
        self.clipboard_append(text)
        self.show_message("Сообщение скопировано в буфер обмена")

    def export_to_ics(self):
        payload = self.reminder_service.get_dashboard_payload(
            days=7,
            responsible=self.responsible_filter.get(),
            circle=self.circle_filter.get(),
        )

        today = date.today()
        contacts = [item for item in payload["contacts_week"] if item.event_date > today]

        rows = []
        if self.reminder_service.interaction_service:
            all_rows = self.reminder_service.interaction_service.list_active_interactions(
                responsible=self.responsible_filter.get(),
            )
            interaction_ids = {item.interaction_id for item in contacts if item.interaction_id}
            for row in all_rows:
                interaction = row["interaction"]
                if interaction.id in interaction_ids and interaction.next_date:
                    rows.append((interaction, row["person"], row["status"]))

        if not rows:
            self.show_message("Нет контактов на ближайшие дни для экспорта в ICS")
            return

        try:
            output_path = export_interactions_to_ics(rows, include_status=False)
            self.show_message(f"Экспорт в ICS завершён:\n{output_path}")
        except Exception as exc:
            self.show_message(f"Ошибка экспорта в ICS:\n{exc}")

    def show_message(self, text: str):
        win = ctk.CTkToplevel(self)
        win.geometry("460x180")
        win.title("Сообщение")
        win.grab_set()

        ctk.CTkLabel(win, text=text, wraplength=420, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy).pack(pady=10)
