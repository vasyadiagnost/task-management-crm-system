
from __future__ import annotations

from pathlib import Path

import customtkinter as ctk
from tkinter import ttk, filedialog

from services.database_editor_service import DatabaseEditorService


class DatabaseEditorTab(ctk.CTkFrame):
    def __init__(self, parent, editor_service: DatabaseEditorService | None = None, database_editor_service: DatabaseEditorService | None = None):
        super().__init__(parent)
        self.editor_service = editor_service or database_editor_service or DatabaseEditorService()

        self.current_table: str | None = None
        self.columns_meta: list[dict] = []
        self.selected_row_id = None
        self.rows_cache: list[dict] = []

        self.configure(fg_color="transparent")

        self._build_filters_row()
        self._build_actions_row()
        self._build_info()
        self._build_table()
        self._load_tables()

    def _build_filters_row(self):
        top = ctk.CTkFrame(self, corner_radius=14)
        top.pack(fill="x", padx=10, pady=(10, 6))

        ctk.CTkLabel(top, text="Таблица", font=ctk.CTkFont(size=12, weight="normal")).pack(side="left", padx=(12, 5), pady=10)
        self.table_combo = ctk.CTkComboBox(top, values=[""], width=220, command=self.on_table_change, corner_radius=10)
        self.table_combo.pack(side="left", padx=5, pady=10)

        ctk.CTkLabel(top, text="Поиск", font=ctk.CTkFont(size=12, weight="normal")).pack(side="left", padx=(12, 5), pady=10)
        self.search_entry = ctk.CTkEntry(top, width=300, placeholder_text="Поиск по выбранной таблице...")
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

    def _build_actions_row(self):
        actions = ctk.CTkFrame(self, corner_radius=14)
        actions.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkButton(actions, text="🔄 Обновить", width=120, command=self.refresh, corner_radius=10, height=34).pack(side="left", padx=10, pady=10)
        ctk.CTkButton(actions, text="✏️ Редактировать ячейку", width=170, command=self.open_edit_cell_window, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(actions, text="🗑 Удалить", width=120, command=self.delete_selected_row, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(actions, text="➕ Новая строка", width=140, command=self.open_add_row_window, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(actions, text="📊 Выгрузка таблицы Excel", width=190, command=self.export_current_table, corner_radius=10, height=34).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(actions, text="📥 Массовая загрузка Excel", width=190, command=self.bulk_import_excel, corner_radius=10, height=34).pack(side="right", padx=5, pady=10)
        ctk.CTkButton(actions, text="📤 Шаблон Excel", width=150, command=self.export_template_excel, corner_radius=10, height=34).pack(side="right", padx=5, pady=10)

    def _build_info(self):
        info = ctk.CTkFrame(self, corner_radius=14)
        info.pack(fill="x", padx=10, pady=(0, 10))
        self.info_label = ctk.CTkLabel(info, text="Выбери таблицу для просмотра и редактирования", font=ctk.CTkFont(size=14, weight="normal"))
        self.info_label.pack(side="left", padx=14, pady=12)

    def _build_table(self):
        card = ctk.CTkFrame(self, corner_radius=14)
        card.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.table_title = ctk.CTkLabel(card, text="Database Editor", anchor="w", font=ctk.CTkFont(size=18, weight="normal"))
        self.table_title.pack(fill="x", padx=14, pady=(12, 8))

        table_holder = ctk.CTkFrame(card, fg_color="transparent")
        table_holder.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tree = ttk.Treeview(table_holder, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_y = ttk.Scrollbar(table_holder, orient="vertical", command=self.tree.yview)
        scroll_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll_y.set)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", lambda event: self.open_edit_cell_window())

    def _load_tables(self):
        tables = self.editor_service.get_table_names()
        self.table_combo.configure(values=tables or [""])
        if tables:
            self.table_combo.set(tables[0])
            self.on_table_change(tables[0])

    def on_table_change(self, table_name: str):
        self.current_table = table_name
        self.search_entry.delete(0, "end")
        self.refresh()

    def refresh(self):
        if not self.current_table:
            return

        self.columns_meta = self.editor_service.get_table_columns(self.current_table)
        self.table_title.configure(text=f"Database Editor — {self.current_table}")

        query = self.search_entry.get().strip()
        if query:
            self.rows_cache = self.editor_service.search_rows(self.current_table, query=query)
        else:
            self.rows_cache = self.editor_service.list_rows(self.current_table)

        self._rebuild_tree()
        self.info_label.configure(text=f"Таблица: {self.current_table} | строк: {len(self.rows_cache)} | столбцов: {len(self.columns_meta)}")
        self.selected_row_id = None

    def _rebuild_tree(self):
        columns = [c["name"] for c in self.columns_meta]
        self.tree.configure(columns=columns)
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
            self.tree.column(col, width=120, anchor="w")
        for col_meta in self.columns_meta:
            name = col_meta["name"]
            label = f"{name} *" if col_meta["primary_key"] else name
            self.tree.heading(name, text=label)
            width = 90 if col_meta["primary_key"] else 150
            self.tree.column(name, width=width, anchor="center" if name.endswith("id") else "w")

        for row in self.tree.get_children():
            self.tree.delete(row)

        for row_data in self.rows_cache:
            values = [self._display_value(row_data.get(col["name"])) for col in self.columns_meta]
            self.tree.insert("", "end", values=values)

    def _display_value(self, value):
        if value is None:
            return ""
        return str(value)

    def on_select(self, event=None):
        selected = self.tree.selection()
        if not selected or not self.columns_meta:
            self.selected_row_id = None
            return
        pk_col = next((c["name"] for c in self.columns_meta if c["primary_key"]), None)
        if not pk_col:
            self.selected_row_id = None
            return
        item_index = self.tree.index(selected[0])
        if 0 <= item_index < len(self.rows_cache):
            self.selected_row_id = self.rows_cache[item_index].get(pk_col)
        else:
            self.selected_row_id = None

    def bulk_import_excel(self):
        if not self.current_table:
            self.show_message("Сначала выбери таблицу")
            return
        file_path = filedialog.askopenfilename(
            title="Выбрать Excel для массовой загрузки",
            filetypes=[("Excel files", "*.xlsx *.xlsm")],
        )
        if not file_path:
            return
        ImportModeWindow(self, file_path)

    def run_bulk_import(self, file_path: str, mode: str):
        try:
            result = self.editor_service.import_from_excel(self.current_table, file_path=file_path, mode=mode)
            self.refresh()
            text = (
                f"Массовая загрузка завершена.\n\n"
                f"Таблица: {self.current_table}\n"
                f"Добавлено строк: {result['inserted']}\n"
                f"Пропущено строк: {result['skipped']}\n"
                f"Совпавшие столбцы: {', '.join(result['matched_headers'])}\n"
            )
            if result["errors"]:
                text += "\nПервые ошибки:\n" + "\n".join(result["errors"][:8])
            self.show_message(text)
        except Exception as exc:
            self.show_message(f"Ошибка массовой загрузки Excel:\n{exc}")

    def export_template_excel(self):
        if not self.current_table:
            self.show_message("Сначала выбери таблицу")
            return
        initial_name = f"{self.current_table}_template.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Сохранить шаблон Excel",
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return
        try:
            self.editor_service.create_excel_template(self.current_table, output_path)
            self.show_message(f"Шаблон Excel сохранён:\n{output_path}")
        except Exception as exc:
            self.show_message(f"Ошибка формирования шаблона Excel:\n{exc}")

    def export_current_table(self):
        if not self.current_table:
            self.show_message("Сначала выбери таблицу")
            return
        initial_name = f"{self.current_table}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Сохранить выгрузку таблицы",
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return
        try:
            self.editor_service.export_table_to_excel(self.current_table, output_path, query=self.search_entry.get().strip())
            self.show_message(f"Выгрузка таблицы сохранена:\n{output_path}")
        except Exception as exc:
            self.show_message(f"Ошибка выгрузки таблицы:\n{exc}")

    def open_edit_cell_window(self):
        if not self.current_table:
            self.show_message("Сначала выбери таблицу")
            return
        if self.selected_row_id is None:
            self.show_message("Сначала выбери строку")
            return
        EditCellWindow(self)

    def open_add_row_window(self):
        if not self.current_table:
            self.show_message("Сначала выбери таблицу")
            return
        AddRowWindow(self)

    def delete_selected_row(self):
        if not self.current_table:
            self.show_message("Сначала выбери таблицу")
            return
        if self.selected_row_id is None:
            self.show_message("Сначала выбери строку")
            return
        try:
            self.editor_service.delete_row(self.current_table, self.selected_row_id)
            self.refresh()
        except Exception as exc:
            self.show_message(f"Ошибка удаления строки:\n{exc}")

    def show_message(self, text: str):
        win = ctk.CTkToplevel(self)
        win.geometry("620x320")
        win.title("Сообщение")
        win.grab_set()
        box = ctk.CTkTextbox(win, width=560, height=220)
        box.pack(padx=20, pady=20)
        box.insert("1.0", text)
        box.configure(state="disabled")
        ctk.CTkButton(win, text="OK", command=win.destroy, width=140, corner_radius=10, height=34).pack(pady=(0, 10))


class ImportModeWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab: DatabaseEditorTab, file_path: str):
        super().__init__()
        self.parent_tab = parent_tab
        self.file_path = file_path
        self.title("Режим массовой загрузки")
        self.geometry("560x300")
        self.resizable(False, False)
        self.grab_set()

        ctk.CTkLabel(self, text="Выбери режим загрузки Excel", font=ctk.CTkFont(size=16, weight="normal")).pack(pady=(16, 8))

        info = (
            f"Файл:\n{file_path}\n\n"
            f"Append — добавить строки к существующим.\n"
            f"Replace — полностью очистить таблицу и затем загрузить Excel."
        )
        box = ctk.CTkTextbox(self, width=500, height=150)
        box.pack(padx=20, pady=10)
        box.insert("1.0", info)
        box.configure(state="disabled")

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(pady=10)
        ctk.CTkButton(btns, text="Append / Добавить", width=170, corner_radius=10, height=34, command=self.run_append).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Replace / Заменить", width=170, corner_radius=10, height=34, command=self.run_replace).pack(side="left", padx=6)

    def run_append(self):
        self.destroy()
        self.parent_tab.run_bulk_import(self.file_path, mode="append")

    def run_replace(self):
        self.destroy()
        self.parent_tab.run_bulk_import(self.file_path, mode="replace")


class EditCellWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab: DatabaseEditorTab):
        super().__init__()
        self.parent_tab = parent_tab
        self.title("Редактирование ячейки")
        self.geometry("620x300")
        self.resizable(False, False)
        self.grab_set()
        self.build()

    def build(self):
        ctk.CTkLabel(self, text="Столбец", font=ctk.CTkFont(size=12, weight="normal")).pack(pady=(14, 4))
        editable_columns = [c["name"] for c in self.parent_tab.columns_meta if not c["primary_key"]]
        self.column_combo = ctk.CTkComboBox(self, values=editable_columns or [""], width=420, corner_radius=10)
        self.column_combo.pack(pady=(0, 8))
        if editable_columns:
            self.column_combo.set(editable_columns[0])

        ctk.CTkLabel(self, text="Новое значение", font=ctk.CTkFont(size=12, weight="normal")).pack(pady=(8, 4))
        self.value_text = ctk.CTkTextbox(self, width=520, height=110)
        self.value_text.pack(pady=(0, 10))

        ctk.CTkButton(self, text="Сохранить", command=self.save, width=180, corner_radius=10, height=34).pack(pady=8)

    def save(self):
        column_name = self.column_combo.get().strip()
        value = self.value_text.get("1.0", "end").strip()
        try:
            self.parent_tab.editor_service.update_cell(
                self.parent_tab.current_table,
                self.parent_tab.selected_row_id,
                column_name,
                value,
            )
            self.parent_tab.refresh()
            self.destroy()
        except Exception as exc:
            self.parent_tab.show_message(f"Ошибка сохранения:\n{exc}")


class AddRowWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab: DatabaseEditorTab):
        super().__init__()
        self.parent_tab = parent_tab
        self.entries = {}
        self.title("Новая строка")
        self.geometry("760x700")
        self.grab_set()

        self.scroll = ctk.CTkScrollableFrame(self, width=700, height=620, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=10, pady=10)

        self.build()

    def build(self):
        for col in self.parent_tab.columns_meta:
            name = col["name"]
            ctk.CTkLabel(self.scroll, text=name, font=ctk.CTkFont(size=12, weight="normal")).pack(pady=(6, 2))
            entry = ctk.CTkEntry(self.scroll, width=620)
            entry.pack(pady=(0, 4))
            self.entries[name] = entry

        ctk.CTkButton(self.scroll, text="Добавить строку", command=self.save, width=180, corner_radius=10, height=34).pack(pady=14)

    def save(self):
        values = {name: entry.get().strip() for name, entry in self.entries.items()}
        try:
            self.parent_tab.editor_service.insert_row(self.parent_tab.current_table, values)
            self.parent_tab.refresh()
            self.destroy()
        except Exception as exc:
            self.parent_tab.show_message(f"Ошибка добавления строки:\n{exc}")
