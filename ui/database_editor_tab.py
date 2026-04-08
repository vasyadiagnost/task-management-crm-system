from __future__ import annotations

import customtkinter as ctk
from tkinter import ttk, filedialog

from services.database_editor_service import DatabaseEditorService


TRANSLATIONS = {
    "ru": {
        "table": "Таблица",
        "search": "Поиск",
        "search_placeholder": "Поиск по выбранной таблице...",
        "refresh": "🔄 Обновить",
        "edit_cell": "✏️ Редактировать ячейку",
        "delete": "🗑 Удалить",
        "new_row": "➕ Новая строка",
        "export_excel": "📊 Выгрузка таблицы Excel",
        "bulk_import": "📥 Массовая загрузка Excel",
        "excel_template": "📤 Шаблон Excel",
        "choose_table_hint": "Выбери таблицу для просмотра и редактирования",
        "editor_title": "Редактор базы данных",
        "editor_title_with_table": "Редактор базы данных — {table}",
        "table_info": "Таблица: {table} | строк: {rows} | столбцов: {cols}",
        "pick_table_first": "Сначала выбери таблицу",
        "pick_row_first": "Сначала выбери строку",
        "select_excel_import": "Выбрать Excel для массовой загрузки",
        "bulk_done": "Массовая загрузка завершена.\n\nТаблица: {table}\nДобавлено строк: {inserted}\nПропущено строк: {skipped}\nСовпавшие столбцы: {headers}\n",
        "first_errors": "\nПервые ошибки:\n{errors}",
        "bulk_error": "Ошибка массовой загрузки Excel:\n{error}",
        "save_excel_template": "Сохранить шаблон Excel",
        "template_saved": "Шаблон Excel сохранён:\n{path}",
        "template_error": "Ошибка формирования шаблона Excel:\n{error}",
        "save_excel_export": "Сохранить выгрузку таблицы",
        "export_saved": "Выгрузка таблицы сохранена:\n{path}",
        "export_error": "Ошибка выгрузки таблицы:\n{error}",
        "delete_error": "Ошибка удаления строки:\n{error}",
        "message": "Сообщение",
        "ok": "OK",
        "import_mode_title": "Режим массовой загрузки",
        "import_mode_label": "Выбери режим загрузки Excel",
        "import_mode_info": "Файл:\n{file}\n\nAppend — добавить строки к существующим.\nReplace — полностью очистить таблицу и затем загрузить Excel.",
        "append": "Append / Добавить",
        "replace": "Replace / Заменить",
        "edit_window_title": "Редактирование ячейки",
        "column": "Столбец",
        "new_value": "Новое значение",
        "save": "Сохранить",
        "save_error": "Ошибка сохранения:\n{error}",
        "add_row_title": "Новая строка",
        "add_row": "Добавить строку",
        "add_row_error": "Ошибка добавления строки:\n{error}",
        "pk_marker": " *",
        "tables": {
            "circles": "Круги общения",
            "interactions": "Взаимодействия",
            "meetings": "Встречи",
            "persons": "Люди",
            "registry_tasks": "Реестр поручений",
            "responsibles": "Ответственные",
            "tasks": "Задачи",
        },
    },
    "en": {
        "table": "Table",
        "search": "Search",
        "search_placeholder": "Search in the selected table...",
        "refresh": "🔄 Refresh",
        "edit_cell": "✏️ Edit cell",
        "delete": "🗑 Delete",
        "new_row": "➕ New row",
        "export_excel": "📊 Export table to Excel",
        "bulk_import": "📥 Bulk import from Excel",
        "excel_template": "📤 Excel template",
        "choose_table_hint": "Choose a table to view and edit",
        "editor_title": "Database Editor",
        "editor_title_with_table": "Database Editor — {table}",
        "table_info": "Table: {table} | rows: {rows} | columns: {cols}",
        "pick_table_first": "Select a table first",
        "pick_row_first": "Select a row first",
        "select_excel_import": "Select an Excel file for bulk import",
        "bulk_done": "Bulk import completed.\n\nTable: {table}\nRows added: {inserted}\nRows skipped: {skipped}\nMatched columns: {headers}\n",
        "first_errors": "\nFirst errors:\n{errors}",
        "bulk_error": "Bulk Excel import error:\n{error}",
        "save_excel_template": "Save Excel template",
        "template_saved": "Excel template saved:\n{path}",
        "template_error": "Failed to create Excel template:\n{error}",
        "save_excel_export": "Save table export",
        "export_saved": "Table export saved:\n{path}",
        "export_error": "Table export error:\n{error}",
        "delete_error": "Row deletion error:\n{error}",
        "message": "Message",
        "ok": "OK",
        "import_mode_title": "Bulk import mode",
        "import_mode_label": "Choose the Excel import mode",
        "import_mode_info": "File:\n{file}\n\nAppend — add rows to the existing data.\nReplace — fully clear the table and then load the Excel file.",
        "append": "Append",
        "replace": "Replace",
        "edit_window_title": "Edit cell",
        "column": "Column",
        "new_value": "New value",
        "save": "Save",
        "save_error": "Save error:\n{error}",
        "add_row_title": "New row",
        "add_row": "Add row",
        "add_row_error": "Row creation error:\n{error}",
        "pk_marker": " *",
        "tables": {
            "circles": "Circles",
            "interactions": "Interactions",
            "meetings": "Meetings",
            "persons": "People",
            "registry_tasks": "Task Registry",
            "responsibles": "Responsibles",
            "tasks": "Tasks",
        },
    },
}


class DatabaseEditorTab(ctk.CTkFrame):
    def __init__(self, parent, editor_service: DatabaseEditorService | None = None, database_editor_service: DatabaseEditorService | None = None):
        super().__init__(parent)
        self.editor_service = editor_service or database_editor_service or DatabaseEditorService()

        self.current_table: str | None = None
        self.columns_meta: list[dict] = []
        self.selected_row_id = None
        self.rows_cache: list[dict] = []
        self.table_name_map: dict[str, str] = {}
        self.reverse_table_name_map: dict[str, str] = {}

        self.configure(fg_color="transparent")

        self._build_filters_row()
        self._build_actions_row()
        self._build_info()
        self._build_table()
        self._load_tables()

    @property
    def current_language(self) -> str:
        lang = getattr(self.winfo_toplevel(), "current_language", "ru")
        return "en" if str(lang).lower() == "en" else "ru"

    def tr(self, key: str, **kwargs) -> str:
        text = TRANSLATIONS[self.current_language].get(key, key)
        if kwargs:
            return text.format(**kwargs)
        return text

    def tr_table(self, table_name: str) -> str:
        return TRANSLATIONS[self.current_language]["tables"].get(table_name, table_name)

    def _build_filters_row(self):
        top = ctk.CTkFrame(self, corner_radius=14)
        top.pack(fill="x", padx=10, pady=(10, 6))

        self.table_label = ctk.CTkLabel(top, text=self.tr("table"), font=ctk.CTkFont(size=12, weight="normal"))
        self.table_label.pack(side="left", padx=(12, 5), pady=10)

        self.table_combo = ctk.CTkComboBox(top, values=[""], width=220, command=self.on_table_change, corner_radius=10)
        self.table_combo.pack(side="left", padx=5, pady=10)

        self.search_label = ctk.CTkLabel(top, text=self.tr("search"), font=ctk.CTkFont(size=12, weight="normal"))
        self.search_label.pack(side="left", padx=(12, 5), pady=10)

        self.search_entry = ctk.CTkEntry(top, width=300, placeholder_text=self.tr("search_placeholder"))
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

    def _build_actions_row(self):
        actions = ctk.CTkFrame(self, corner_radius=14)
        actions.pack(fill="x", padx=10, pady=(0, 10))

        self.refresh_btn = ctk.CTkButton(actions, text=self.tr("refresh"), width=120, command=self.refresh, corner_radius=10, height=34)
        self.refresh_btn.pack(side="left", padx=10, pady=10)

        self.edit_btn = ctk.CTkButton(actions, text=self.tr("edit_cell"), width=170, command=self.open_edit_cell_window, corner_radius=10, height=34)
        self.edit_btn.pack(side="left", padx=5, pady=10)

        self.delete_btn = ctk.CTkButton(actions, text=self.tr("delete"), width=120, command=self.delete_selected_row, corner_radius=10, height=34)
        self.delete_btn.pack(side="left", padx=5, pady=10)

        self.new_row_btn = ctk.CTkButton(actions, text=self.tr("new_row"), width=140, command=self.open_add_row_window, corner_radius=10, height=34)
        self.new_row_btn.pack(side="left", padx=5, pady=10)

        self.export_btn = ctk.CTkButton(actions, text=self.tr("export_excel"), width=190, command=self.export_current_table, corner_radius=10, height=34)
        self.export_btn.pack(side="right", padx=10, pady=10)

        self.import_btn = ctk.CTkButton(actions, text=self.tr("bulk_import"), width=190, command=self.bulk_import_excel, corner_radius=10, height=34)
        self.import_btn.pack(side="right", padx=5, pady=10)

        self.template_btn = ctk.CTkButton(actions, text=self.tr("excel_template"), width=150, command=self.export_template_excel, corner_radius=10, height=34)
        self.template_btn.pack(side="right", padx=5, pady=10)

    def _build_info(self):
        info = ctk.CTkFrame(self, corner_radius=14)
        info.pack(fill="x", padx=10, pady=(0, 10))
        self.info_label = ctk.CTkLabel(info, text=self.tr("choose_table_hint"), font=ctk.CTkFont(size=14, weight="normal"))
        self.info_label.pack(side="left", padx=14, pady=12)

    def _build_table(self):
        card = ctk.CTkFrame(self, corner_radius=14)
        card.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.table_title = ctk.CTkLabel(card, text=self.tr("editor_title"), anchor="w", font=ctk.CTkFont(size=18, weight="normal"))
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
        self.table_name_map = {self.tr_table(name): name for name in tables}
        self.reverse_table_name_map = {v: k for k, v in self.table_name_map.items()}
        display_values = list(self.table_name_map.keys()) or [""]
        self.table_combo.configure(values=display_values)
        if tables:
            first_table = tables[0]
            self.current_table = first_table
            self.table_combo.set(self.reverse_table_name_map.get(first_table, first_table))
            self.refresh()

    def on_table_change(self, selected_value: str):
        self.current_table = self.table_name_map.get(selected_value, selected_value)
        self.search_entry.delete(0, "end")
        self.refresh()

    def refresh(self):
        if not self.current_table:
            return

        self.columns_meta = self.editor_service.get_table_columns(self.current_table)
        self.table_title.configure(text=self.tr("editor_title_with_table", table=self.tr_table(self.current_table)))

        query = self.search_entry.get().strip()
        if query:
            self.rows_cache = self.editor_service.search_rows(self.current_table, query=query)
        else:
            self.rows_cache = self.editor_service.list_rows(self.current_table)

        self._rebuild_tree()
        self.info_label.configure(
            text=self.tr(
                "table_info",
                table=self.tr_table(self.current_table),
                rows=len(self.rows_cache),
                cols=len(self.columns_meta),
            )
        )
        self.selected_row_id = None

    def _rebuild_tree(self):
        columns = [c["name"] for c in self.columns_meta]
        self.tree.configure(columns=columns)
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
            self.tree.column(col, width=120, anchor="w")
        for col_meta in self.columns_meta:
            name = col_meta["name"]
            label = f"{name}{self.tr('pk_marker')}" if col_meta["primary_key"] else name
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
            self.show_message(self.tr("pick_table_first"))
            return
        file_path = filedialog.askopenfilename(
            title=self.tr("select_excel_import"),
            filetypes=[("Excel files", "*.xlsx *.xlsm")],
        )
        if not file_path:
            return
        ImportModeWindow(self, file_path)

    def run_bulk_import(self, file_path: str, mode: str):
        try:
            result = self.editor_service.import_from_excel(self.current_table, file_path=file_path, mode=mode)
            self.refresh()
            text = self.tr(
                "bulk_done",
                table=self.tr_table(self.current_table),
                inserted=result["inserted"],
                skipped=result["skipped"],
                headers=", ".join(result["matched_headers"]),
            )
            if result["errors"]:
                text += self.tr("first_errors", errors="\n".join(result["errors"][:8]))
            self.show_message(text)
        except Exception as exc:
            self.show_message(self.tr("bulk_error", error=exc))

    def export_template_excel(self):
        if not self.current_table:
            self.show_message(self.tr("pick_table_first"))
            return
        initial_name = f"{self.current_table}_template.xlsx"
        output_path = filedialog.asksaveasfilename(
            title=self.tr("save_excel_template"),
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return
        try:
            self.editor_service.create_excel_template(self.current_table, output_path)
            self.show_message(self.tr("template_saved", path=output_path))
        except Exception as exc:
            self.show_message(self.tr("template_error", error=exc))

    def export_current_table(self):
        if not self.current_table:
            self.show_message(self.tr("pick_table_first"))
            return
        initial_name = f"{self.current_table}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title=self.tr("save_excel_export"),
            defaultextension=".xlsx",
            initialfile=initial_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return
        try:
            self.editor_service.export_table_to_excel(self.current_table, output_path, query=self.search_entry.get().strip())
            self.show_message(self.tr("export_saved", path=output_path))
        except Exception as exc:
            self.show_message(self.tr("export_error", error=exc))

    def open_edit_cell_window(self):
        if not self.current_table:
            self.show_message(self.tr("pick_table_first"))
            return
        if self.selected_row_id is None:
            self.show_message(self.tr("pick_row_first"))
            return
        EditCellWindow(self)

    def open_add_row_window(self):
        if not self.current_table:
            self.show_message(self.tr("pick_table_first"))
            return
        AddRowWindow(self)

    def delete_selected_row(self):
        if not self.current_table:
            self.show_message(self.tr("pick_table_first"))
            return
        if self.selected_row_id is None:
            self.show_message(self.tr("pick_row_first"))
            return
        try:
            self.editor_service.delete_row(self.current_table, self.selected_row_id)
            self.refresh()
        except Exception as exc:
            self.show_message(self.tr("delete_error", error=exc))

    def show_message(self, text: str):
        win = ctk.CTkToplevel(self)
        win.geometry("620x320")
        win.title(self.tr("message"))
        win.grab_set()
        box = ctk.CTkTextbox(win, width=560, height=220)
        box.pack(padx=20, pady=20)
        box.insert("1.0", text)
        box.configure(state="disabled")
        ctk.CTkButton(win, text=self.tr("ok"), command=win.destroy, width=140, corner_radius=10, height=34).pack(pady=(0, 10))


class ImportModeWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab: DatabaseEditorTab, file_path: str):
        super().__init__()
        self.parent_tab = parent_tab
        self.file_path = file_path

        self.title(self.parent_tab.tr("import_mode_title"))
        self.geometry("560x300")
        self.resizable(False, False)
        self.grab_set()

        ctk.CTkLabel(self, text=self.parent_tab.tr("import_mode_label"), font=ctk.CTkFont(size=16, weight="normal")).pack(pady=(16, 8))

        info = self.parent_tab.tr("import_mode_info", file=file_path)
        box = ctk.CTkTextbox(self, width=500, height=150)
        box.pack(padx=20, pady=10)
        box.insert("1.0", info)
        box.configure(state="disabled")

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(pady=10)
        ctk.CTkButton(btns, text=self.parent_tab.tr("append"), width=170, corner_radius=10, height=34, command=self.run_append).pack(side="left", padx=6)
        ctk.CTkButton(btns, text=self.parent_tab.tr("replace"), width=170, corner_radius=10, height=34, command=self.run_replace).pack(side="left", padx=6)

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
        self.title(self.parent_tab.tr("edit_window_title"))
        self.geometry("620x300")
        self.resizable(False, False)
        self.grab_set()
        self.build()

    def build(self):
        ctk.CTkLabel(self, text=self.parent_tab.tr("column"), font=ctk.CTkFont(size=12, weight="normal")).pack(pady=(14, 4))
        editable_columns = [c["name"] for c in self.parent_tab.columns_meta if not c["primary_key"]]
        self.column_combo = ctk.CTkComboBox(self, values=editable_columns or [""], width=420, corner_radius=10)
        self.column_combo.pack(pady=(0, 8))
        if editable_columns:
            self.column_combo.set(editable_columns[0])

        ctk.CTkLabel(self, text=self.parent_tab.tr("new_value"), font=ctk.CTkFont(size=12, weight="normal")).pack(pady=(8, 4))
        self.value_text = ctk.CTkTextbox(self, width=520, height=110)
        self.value_text.pack(pady=(0, 10))

        ctk.CTkButton(self, text=self.parent_tab.tr("save"), command=self.save, width=180, corner_radius=10, height=34).pack(pady=8)

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
            self.parent_tab.show_message(self.parent_tab.tr("save_error", error=exc))


class AddRowWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab: DatabaseEditorTab):
        super().__init__()
        self.parent_tab = parent_tab
        self.entries = {}
        self.title(self.parent_tab.tr("add_row_title"))
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

        ctk.CTkButton(self.scroll, text=self.parent_tab.tr("add_row"), command=self.save, width=180, corner_radius=10, height=34).pack(pady=14)

    def save(self):
        values = {name: entry.get().strip() for name, entry in self.entries.items()}
        try:
            self.parent_tab.editor_service.insert_row(self.parent_tab.current_table, values)
            self.parent_tab.refresh()
            self.destroy()
        except Exception as exc:
            self.parent_tab.show_message(self.parent_tab.tr("add_row_error", error=exc))
