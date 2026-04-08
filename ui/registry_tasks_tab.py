import customtkinter as ctk
from tkinter import ttk, filedialog
from tkcalendar import Calendar
from datetime import datetime, date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

from models.registry_task import (
    REGISTRY_TASK_STATUS_VALUES,
    REGISTRY_TASK_STATUS_NEW,
    REGISTRY_TASK_STATUS_IN_PROGRESS,
    REGISTRY_TASK_STATUS_DONE,
    REGISTRY_TASK_STATUS_CANCELLED,
    REGISTRY_TASK_STATUS_OVERDUE,
)


DISPLAY_STATUS_ORDER = {
    REGISTRY_TASK_STATUS_OVERDUE: 0,
    REGISTRY_TASK_STATUS_IN_PROGRESS: 1,
    REGISTRY_TASK_STATUS_NEW: 2,
    REGISTRY_TASK_STATUS_DONE: 3,
    REGISTRY_TASK_STATUS_CANCELLED: 4,
}


class RegistryTasksTab(ctk.CTkFrame):
    def __init__(self, parent, registry_task_service):
        super().__init__(parent)
        self.registry_task_service = registry_task_service
        self.selected_task_id = None

        top_frame = ctk.CTkFrame(self, corner_radius=14)
        top_frame.pack(fill="x", padx=10, pady=(10, 6))

        ctk.CTkButton(top_frame, text="➕ Новое поручение", command=self.open_add_window, width=170, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(top_frame, text="✏️ Редактировать", command=self.open_edit_window, width=150, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(top_frame, text="📄 Дублировать", command=self.duplicate_selected, width=140, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(top_frame, text="🗑 Удалить", command=self.delete_selected, width=120, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(top_frame, text="📑 Сформировать отчёт", command=self.export_report_pdf, width=180, corner_radius=10, height=34).pack(side="left", padx=5, pady=10)
        ctk.CTkButton(top_frame, text="🔄 Обновить", command=self.refresh, width=140, corner_radius=10, height=34).pack(side="right", padx=10, pady=10)

        filter_frame = ctk.CTkFrame(self, corner_radius=14)
        filter_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(filter_frame, text="Поиск", font=ctk.CTkFont(size=12, weight="normal")).pack(side="left", padx=(10, 5), pady=10)
        self.search_entry = ctk.CTkEntry(filter_frame, width=260, placeholder_text="Название / источник / ответственный...")
        self.search_entry.pack(side="left", padx=5, pady=10)
        self.search_entry.bind("<KeyRelease>", lambda event: self.refresh())

        ctk.CTkLabel(filter_frame, text="Статус", font=ctk.CTkFont(size=12, weight="normal")).pack(side="left", padx=(15, 5), pady=10)

        self.status_new_var = ctk.BooleanVar(value=True)
        self.status_progress_var = ctk.BooleanVar(value=True)
        self.status_overdue_var = ctk.BooleanVar(value=True)
        self.status_done_var = ctk.BooleanVar(value=False)
        self.status_cancelled_var = ctk.BooleanVar(value=False)

        self.status_new_cb = ctk.CTkCheckBox(filter_frame, text="Новые", variable=self.status_new_var, command=self.refresh)
        self.status_new_cb.pack(side="left", padx=(2, 4), pady=10)
        self.status_progress_cb = ctk.CTkCheckBox(filter_frame, text="В работе", variable=self.status_progress_var, command=self.refresh)
        self.status_progress_cb.pack(side="left", padx=4, pady=10)
        self.status_overdue_cb = ctk.CTkCheckBox(filter_frame, text="Просроченные", variable=self.status_overdue_var, command=self.refresh)
        self.status_overdue_cb.pack(side="left", padx=4, pady=10)
        self.status_done_cb = ctk.CTkCheckBox(filter_frame, text="Выполненные", variable=self.status_done_var, command=self.refresh)
        self.status_done_cb.pack(side="left", padx=4, pady=10)
        self.status_cancelled_cb = ctk.CTkCheckBox(filter_frame, text="Отменённые", variable=self.status_cancelled_var, command=self.refresh)
        self.status_cancelled_cb.pack(side="left", padx=4, pady=10)

        ctk.CTkLabel(filter_frame, text="Ответственный", font=ctk.CTkFont(size=12, weight="normal")).pack(side="left", padx=(15, 5), pady=10)
        self.responsible_filter = ctk.CTkComboBox(
            filter_frame,
            values=self.registry_task_service.get_responsible_values(),
            width=180,
            command=lambda _: self.refresh(),
            corner_radius=10,
        )
        self.responsible_filter.pack(side="left", padx=5, pady=10)
        self.responsible_filter.set("Все")

        ctk.CTkButton(filter_frame, text="Сбросить фильтры", width=150, command=self.reset_filters, corner_radius=10, height=34).pack(side="right", padx=10, pady=10)

        self.counter_frame = ctk.CTkFrame(self, corner_radius=14)
        self.counter_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.new_label = ctk.CTkLabel(self.counter_frame, text="🟦 Новые: 0", font=ctk.CTkFont(size=16, weight="normal"))
        self.new_label.pack(side="left", padx=16, pady=12)
        self.progress_label = ctk.CTkLabel(self.counter_frame, text="🟨 В работе: 0", font=ctk.CTkFont(size=16, weight="normal"))
        self.progress_label.pack(side="left", padx=16, pady=12)
        self.overdue_label = ctk.CTkLabel(self.counter_frame, text="🟥 Просрочены: 0", font=ctk.CTkFont(size=16, weight="normal"))
        self.overdue_label.pack(side="left", padx=16, pady=12)
        self.done_label = ctk.CTkLabel(self.counter_frame, text="🟩 Выполнены: 0", font=ctk.CTkFont(size=16, weight="normal"))
        self.done_label.pack(side="left", padx=16, pady=12)

        table_card = ctk.CTkFrame(self, corner_radius=14)
        table_card.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        ctk.CTkLabel(table_card, text="Реестр поручений", anchor="w", font=ctk.CTkFont(size=18, weight="normal")).pack(fill="x", padx=14, pady=(12, 8))

        table_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("id", "title", "source", "main_responsible", "controller", "due_date", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        headings = {
            "id": "ID",
            "title": "Поручение",
            "source": "Источник",
            "main_responsible": "Основной",
            "controller": "Контроль",
            "due_date": "Срок",
            "status": "Статус",
        }
        for key, text in headings.items():
            self.tree.heading(key, text=text)

        self.tree.column("id", width=55, anchor="center")
        self.tree.column("title", width=360)
        self.tree.column("source", width=180)
        self.tree.column("main_responsible", width=140)
        self.tree.column("controller", width=140)
        self.tree.column("due_date", width=100, anchor="center")
        self.tree.column("status", width=120, anchor="center")

        self.tree.tag_configure("overdue", background="#ffdddd")
        self.tree.tag_configure("in_progress", background="#fff4cc")
        self.tree.tag_configure("done", background="#ddffdd")

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.on_double_click)

        self.refresh()

    def reset_filters(self):
        self.search_entry.delete(0, "end")
        self.status_new_var.set(True)
        self.status_progress_var.set(True)
        self.status_overdue_var.set(True)
        self.status_done_var.set(False)
        self.status_cancelled_var.set(False)
        self.responsible_filter.configure(values=self.registry_task_service.get_responsible_values())
        self.responsible_filter.set("Все")
        self.refresh()

    def _get_selected_statuses(self):
        statuses = []
        if self.status_new_var.get():
            statuses.append(REGISTRY_TASK_STATUS_NEW)
        if self.status_progress_var.get():
            statuses.append(REGISTRY_TASK_STATUS_IN_PROGRESS)
        if self.status_overdue_var.get():
            statuses.append(REGISTRY_TASK_STATUS_OVERDUE)
        if self.status_done_var.get():
            statuses.append(REGISTRY_TASK_STATUS_DONE)
        if self.status_cancelled_var.get():
            statuses.append(REGISTRY_TASK_STATUS_CANCELLED)
        return statuses

    def _row_sort_key(self, row):
        task = row["task"]
        status = row["status"]
        due = task.due_date.date() if task.due_date else date.max
        return (DISPLAY_STATUS_ORDER.get(status, 99), due, (task.title or "").lower())

    def _get_filtered_rows(self):
        rows = self.registry_task_service.list_tasks(
            search=self.search_entry.get().strip(),
            status="Все",
            responsible=self.responsible_filter.get().strip(),
        )
        selected_statuses = self._get_selected_statuses()
        if selected_statuses:
            rows = [row for row in rows if row["status"] in selected_statuses]
        rows = sorted(rows, key=self._row_sort_key)
        return rows

    def refresh(self):
        values = self.registry_task_service.get_responsible_values()
        self.responsible_filter.configure(values=values)
        if self.responsible_filter.get() not in values:
            self.responsible_filter.set("Все")

        for row in self.tree.get_children():
            self.tree.delete(row)

        rows = self._get_filtered_rows()

        counts = self.registry_task_service.get_status_counts()
        self.new_label.configure(text=f"🟦 Новые: {counts.get(REGISTRY_TASK_STATUS_NEW, 0)}")
        self.progress_label.configure(text=f"🟨 В работе: {counts.get(REGISTRY_TASK_STATUS_IN_PROGRESS, 0)}")
        self.overdue_label.configure(text=f"🟥 Просрочены: {counts.get(REGISTRY_TASK_STATUS_OVERDUE, 0)}")
        self.done_label.configure(text=f"🟩 Выполнены: {counts.get(REGISTRY_TASK_STATUS_DONE, 0)}")

        for row in rows:
            task = row["task"]
            status = row["status"]
            due_date = task.due_date.strftime("%d.%m.%Y") if task.due_date else ""

            tag = ""
            if status == REGISTRY_TASK_STATUS_OVERDUE:
                tag = "overdue"
            elif status == REGISTRY_TASK_STATUS_IN_PROGRESS:
                tag = "in_progress"
            elif status == REGISTRY_TASK_STATUS_DONE:
                tag = "done"

            self.tree.insert(
                "",
                "end",
                values=(
                    task.id,
                    task.title or "",
                    task.source or "",
                    task.main_responsible or "",
                    task.controller or "",
                    due_date,
                    status,
                ),
                tags=(tag,) if tag else (),
            )

        self.selected_task_id = None

    def export_report_pdf(self):
        rows = self._get_filtered_rows()
        if not rows:
            self.show_message("По текущим фильтрам нет поручений для выгрузки в PDF")
            return

        initial_name = f"реестр_поручений_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        output_path = filedialog.asksaveasfilename(
            title="Сохранить отчёт PDF",
            defaultextension=".pdf",
            initialfile=initial_name,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output_path:
            return

        try:
            self._build_pdf_report(Path(output_path), rows)
            self.show_message(f"PDF-отчёт сформирован:\n{output_path}")
        except Exception as exc:
            self.show_message(f"Ошибка формирования PDF:\n{exc}")

    def _build_pdf_report(self, output_path: Path, rows):
        font_name = self._ensure_pdf_font()
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=landscape(A4),
            leftMargin=14 * mm,
            rightMargin=14 * mm,
            topMargin=12 * mm,
            bottomMargin=12 * mm,
        )

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "ReportTitle",
            parent=styles["Normal"],
            fontName=font_name,
            fontSize=15,
            leading=18,
            alignment=TA_LEFT,
            spaceAfter=4,
        )
        meta_style = ParagraphStyle(
            "ReportMeta",
            parent=styles["Normal"],
            fontName=font_name,
            fontSize=9,
            leading=11,
            alignment=TA_LEFT,
            textColor=colors.HexColor("#555555"),
            spaceAfter=2,
        )
        table_style_text = ParagraphStyle(
            "Cell",
            parent=styles["Normal"],
            fontName=font_name,
            fontSize=8,
            leading=10,
            alignment=TA_LEFT,
        )
        table_style_center = ParagraphStyle(
            "CellCenter",
            parent=table_style_text,
            alignment=TA_CENTER,
        )

        elements = []
        elements.append(Paragraph("Реестр поручений — отчёт по применённым фильтрам", title_style))
        elements.append(Paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}", meta_style))
        elements.append(Paragraph(f"Поиск: {self.search_entry.get().strip() or '—'}", meta_style))
        elements.append(Paragraph(f"Ответственный: {self.responsible_filter.get().strip() or 'Все'}", meta_style))
        elements.append(Paragraph(f"Статусы: {self._selected_statuses_text()}", meta_style))
        elements.append(Paragraph(f"Количество поручений: {len(rows)}", meta_style))
        elements.append(Spacer(1, 6))

        data = [[
            Paragraph("ID", table_style_center),
            Paragraph("Поручение", table_style_text),
            Paragraph("Источник", table_style_text),
            Paragraph("Основной", table_style_text),
            Paragraph("Контроль", table_style_text),
            Paragraph("Срок", table_style_center),
            Paragraph("Статус", table_style_center),
        ]]

        for row in rows:
            task = row["task"]
            due = task.due_date.strftime("%d.%m.%Y") if task.due_date else ""
            data.append([
                Paragraph(str(task.id), table_style_center),
                Paragraph(self._safe(task.title), table_style_text),
                Paragraph(self._safe(task.source), table_style_text),
                Paragraph(self._safe(task.main_responsible), table_style_text),
                Paragraph(self._safe(task.controller), table_style_text),
                Paragraph(due, table_style_center),
                Paragraph(self._safe(row["status"]), table_style_center),
            ])

        table = Table(
            data,
            colWidths=[18 * mm, 78 * mm, 48 * mm, 35 * mm, 35 * mm, 24 * mm, 28 * mm],
            repeatRows=1,
        )
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DCE6F1")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("LEADING", (0, 0), (-1, -1), 10),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9AA5B1")),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#7B8794")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))

        for row_index, row in enumerate(rows, start=1):
            status = row["status"]
            if status == REGISTRY_TASK_STATUS_OVERDUE:
                bg = colors.HexColor("#FDE2E1")
            elif status == REGISTRY_TASK_STATUS_IN_PROGRESS:
                bg = colors.HexColor("#FFF3CD")
            elif status == REGISTRY_TASK_STATUS_DONE:
                bg = colors.HexColor("#DFF4E4")
            elif status == REGISTRY_TASK_STATUS_CANCELLED:
                bg = colors.HexColor("#ECECEC")
            else:
                bg = colors.white
            table.setStyle(TableStyle([("BACKGROUND", (0, row_index), (-1, row_index), bg)]))

        elements.append(table)
        doc.build(elements)

    def _selected_statuses_text(self):
        mapping = [
            (self.status_new_var.get(), "Новые"),
            (self.status_progress_var.get(), "В работе"),
            (self.status_overdue_var.get(), "Просроченные"),
            (self.status_done_var.get(), "Выполненные"),
            (self.status_cancelled_var.get(), "Отменённые"),
        ]
        selected = [label for is_on, label in mapping if is_on]
        return ", ".join(selected) if selected else "ничего не выбрано"

    def _safe(self, value):
        text = (value or "").strip()
        if not text:
            return "—"
        text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return text

    def _ensure_pdf_font(self):
        font_name = "ArialUnicodeReport"
        if font_name in pdfmetrics.getRegisteredFontNames():
            return font_name

        candidates = [
            Path("C:/Windows/Fonts/arial.ttf"),
            Path("C:/Windows/Fonts/ARIAL.TTF"),
            Path("C:/Windows/Fonts/calibri.ttf"),
            Path("C:/Windows/Fonts/CALIBRI.TTF"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ]
        for path in candidates:
            if path.exists():
                pdfmetrics.registerFont(TTFont(font_name, str(path)))
                return font_name

        fallback = "Helvetica"
        return fallback

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
        self.open_edit_window()

    def open_add_window(self):
        RegistryTaskWindow(self, mode="add")

    def open_edit_window(self):
        if self.selected_task_id is None:
            self.show_message("Выбери поручение")
            return
        RegistryTaskWindow(self, mode="edit", task_id=self.selected_task_id)

    def duplicate_selected(self):
        if self.selected_task_id is None:
            self.show_message("Выбери поручение")
            return
        task = self.registry_task_service.get_task(self.selected_task_id)
        if not task:
            self.show_message("Поручение не найдено")
            return
        payload = {
            "title": f"{(task.title or '').strip()} (копия)",
            "description": task.description or "",
            "source": task.source or "",
            "main_responsible": task.main_responsible or "",
            "co_executors": task.co_executors or "",
            "controller": task.controller or "",
            "due_date": task.due_date,
            "status": REGISTRY_TASK_STATUS_NEW,
            "comment": task.comment or "",
        }
        try:
            self.registry_task_service.create_task(payload)
        except ValueError as e:
            self.show_message(str(e))
            return
        self.refresh()

    def delete_selected(self):
        if self.selected_task_id is None:
            self.show_message("Выбери поручение")
            return
        deleted = self.registry_task_service.delete_task(self.selected_task_id)
        if not deleted:
            self.show_message("Поручение не найдено")
            return
        self.refresh()

    def show_message(self, text):
        win = ctk.CTkToplevel(self)
        win.geometry("460x180")
        win.title("Сообщение")
        win.grab_set()
        ctk.CTkLabel(win, text=text, wraplength=420, justify="left").pack(pady=20, padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy, width=140, corner_radius=10, height=34).pack(pady=10)


class RegistryTaskWindow(ctk.CTkToplevel):
    def __init__(self, parent_tab, mode="add", task_id=None):
        super().__init__()
        self.parent_tab = parent_tab
        self.mode = mode
        self.task_id = task_id
        self.task = None
        if mode == "edit":
            self.task = self.parent_tab.registry_task_service.get_task(task_id)

        self.title("Новое поручение" if mode == "add" else "Редактирование поручения")
        self.geometry("980x760")
        self.minsize(900, 680)
        self.resizable(True, True)
        self.grab_set()

        self.build()

    def build(self):
        form_width = 760

        self.scroll = ctk.CTkScrollableFrame(self, width=900, height=660, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=10, pady=10)

        self.title_label = ctk.CTkLabel(self.scroll, text="Название поручения", font=ctk.CTkFont(size=12, weight="normal"))
        self.title_label.pack(pady=(8, 3))
        self.title_entry = ctk.CTkEntry(self.scroll, width=form_width)
        self.title_entry.pack(pady=(0, 6))

        self.source_label = ctk.CTkLabel(self.scroll, text="Источник / кто поручил", font=ctk.CTkFont(size=12, weight="normal"))
        self.source_label.pack(pady=(4, 2))
        self.source_entry = ctk.CTkEntry(self.scroll, width=form_width)
        self.source_entry.pack(pady=(0, 6))

        self.main_responsible_label = ctk.CTkLabel(self.scroll, text="Основной ответственный", font=ctk.CTkFont(size=12, weight="normal"))
        self.main_responsible_label.pack(pady=(4, 2))
        self.main_responsible_combo = ctk.CTkComboBox(
            self.scroll,
            values=[v for v in self.parent_tab.registry_task_service.get_responsible_values() if v != "Все"] or [""],
            width=form_width,
            corner_radius=10,
        )
        self.main_responsible_combo.pack(pady=(0, 6))

        self.co_exec_label = ctk.CTkLabel(self.scroll, text="Соисполнители", font=ctk.CTkFont(size=12, weight="normal"))
        self.co_exec_label.pack(pady=(4, 2))
        self.co_exec_entry = ctk.CTkEntry(self.scroll, width=form_width, placeholder_text="Через запятую")
        self.co_exec_entry.pack(pady=(0, 6))

        self.controller_label = ctk.CTkLabel(self.scroll, text="Контроль", font=ctk.CTkFont(size=12, weight="normal"))
        self.controller_label.pack(pady=(4, 2))
        self.controller_entry = ctk.CTkEntry(self.scroll, width=form_width)
        self.controller_entry.pack(pady=(0, 6))

        self.due_date_label = ctk.CTkLabel(self.scroll, text="Срок", font=ctk.CTkFont(size=12, weight="normal"))
        self.due_date_label.pack(pady=(4, 2))
        due_frame = ctk.CTkFrame(self.scroll, fg_color="transparent")
        due_frame.pack(pady=(0, 6))
        self.due_date_entry = ctk.CTkEntry(due_frame, width=460, placeholder_text="дд.мм.гггг")
        self.due_date_entry.pack(side="left", padx=5)
        ctk.CTkButton(due_frame, text="📅 Выбрать", width=120, command=lambda: self.open_calendar(self.due_date_entry), corner_radius=10, height=34).pack(side="left", padx=5)

        self.status_label = ctk.CTkLabel(self.scroll, text="Статус", font=ctk.CTkFont(size=12, weight="normal"))
        self.status_label.pack(pady=(4, 2))
        self.status_combo = ctk.CTkComboBox(self.scroll, values=REGISTRY_TASK_STATUS_VALUES, width=form_width, corner_radius=10)
        self.status_combo.pack(pady=(0, 6))
        self.status_combo.set(REGISTRY_TASK_STATUS_NEW)

        self.description_label = ctk.CTkLabel(self.scroll, text="Описание", font=ctk.CTkFont(size=12, weight="normal"))
        self.description_label.pack(pady=(4, 2))
        self.description_text = ctk.CTkTextbox(self.scroll, width=form_width, height=100)
        self.description_text.pack(pady=(0, 6))

        self.comment_label = ctk.CTkLabel(self.scroll, text="Комментарий", font=ctk.CTkFont(size=12, weight="normal"))
        self.comment_label.pack(pady=(4, 2))
        self.comment_text = ctk.CTkTextbox(self.scroll, width=form_width, height=85)
        self.comment_text.pack(pady=(0, 8))

        ctk.CTkButton(self.scroll, text="Сохранить", command=self.save, width=220, corner_radius=10, height=34).pack(pady=(8, 14))

        if self.task:
            self.fill()

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
        ctk.CTkButton(btn_frame, text="OK", width=100, command=apply, corner_radius=10, height=34).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Отмена", width=100, command=win.destroy, corner_radius=10, height=34).pack(side="left", padx=5)

    def fill(self):
        self.title_entry.insert(0, self.task.title or "")
        self.source_entry.insert(0, self.task.source or "")
        self.main_responsible_combo.set(self.task.main_responsible or "")
        self.co_exec_entry.insert(0, self.task.co_executors or "")
        self.controller_entry.insert(0, self.task.controller or "")
        if self.task.due_date:
            self.due_date_entry.insert(0, self.task.due_date.strftime("%d.%m.%Y"))
        self.status_combo.set(self.task.status or REGISTRY_TASK_STATUS_NEW)
        self.description_text.insert("1.0", self.task.description or "")
        self.comment_text.insert("1.0", self.task.comment or "")

    def save(self):
        due_date = self.parse_date(self.due_date_entry.get())
        if due_date == "INVALID":
            self.show_error("Срок должен быть в формате дд.мм.гггг")
            return

        payload = {
            "title": self.title_entry.get().strip(),
            "source": self.source_entry.get().strip(),
            "main_responsible": self.main_responsible_combo.get().strip(),
            "co_executors": self.co_exec_entry.get().strip(),
            "controller": self.controller_entry.get().strip(),
            "due_date": due_date,
            "status": self.status_combo.get().strip(),
            "description": self.description_text.get("1.0", "end").strip(),
            "comment": self.comment_text.get("1.0", "end").strip(),
        }

        try:
            if self.mode == "edit":
                task = self.parent_tab.registry_task_service.update_task(self.task_id, payload)
                if not task:
                    self.show_error("Поручение не найдено")
                    return
            else:
                self.parent_tab.registry_task_service.create_task(payload)
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
        ctk.CTkButton(win, text="OK", command=win.destroy, width=140, corner_radius=10, height=34).pack(pady=10)
