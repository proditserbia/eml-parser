"""Main Tkinter GUI for the EML parser application."""

import dataclasses
import json
import logging
import os
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from exporters import ExcelExporter, PdfExporter
from models import ParsedReport
from parser import EmlParser

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Colour / style constants (dark theme)
# ---------------------------------------------------------------------------

_BG = "#1E1E2E"          # main background
_BG2 = "#2A2A3E"         # slightly lighter bg (frames, panels)
_BG3 = "#313145"         # tree / table background
_ACCENT = "#4A9EFF"      # buttons, highlights
_ACCENT_HOVER = "#6AB4FF"
_FG = "#D4D4F0"          # primary text
_FG_DIM = "#8888AA"      # secondary text / placeholders
_SEL_BG = "#3A5A8A"      # selection background
_SEL_FG = "#FFFFFF"
_LOG_BG = "#111120"
_LOG_FG = "#88FFAA"
_BTN_BG = "#2D6A9F"
_BTN_FG = "#FFFFFF"
_FONT_FAMILY = "Segoe UI"   # good Cyrillic coverage on Windows; falls back gracefully


class _HoverButton(tk.Button):
    """A Tkinter Button with a simple hover colour change."""

    def __init__(self, master: tk.Widget, **kwargs) -> None:  # type: ignore[override]
        self._bg_normal = kwargs.get("background", _BTN_BG)
        self._bg_hover = kwargs.pop("hover_bg", _ACCENT_HOVER)
        super().__init__(master, **kwargs)
        self.bind("<Enter>", lambda _e: self.configure(background=self._bg_hover))
        self.bind("<Leave>", lambda _e: self.configure(background=self._bg_normal))


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow:
    """Top-level application window."""

    def __init__(self, root: tk.Tk) -> None:
        self._root = root
        self._parser = EmlParser()
        self._excel = ExcelExporter()
        self._pdf = PdfExporter()
        self._report: Optional[ParsedReport] = None
        self._all_active: List[str] = []
        self._all_old: List[str] = []

        self._build_root()
        self._build_toolbar()
        self._build_main()
        self._build_log_panel()
        self._apply_treeview_style()

        self._log("Добро дошли. Учитајте EML датотеку да бисте почели.")

    # ------------------------------------------------------------------
    # Root window setup
    # ------------------------------------------------------------------

    def _build_root(self) -> None:
        self._root.title("ELEMENTS EML Парсер")
        self._root.configure(bg=_BG)
        self._root.geometry("1100x780")
        self._root.minsize(800, 560)
        self._root.columnconfigure(0, weight=1)
        self._root.rowconfigure(1, weight=1)
        self._root.rowconfigure(2, weight=0)

    # ------------------------------------------------------------------
    # Toolbar
    # ------------------------------------------------------------------

    def _build_toolbar(self) -> None:
        toolbar = tk.Frame(self._root, bg=_BG2, pady=6, padx=8)
        toolbar.grid(row=0, column=0, sticky="ew")

        btn_cfg = dict(
            font=(_FONT_FAMILY, 10, "bold"),
            fg=_BTN_FG,
            background=_BTN_BG,
            activebackground=_ACCENT_HOVER,
            activeforeground=_BTN_FG,
            relief="flat",
            padx=12,
            pady=5,
            cursor="hand2",
            borderwidth=0,
        )

        _HoverButton(toolbar, text="📂  Учитај EML",
                     command=self._load_eml, **btn_cfg).pack(side="left", padx=4)
        _HoverButton(toolbar, text="🗑  Очисти",
                     command=self._clear, **btn_cfg).pack(side="left", padx=4)

        sep = tk.Frame(toolbar, width=2, bg=_FG_DIM)
        sep.pack(side="left", padx=8, fill="y", pady=2)

        _HoverButton(toolbar, text="📊  Извези у Excel",
                     command=self._export_excel, **btn_cfg).pack(side="left", padx=4)
        _HoverButton(toolbar, text="📄  Извези у PDF",
                     command=self._export_pdf, **btn_cfg).pack(side="left", padx=4)
        _HoverButton(toolbar, text="💾  Сачувај JSON",
                     command=self._export_json, **btn_cfg).pack(side="left", padx=4)

    # ------------------------------------------------------------------
    # Main content area
    # ------------------------------------------------------------------

    def _build_main(self) -> None:
        content = tk.Frame(self._root, bg=_BG)
        content.grid(row=1, column=0, sticky="nsew", padx=10, pady=(6, 0))
        content.columnconfigure(1, weight=1)
        content.rowconfigure(1, weight=1)

        # Left column: metadata + storage
        left = tk.Frame(content, bg=_BG)
        left.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 8))
        left.columnconfigure(0, weight=1)
        self._build_metadata_frame(left)
        self._build_storage_frame(left)

        # Right column: search + notebook
        right = tk.Frame(content, bg=_BG)
        right.grid(row=0, column=1, rowspan=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        self._build_search_bar(right)
        self._build_notebook(right)

    # ------------------------------------------------------------------
    # Metadata frame
    # ------------------------------------------------------------------

    def _build_metadata_frame(self, parent: tk.Widget) -> None:
        frame = tk.LabelFrame(
            parent,
            text=" Метаподаци ",
            bg=_BG2,
            fg=_ACCENT,
            font=(_FONT_FAMILY, 10, "bold"),
            bd=1,
            relief="solid",
        )
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        frame.columnconfigure(1, weight=1)

        fields = [
            ("Пошиљалац:", "_lbl_from"),
            ("Прималац:", "_lbl_to"),
            ("Датум:", "_lbl_date"),
            ("Наслов:", "_lbl_subject"),
        ]
        for row_idx, (label, attr) in enumerate(fields):
            tk.Label(frame, text=label, bg=_BG2, fg=_FG_DIM,
                     font=(_FONT_FAMILY, 9), anchor="e", width=12
                     ).grid(row=row_idx, column=0, padx=(8, 4), pady=3, sticky="e")
            var = tk.StringVar(value="—")
            lbl = tk.Label(frame, textvariable=var, bg=_BG2, fg=_FG,
                           font=(_FONT_FAMILY, 9), anchor="w",
                           wraplength=260, justify="left")
            lbl.grid(row=row_idx, column=1, padx=(0, 8), pady=3, sticky="ew")
            setattr(self, attr, var)

    # ------------------------------------------------------------------
    # Storage summary frame
    # ------------------------------------------------------------------

    def _build_storage_frame(self, parent: tk.Widget) -> None:
        frame = tk.LabelFrame(
            parent,
            text=" Коришћење складишта ",
            bg=_BG2,
            fg=_ACCENT,
            font=(_FONT_FAMILY, 10, "bold"),
            bd=1,
            relief="solid",
        )
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        frame.columnconfigure(1, weight=1)

        fields = [
            ("Волумен:", "_lbl_volume"),
            ("Слободан простор:", "_lbl_free"),
            ("Искоришћено %:", "_lbl_used"),
            ("Промена %:", "_lbl_change"),
        ]
        for row_idx, (label, attr) in enumerate(fields):
            tk.Label(frame, text=label, bg=_BG2, fg=_FG_DIM,
                     font=(_FONT_FAMILY, 9), anchor="e", width=18
                     ).grid(row=row_idx, column=0, padx=(8, 4), pady=3, sticky="e")
            var = tk.StringVar(value="—")
            lbl = tk.Label(frame, textvariable=var, bg=_BG2, fg=_FG,
                           font=(_FONT_FAMILY, 9), anchor="w")
            lbl.grid(row=row_idx, column=1, padx=(0, 8), pady=3, sticky="ew")
            setattr(self, attr, var)

    # ------------------------------------------------------------------
    # Search / filter bar
    # ------------------------------------------------------------------

    def _build_search_bar(self, parent: tk.Widget) -> None:
        bar = tk.Frame(parent, bg=_BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        bar.columnconfigure(1, weight=1)

        tk.Label(bar, text="🔍  Претрага:", bg=_BG, fg=_FG_DIM,
                 font=(_FONT_FAMILY, 9)).grid(row=0, column=0, padx=(0, 6), sticky="w")

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filter())

        entry = tk.Entry(
            bar,
            textvariable=self._search_var,
            bg=_BG3,
            fg=_FG,
            insertbackground=_FG,
            font=(_FONT_FAMILY, 10),
            relief="flat",
            bd=4,
        )
        entry.grid(row=0, column=1, sticky="ew", ipady=4)

    # ------------------------------------------------------------------
    # Notebook with two workspace tables
    # ------------------------------------------------------------------

    def _build_notebook(self, parent: tk.Widget) -> None:
        notebook = ttk.Notebook(parent)
        notebook.grid(row=1, column=0, sticky="nsew")

        self._tab_active = tk.Frame(notebook, bg=_BG3)
        self._tab_old = tk.Frame(notebook, bg=_BG3)

        notebook.add(self._tab_active, text="  Активни  ")
        notebook.add(self._tab_old, text="  Старији од 90 дана  ")

        self._tree_active = self._make_tree(self._tab_active)
        self._tree_old = self._make_tree(self._tab_old)

    def _make_tree(self, parent: tk.Widget) -> ttk.Treeview:
        frame = tk.Frame(parent, bg=_BG3)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        columns = ("#", "Назив")
        tree = ttk.Treeview(frame, columns=columns, show="headings",
                            selectmode="extended", style="Dark.Treeview")
        tree.heading("#", text="#")
        tree.heading("Назив", text="Назив радног простора")
        tree.column("#", width=50, minwidth=40, anchor="center", stretch=False)
        tree.column("Назив", width=500, minwidth=200, anchor="w")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        return tree

    # ------------------------------------------------------------------
    # Log panel
    # ------------------------------------------------------------------

    def _build_log_panel(self) -> None:
        frame = tk.Frame(self._root, bg=_LOG_BG, height=110)
        frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(4, 8))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        frame.grid_propagate(False)

        self._log_text = tk.Text(
            frame,
            bg=_LOG_BG,
            fg=_LOG_FG,
            font=(_FONT_FAMILY, 8),
            relief="flat",
            state="disabled",
            wrap="word",
        )
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=vsb.set)
        self._log_text.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

    # ------------------------------------------------------------------
    # Treeview dark style
    # ------------------------------------------------------------------

    def _apply_treeview_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Dark.Treeview",
            background=_BG3,
            fieldbackground=_BG3,
            foreground=_FG,
            rowheight=22,
            font=(_FONT_FAMILY, 9),
        )
        style.configure(
            "Dark.Treeview.Heading",
            background=_BTN_BG,
            foreground=_BTN_FG,
            font=(_FONT_FAMILY, 9, "bold"),
            relief="flat",
        )
        style.map(
            "Dark.Treeview",
            background=[("selected", _SEL_BG)],
            foreground=[("selected", _SEL_FG)],
        )

        # Notebook tabs
        style.configure(
            "TNotebook",
            background=_BG,
            borderwidth=0,
        )
        style.configure(
            "TNotebook.Tab",
            background=_BG2,
            foreground=_FG,
            font=(_FONT_FAMILY, 9, "bold"),
            padding=(10, 4),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", _BTN_BG)],
            foreground=[("selected", _BTN_FG)],
        )

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _load_eml(self) -> None:
        path = filedialog.askopenfilename(
            title="Одаберите EML датотеку",
            filetypes=[("EML датотеке", "*.eml"), ("Све датотеке", "*.*")],
        )
        if not path:
            return

        self._log(f"Учитавање: {path}")
        try:
            report = self._parser.parse_file(path)
        except FileNotFoundError as exc:
            self._error(str(exc))
            return
        except ValueError as exc:
            self._error(f"Грешка при парсирању: {exc}")
            return
        except Exception as exc:
            self._error(f"Неочекивана грешка: {exc}")
            return

        source = self._parser.last_source
        if source == "html":
            self._log("⚠ Равни текст није пронађен; коришћено HTML уклањање форматирања.")

        self._report = report
        self._all_active = list(report.active_workspaces)
        self._all_old = list(report.old_workspaces)

        self._populate(report)
        self._log(
            f"✓ Учитано: {len(report.active_workspaces)} активних, "
            f"{len(report.old_workspaces)} старих радних простора."
        )

    def _clear(self) -> None:
        self._report = None
        self._all_active = []
        self._all_old = []
        for attr in ("_lbl_from", "_lbl_to", "_lbl_date", "_lbl_subject",
                     "_lbl_volume", "_lbl_free", "_lbl_used", "_lbl_change"):
            getattr(self, attr).set("—")
        self._tree_active.delete(*self._tree_active.get_children())
        self._tree_old.delete(*self._tree_old.get_children())
        self._search_var.set("")
        self._log("Подаци очишћени.")

    def _export_excel(self) -> None:
        if not self._report:
            self._warn("Нема учитаних података. Прво учитајте EML датотеку.")
            return
        path = filedialog.asksaveasfilename(
            title="Сачувај Excel извештај",
            defaultextension=".xlsx",
            filetypes=[("Excel датотеке", "*.xlsx")],
            initialfile=f"eml_report_{_timestamp()}.xlsx",
        )
        if not path:
            return
        try:
            self._excel.export(self._report, path)
            self._log(f"✓ Excel извештај сачуван: {path}")
        except Exception as exc:
            self._error(f"Грешка при извозу у Excel: {exc}")

    def _export_pdf(self) -> None:
        if not self._report:
            self._warn("Нема учитаних података. Прво учитајте EML датотеку.")
            return
        path = filedialog.asksaveasfilename(
            title="Сачувај PDF извештај",
            defaultextension=".pdf",
            filetypes=[("PDF датотеке", "*.pdf")],
            initialfile=f"eml_report_{_timestamp()}.pdf",
        )
        if not path:
            return
        try:
            self._pdf.export(self._report, path)
            self._log(f"✓ PDF извештај сачуван: {path}")
        except Exception as exc:
            self._error(f"Грешка при извозу у PDF: {exc}")

    def _export_json(self) -> None:
        if not self._report:
            self._warn("Нема учитаних података. Прво учитајте EML датотеку.")
            return
        path = filedialog.asksaveasfilename(
            title="Сачувај JSON датотеку",
            defaultextension=".json",
            filetypes=[("JSON датотеке", "*.json")],
            initialfile=f"eml_report_{_timestamp()}.json",
        )
        if not path:
            return
        try:
            data = dataclasses.asdict(self._report)
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
            self._log(f"✓ JSON датотека сачувана: {path}")
        except Exception as exc:
            self._error(f"Грешка при чувању JSON: {exc}")

    # ------------------------------------------------------------------
    # Populate UI from report
    # ------------------------------------------------------------------

    def _populate(self, report: ParsedReport) -> None:
        self._lbl_from.set(report.from_addr or "—")
        self._lbl_to.set(report.to_addr or "—")
        self._lbl_date.set(report.date or "—")
        self._lbl_subject.set(report.subject or "—")

        self._lbl_volume.set(report.volume_name or "—")
        self._lbl_free.set(report.free_space or "—")
        self._lbl_used.set(report.used_percent or "—")
        self._lbl_change.set(report.change_percent or "—")

        self._fill_tree(self._tree_active, report.active_workspaces)
        self._fill_tree(self._tree_old, report.old_workspaces)

    def _fill_tree(self, tree: ttk.Treeview, items: List[str]) -> None:
        tree.delete(*tree.get_children())
        for i, name in enumerate(items, start=1):
            tag = "odd" if i % 2 else "even"
            tree.insert("", "end", values=(i, name), tags=(tag,))
        tree.tag_configure("odd", background=_BG3)
        tree.tag_configure("even", background=_BG2)

    # ------------------------------------------------------------------
    # Filter
    # ------------------------------------------------------------------

    def _apply_filter(self) -> None:
        query = self._search_var.get().strip().lower()
        active = [w for w in self._all_active if query in w.lower()] if query else self._all_active
        old = [w for w in self._all_old if query in w.lower()] if query else self._all_old
        self._fill_tree(self._tree_active, active)
        self._fill_tree(self._tree_old, old)

    # ------------------------------------------------------------------
    # Logging helpers
    # ------------------------------------------------------------------

    def _log(self, msg: str) -> None:
        self._append_log(f"[{_now()}]  {msg}")
        logger.info(msg)

    def _warn(self, msg: str) -> None:
        self._append_log(f"[{_now()}] ⚠  {msg}")
        messagebox.showwarning("Упозорење", msg, parent=self._root)

    def _error(self, msg: str) -> None:
        self._append_log(f"[{_now()}] ✗  {msg}")
        logger.error(msg)
        messagebox.showerror("Грешка", msg, parent=self._root)

    def _append_log(self, text: str) -> None:
        self._log_text.configure(state="normal")
        self._log_text.insert("end", text + "\n")
        self._log_text.see("end")
        self._log_text.configure(state="disabled")


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _now() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")
