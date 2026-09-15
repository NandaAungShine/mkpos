import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from tkinter.filedialog import asksaveasfilename
from tkinter import ttk
from datetime import date
import tempfile
import os
import inspect
import json

from tkcalendar import Calendar
from tkcalendar import DateEntry
# database
from real.dblogin import Registerdata
from real.dbWarehouse import warehouse
from real.dbTotal_TempReal import DatabaseTotal
from real.dbSalesReportReal import DatabaseSales
import random
import customtkinter as ctk
from tkinter import ttk, messagebox
import sqlite3
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime


# ============================================================
# CUSTOMTKINTER SETTINGS
# ============================================================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ============================================================
# LOAD MYANMAR FONTS (before any UI creation)
# ============================================================

from font_loader import load_bundled_fonts

# This must run BEFORE any ctk.CTkFont() is created
MYANMAR_FONT = load_bundled_fonts()
MYANMAR_FONT_BOLD = MYANMAR_FONT  # Same family, use weight="bold" instead

print(f"🇲🇲 Myanmar font family: {MYANMAR_FONT}")


# ============================================================
# EXCEL EXPORT
# ============================================================

def export_to_excel(treeview, default_filename):

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=default_filename
    )

    if not file_path:
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Report"
        ws.append(list(treeview["columns"]))
        for item in treeview.get_children():
            ws.append(list(treeview.item(item)["values"]))
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        wb.save(file_path)
        messagebox.showinfo("Success", f"Data successfully exported to:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export data: {e}")


# ============================================================
# ROBUST DB INSERT HELPER
# ============================================================

def robust_db_insert(db_obj, *args):
    """Try db.insert() with fewer args if the method rejects them."""

    if db_obj is None:
        raise RuntimeError("Database object is None.")

    try:
        sig = inspect.signature(db_obj.insert)
        params = [
            p for p in sig.parameters.values()
            if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)
        ]
        if params and params[0].name in ("self", "cls"):
            params = params[1:]
        max_args = len(params)
    except Exception:
        max_args = len(args)

    candidates = []
    if len(args) <= max_args:
        candidates.append(args)
    for n in range(min(len(args), max_args), 0, -1):
        cand = args[:n]
        if cand not in candidates:
            candidates.append(cand)

    last_error = None
    for cand in candidates:
        try:
            db_obj.insert(*cand)
            print(f"✅ insert succeeded with {len(cand)} args")
            return True
        except TypeError as e:
            msg = str(e)
            if "positional arguments" in msg or "argument" in msg.lower():
                last_error = e
                continue
            else:
                raise

    if last_error:
        raise last_error
    return False


# ============================================================
# SAFE TABLE INFO (get columns of a table)
# ============================================================

def get_table_columns(cur, table_name):
    """Return list of column names for a table. Empty if error."""
    try:
        cur.execute(f"PRAGMA table_info({table_name})")
        return [row[1] for row in cur.fetchall()]
    except Exception:
        return []


# ============================================================
# DATA PERSISTENCE
# ============================================================

class DataPersistence:

    def __init__(self, db_path="real/KMPOS_data.db"):
        self.db_path = db_path
        try:
            self.con = sqlite3.connect(db_path)
            self.cur = self.con.cursor()
            self.cur.execute("""
                CREATE TABLE IF NOT EXISTS pos_data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    data_key TEXT NOT NULL,
                    row_id INTEGER NOT NULL,
                    data TEXT NOT NULL
                )
            """)
            self.con.commit()
            print(f"✅ DataPersistence initialized: {db_path}")
        except Exception as e:
            print(f"❌ DataPersistence init failed: {e}")
            self.con = None
            self.cur = None

    def load_all(self, data_key):
        if not self.cur:
            return []
        try:
            rows = self.cur.execute(
                "SELECT row_id, data FROM pos_data WHERE data_key=? ORDER BY row_id",
                (data_key,)
            ).fetchall()
            result = []
            for row_id, data_str in rows:
                try:
                    rec = json.loads(data_str)
                    rec["_id"] = row_id
                    result.append(rec)
                except Exception:
                    pass
            return result
        except Exception as e:
            print(f"load_all error: {e}")
            return []

    def add(self, data_key, row_id, record):
        if not self.cur:
            return
        try:
            data_str = json.dumps({k: v for k, v in record.items() if k != "_id"})
            self.cur.execute(
                "INSERT INTO pos_data (data_key, row_id, data) VALUES (?, ?, ?)",
                (data_key, row_id, data_str)
            )
            self.con.commit()
        except Exception as e:
            print(f"add error: {e}")

    def update(self, data_key, row_id, record):
        """Delete then re-add (update pattern)."""
        if not self.cur:
            return
        try:
            self.cur.execute(
                "DELETE FROM pos_data WHERE data_key=? AND row_id=?",
                (data_key, row_id)
            )
            self.con.commit()
        except Exception as e:
            print(f"update-delete error: {e}")
        self.add(data_key, row_id, record)

    def delete(self, data_key, row_id):
        if not self.cur:
            return
        try:
            self.cur.execute(
                "DELETE FROM pos_data WHERE data_key=? AND row_id=?",
                (data_key, row_id)
            )
            self.con.commit()
        except Exception as e:
            print(f"delete error: {e}")

    def get_max_id(self, data_key):
        if not self.cur:
            return 0
        try:
            row = self.cur.execute(
                "SELECT MAX(row_id) FROM pos_data WHERE data_key=?",
                (data_key,)
            ).fetchone()
            return row[0] or 0
        except Exception:
            return 0


# ============================================================
# HELPERS
# ============================================================

def to_number(value):
    try:
        return float(str(value).replace(",", "").replace(" ", "") or 0)
    except Exception:
        return 0.0


def get_recent_months(count=24):
    months = []
    today = datetime.now()
    y, m = today.year, today.month
    for _ in range(count):
        months.append(f"{y:04d}-{m:02d}")
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return months


# ============================================================
# BASE TABLE PAGE
# ============================================================

class BaseTablePage(ctk.CTkFrame):

    def __init__(
        self,
        parent,
        controller,
        title,
        subtitle,
        columns,
        default_filename,
        input_fields=None,
        data_key=None,
        computed_columns=None,
        total_column=None,
        filter_options=None
    ):

        super().__init__(parent, fg_color="#f3f5f9")

        self.controller = controller
        self.title_text = title
        self.subtitle_text = subtitle
        self.columns = columns
        self.default_filename = default_filename
        self.data_key = data_key
        self.computed_columns = computed_columns or {}
        self.total_column = total_column
        self.filter_options = filter_options or []

        self.all_records = []
        self._next_id = 1
        self.field_widgets = {}
        self.filter_widgets = {}

        if input_fields is None:
            input_fields = [
                {"name": c, "type": "entry"}
                for c in columns
                if c != "Date" and c not in self.computed_columns
            ]

        self.input_fields = input_fields

        self._build_ui()
        self._load_from_shared_store()

    def _load_from_shared_store(self):

        if not self.data_key:
            return

        store = getattr(self.controller, "data_store", {})
        records = store.get(self.data_key, [])

        if not records and hasattr(self.controller, "persistence"):
            records = self.controller.persistence.load_all(self.data_key)
            store[self.data_key] = records

        self.all_records = list(records)

        max_id = 0
        for rec in self.all_records:
            try:
                rid = int(rec.get("_id", 0))
                if rid > max_id:
                    max_id = rid
            except Exception:
                pass

        self._next_id = max_id + 1
        self.refresh_table()

    def _build_ui(self):

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_scroll = ctk.CTkScrollableFrame(self, fg_color="#f3f5f9")
        self.main_scroll.grid(row=0, column=0, sticky="nsew")
        self.main_scroll.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 10))

        header_left = ctk.CTkFrame(header, fg_color="transparent")
        header_left.pack(side="left")

        ctk.CTkLabel(
            header_left, text=self.title_text,
            font=ctk.CTkFont(family=MYANMAR_FONT, size=28, weight="bold"),
            text_color="#111827"
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_left, text=self.subtitle_text,
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
            text_color="#6b7280"
        ).pack(anchor="w")

        ctk.CTkLabel(
            header, text=datetime.now().strftime("%d %B %Y"),
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            text_color="#374151"
        ).pack(side="right", pady=10)

        input_card = ctk.CTkFrame(self.main_scroll, fg_color="#ffffff", corner_radius=14)
        input_card.pack(fill="x", padx=15, pady=(0, 10))

        title_row = ctk.CTkFrame(input_card, fg_color="transparent")
        title_row.pack(fill="x", padx=20, pady=(15, 8))

        ctk.CTkLabel(
            title_row, text="📝 Import Data / ဒေတာထည့်ရန်",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=16, weight="bold"),
            text_color="#111827"
        ).pack(side="left")

        self.month_total_label = ctk.CTkLabel(
            title_row, text="📊 Month Total: 0",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            text_color="#2563eb", fg_color="#dbeafe",
            corner_radius=8, padx=14, pady=6
        )
        self.month_total_label.pack(side="right", padx=(10, 0))

        date_frame = ctk.CTkFrame(title_row, fg_color="transparent")
        date_frame.pack(side="right", padx=(10, 0))

        ctk.CTkLabel(
            date_frame, text="📅 Date:",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            text_color="#374151"
        ).pack(side="left", padx=(0, 6))

        self.date_widget = DateEntry(
            date_frame, date_pattern="yyyy-mm-dd",
            font=(MYANMAR_FONT, 12), background="#2563eb",
            foreground="white", borderwidth=0,
            justify="center", width=12
        )
        self.date_widget.set_date(date.today())
        self.date_widget.pack(side="left")

        fields_container = ctk.CTkFrame(input_card, fg_color="transparent")
        fields_container.pack(fill="x", padx=20, pady=(0, 10))

        MAX_COLS = 4

        for i, field in enumerate(self.input_fields):

            row = i // MAX_COLS
            col = i % MAX_COLS

            col_frame = ctk.CTkFrame(fields_container, fg_color="transparent")
            col_frame.grid(row=row, column=col, padx=4, pady=4, sticky="ew")

            ctk.CTkLabel(
                col_frame, text=field["name"],
                font=ctk.CTkFont(family=MYANMAR_FONT, size=11, weight="bold"),
                text_color="#374151"
            ).pack(anchor="w")

            widget = self._create_field_widget(col_frame, field)

            if widget is not None:
                widget.pack(fill="x", pady=(2, 0))
                self.field_widgets[field["name"]] = (field, widget)
                try:
                    widget.bind("<Return>", self._on_enter_key)
                    widget.bind("<KP_Enter>", self._on_enter_key)
                except Exception:
                    pass

        for c in range(MAX_COLS):
            fields_container.grid_columnconfigure(c, weight=1)

        btn_frame = ctk.CTkFrame(input_card, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(5, 15))

        ctk.CTkButton(
            btn_frame, text="➕ Add", width=110, height=38,
            fg_color="#2563eb", hover_color="#1d4ed8",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=self.add_row
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame, text="✖ Clear", width=110, height=38,
            fg_color="#ef4444", hover_color="#dc2626",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=self.clear_entries
        ).pack(side="left")

        table_card = ctk.CTkFrame(self.main_scroll, fg_color="#ffffff", corner_radius=14)
        table_card.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        toolbar = ctk.CTkFrame(table_card, fg_color="transparent")
        toolbar.pack(fill="x", padx=20, pady=(15, 5))

        ctk.CTkLabel(
            toolbar, text="📊 Data Table",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=14, weight="bold"),
            text_color="#111827"
        ).pack(side="left")

        ctk.CTkButton(
            toolbar, text="📥 Export to Excel", width=150, height=35,
            fg_color="#16a34a", hover_color="#15803d",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=lambda: export_to_excel(self.tree, self.default_filename)
        ).pack(side="right")

        ctk.CTkButton(
            toolbar, text="🗑 Delete Row", width=120, height=35,
            fg_color="#dc2626", hover_color="#b91c1c",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=self.delete_row
        ).pack(side="right", padx=(0, 8))

        for fopt in reversed(self.filter_options):
            self._create_filter_widget(toolbar, fopt)

        month_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        month_frame.pack(side="right", padx=4)

        ctk.CTkLabel(month_frame, text="📅",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=13)).pack(side="left", padx=(0, 3))

        recent_months = get_recent_months(24)
        current_month = datetime.now().strftime("%Y-%m")

        self.month_filter = ctk.CTkComboBox(
            month_frame,
            values=["All"] + recent_months,
            width=110, height=35, corner_radius=8,
            border_width=1, border_color="#cbd5e1",
            fg_color="#f8fafc", text_color="#111827",
            button_color="#7c3aed", button_hover_color="#6d28d9",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=12),
            dropdown_font=ctk.CTkFont(family=MYANMAR_FONT, size=12),
            state="readonly",
            command=lambda v: self.refresh_table()
        )
        self.month_filter.set(current_month)
        self.month_filter.pack(side="left")

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Treeview", background="#ffffff", foreground="#111827",
            rowheight=35, fieldbackground="#ffffff", font=(MYANMAR_FONT, 11)
        )
        style.map("Treeview", background=[("selected", "#dbeafe")])
        style.configure(
            "Treeview.Heading", font=(MYANMAR_FONT, 11, "bold"),
            background="#f3f4f6", foreground="#111827"
        )

        self.tree_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        self.tree_frame.pack(fill="both", expand=True, padx=20, pady=(5, 20))

        self.tree = ttk.Treeview(
            self.tree_frame, columns=self.columns,
            show="headings", height=12
        )

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)

    def _on_enter_key(self, event=None):
        try:
            self.add_row()
        except Exception as e:
            print("Enter key error:", e)
        return "break"

    def _create_field_widget(self, parent, field):

        ftype = field.get("type", "entry")

        if ftype == "entry":
            return ctk.CTkEntry(
                parent, height=38, corner_radius=8,
                border_width=1, border_color="#cbd5e1",
                fg_color="#f8fafc", text_color="#111827",
                font=ctk.CTkFont(family=MYANMAR_FONT, size=13)
            )

        elif ftype == "dropdown":
            w = ctk.CTkComboBox(
                parent,
                values=field.get("values", []),
                height=38, corner_radius=8,
                border_width=1, border_color="#cbd5e1",
                fg_color="#f8fafc", text_color="#111827",
                dropdown_text_color="#111827",
                button_color="#2563eb", button_hover_color="#1d4ed8",
                font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
                dropdown_font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
                state="readonly"
            )
            vals = field.get("values", [])
            if vals:
                w.set(vals[0])
            return w

        return None

    def _create_filter_widget(self, parent, fopt):

        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="right", padx=4)

        ctk.CTkLabel(
            frame, text=fopt.get("label", fopt["name"]) + ":",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=11, weight="bold"),
            text_color="#374151"
        ).pack(side="left", padx=(0, 3))

        w = ctk.CTkComboBox(
            frame,
            values=["All"] + fopt["values"],
            width=110, height=35, corner_radius=8,
            border_width=1, border_color="#cbd5e1",
            fg_color="#f8fafc", text_color="#111827",
            button_color="#f59e0b", button_hover_color="#d97706",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=12),
            dropdown_font=ctk.CTkFont(family=MYANMAR_FONT, size=12),
            state="readonly",
            command=lambda v: self.refresh_table()
        )
        w.set("All")
        w.pack(side="left")
        self.filter_widgets[fopt["column"]] = w

    def _get_field_value(self, field, widget):
        return widget.get().strip()

    def _clear_field(self, field, widget):
        ftype = field.get("type", "entry")
        if ftype == "entry":
            widget.delete(0, "end")
        elif ftype == "dropdown":
            vals = field.get("values", [])
            widget.set(vals[0] if vals else "")

    def add_row(self):

        field_values = {}

        for field in self.input_fields:
            if field["name"] in self.field_widgets:
                _, widget = self.field_widgets[field["name"]]
                field_values[field["name"]] = self._get_field_value(field, widget)
            else:
                field_values[field["name"]] = ""

        try:
            field_values["Date"] = self.date_widget.get_date().strftime("%Y-%m-%d")
        except Exception:
            field_values["Date"] = datetime.now().strftime("%Y-%m-%d")

        for col_name, config in self.computed_columns.items():
            source_cols = config.get("source", [])
            total = 0.0
            for src in source_cols:
                total += to_number(field_values.get(src, 0))
            field_values[col_name] = f"{total:,.0f}"

        values = {col: field_values.get(col, "") for col in self.columns}

        if not any(str(v).strip() for v in values.values()):
            messagebox.showwarning("Empty Input", "ကျေးဇူးပြု၍ အနည်းဆုံး တစ်ခုထည့်ပါ။")
            return

        record = dict(values)
        record["_id"] = self._next_id
        self._next_id += 1

        self.all_records.append(record)

        if self.data_key and hasattr(self.controller, "data_store"):
            self.controller.data_store.setdefault(self.data_key, []).append(record)

        if self.data_key and hasattr(self.controller, "persistence"):
            self.controller.persistence.add(self.data_key, record["_id"], record)

        self.refresh_table()
        self.clear_entries()

    def clear_entries(self):
        for field in self.input_fields:
            if field["name"] in self.field_widgets:
                _, widget = self.field_widgets[field["name"]]
                self._clear_field(field, widget)

    def delete_row(self):

        selected = self.tree.selection()

        if not selected:
            messagebox.showwarning("No Selection",
                                   "ကျေးဇူးပြု၍ ဖျက်လိုသော Row ကို ရွေးပါ။")
            return

        ids_to_remove = set(int(iid) for iid in selected)
        removed = [r for r in self.all_records if r["_id"] in ids_to_remove]

        self.all_records = [r for r in self.all_records if r["_id"] not in ids_to_remove]

        if self.data_key and hasattr(self.controller, "data_store"):
            store = self.controller.data_store.setdefault(self.data_key, [])
            for rec in removed:
                try:
                    store.remove(rec)
                except ValueError:
                    pass

        if self.data_key and hasattr(self.controller, "persistence"):
            for rec in removed:
                self.controller.persistence.delete(self.data_key, rec["_id"])

        self.refresh_table()

    def _record_matches_filters(self, record):

        month_val = self.month_filter.get()

        if month_val != "All":
            rec_date = str(record.get("Date", ""))
            if not rec_date.startswith(month_val):
                return False

        for col, widget in self.filter_widgets.items():
            val = widget.get()
            if val != "All":
                if str(record.get(col, "")) != val:
                    return False

        return True

    def refresh_table(self):

        for item in self.tree.get_children():
            self.tree.delete(item)

        visible_records = [r for r in self.all_records if self._record_matches_filters(r)]

        for rec in visible_records:
            values = [rec.get(col, "") for col in self.columns]
            self.tree.insert("", "end", iid=str(rec["_id"]), values=values)

        if self.total_column:
            total = 0.0
            for rec in visible_records:
                total += to_number(rec.get(self.total_column, "0"))
            self.month_total_label.configure(text=f"📊 Month Total: {total:,.0f}")


# ============================================================
# INVENTORY PAGE
# ============================================================

class InventoryPage(BaseTablePage):

    def __init__(self, parent, controller):

        columns = (
            "Date",
            "ပုလင်းခွံ/ ဘူးခွံ",
            "ဂျပ်ဖာ",
            "တံဆိပ်",
            "အရက်ကုန်ကြမ်း",
            "ရေသန့်",
            "အရောင်",
            "အရသာ",
            "others",
            "Total"
        )

        input_fields = [
            {"name": "ပုလင်းခွံ/ ဘူးခွံ", "type": "entry"},
            {"name": "ဂျပ်ဖာ", "type": "entry"},
            {"name": "တံဆိပ်", "type": "entry"},
            {"name": "အရက်ကုန်ကြမ်း", "type": "entry"},
            {"name": "ရေသန့်", "type": "entry"},
            {"name": "အရောင်", "type": "entry"},
            {"name": "အရသာ", "type": "entry"},
            {"name": "others", "type": "entry"},
        ]

        computed_columns = {
            "Total": {
                "source": [
                    "ပုလင်းခွံ/ ဘူးခွံ",
                    "ဂျပ်ဖာ",
                    "တံဆိပ်",
                    "အရက်ကုန်ကြမ်း",
                    "ရေသန့်",
                    "အရောင်",
                    "အရသာ",
                    "others",
                ]
            }
        }

        super().__init__(
            parent, controller,
            "ကုန်ပစ္စည်း",
            "Stock Management System",
            columns,
            "Inventory_Report.xlsx",
            input_fields=input_fields,
            data_key="inventory",
            computed_columns=computed_columns,
            total_column="Total"
        )


# ============================================================
# EMPLOYEE PAGE
# ============================================================

class EmployeePage(BaseTablePage):

    def __init__(self, parent, controller):

        columns = (
            "Date",
            "အလုပ်သမားနာမည်",
            "အမျိုးအစား",
            "Salary",
            "ရှင်းပြီး/မပြီး",
            "ခွင့်",
            "Remark"
        )

        input_fields = [
            {"name": "အလုပ်သမားနာမည်", "type": "entry"},
            {"name": "အမျိုးအစား", "type": "dropdown", "values": ["နေ့စား", "လခစား"]},
            {"name": "Salary", "type": "entry"},
            {"name": "ရှင်းပြီး/မပြီး", "type": "dropdown", "values": ["ရှင်းပြီး", "မပြီး"]},
            {"name": "ခွင့်", "type": "dropdown", "values": ["ခွင့်မယူ", "ခွင့်ယူ"]},
            {"name": "Remark", "type": "entry"},
        ]

        filter_options = [
            {"name": "အမျိုးအစား", "label": "အမျိုးအစား",
             "column": "အမျိုးအစား", "values": ["နေ့စား", "လခစား"]},
            {"name": "ရှင်းပြီး", "label": "ရှင်းပြီး",
             "column": "ရှင်းပြီး/မပြီး", "values": ["ရှင်းပြီး", "မပြီး"]},
            {"name": "ခွင့်", "label": "ခွင့်",
             "column": "ခွင့်", "values": ["ခွင့်ယူ", "ခွင့်မယူ"]},
        ]

        super().__init__(
            parent, controller,
            "အလုပ်သမားစာရင်း",
            "Employee Management",
            columns,
            "Employee_List.xlsx",
            input_fields=input_fields,
            data_key="employee",
            total_column="Salary",
            filter_options=filter_options
        )


# ============================================================
# EXPENSE PAGE
# ============================================================

class ExpensePage(BaseTablePage):

    def __init__(self, parent, controller):

        columns = ("Date", "Expense Type", "Amount", "Remark")

        input_fields = [
            {"name": "Expense Type", "type": "dropdown",
             "values": ["ကားခ", "ဆီဖိုး", "ပြင်ဆင်စရိတ်", "မီတာခ", "ဂိတ်ကြေး", "Others"]},
            {"name": "Amount", "type": "entry"},
            {"name": "Remark", "type": "entry"},
        ]

        super().__init__(
            parent, controller,
            "အထွေထွေကုန်ကျစရိတ်",
            "General Expenses",
            columns,
            "Expense_Report.xlsx",
            input_fields=input_fields,
            data_key="expense",
            total_column="Amount"
        )


# ============================================================
# ITEM LIST PAGE
# ============================================================

class ItemListPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(parent, fg_color="#f3f5f9")
        self.controller = controller

        try:
            self.dbWarehouse = warehouse("real/Warehousedatabase.db")
        except Exception as e:
            print("Warehouse DB error:", e)
            self.dbWarehouse = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(self, fg_color="#f3f5f9")
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(scroll, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 10))

        header_left = ctk.CTkFrame(header, fg_color="transparent")
        header_left.pack(side="left")

        ctk.CTkLabel(
            header_left, text="ပစ္စည်းစာရင်း",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=28, weight="bold"),
            text_color="#111827"
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_left,
            text="Product List — Bill page မှာ ရောင်းတိုင်း Stock auto လျော့ပါတယ်",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
            text_color="#6b7280"
        ).pack(anchor="w")

        ctk.CTkLabel(
            header, text=datetime.now().strftime("%d %B %Y"),
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            text_color="#374151"
        ).pack(side="right", pady=10)

        input_card = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=14)
        input_card.pack(fill="x", padx=15, pady=(0, 10))

        title_row = ctk.CTkFrame(input_card, fg_color="transparent")
        title_row.pack(fill="x", padx=20, pady=(15, 8))

        ctk.CTkLabel(
            title_row, text="📝 Import Data / ဒေတာထည့်ရန်",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=16, weight="bold"),
            text_color="#111827"
        ).pack(side="left")

        self.month_total_label = ctk.CTkLabel(
            title_row, text="📊 Month Total: 0",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            text_color="#2563eb", fg_color="#dbeafe",
            corner_radius=8, padx=14, pady=6
        )
        self.month_total_label.pack(side="right", padx=(10, 0))

        date_frame = ctk.CTkFrame(title_row, fg_color="transparent")
        date_frame.pack(side="right", padx=(10, 0))

        ctk.CTkLabel(
            date_frame, text="📅 Date:",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            text_color="#374151"
        ).pack(side="left", padx=(0, 6))

        self.date_widget = DateEntry(
            date_frame, date_pattern="yyyy-mm-dd",
            font=(MYANMAR_FONT, 12), background="#2563eb",
            foreground="white", borderwidth=0,
            justify="center", width=12
        )
        self.date_widget.set_date(date.today())
        self.date_widget.pack(side="left")

        fields_container = ctk.CTkFrame(input_card, fg_color="transparent")
        fields_container.pack(fill="x", padx=20, pady=(0, 10))

        self.input_fields = [
            {"name": "Item Code", "type": "entry"},
            {"name": "Item Name", "type": "entry"},
            {"name": "Category", "type": "entry"},
            {"name": "Price", "type": "entry"},
            {"name": "Stock", "type": "entry"},
        ]

        self.field_widgets = {}
        MAX_COLS = 5

        for i, field in enumerate(self.input_fields):

            col_frame = ctk.CTkFrame(fields_container, fg_color="transparent")
            col_frame.grid(row=0, column=i, padx=4, pady=4, sticky="ew")

            ctk.CTkLabel(
                col_frame, text=field["name"],
                font=ctk.CTkFont(family=MYANMAR_FONT, size=11, weight="bold"),
                text_color="#374151"
            ).pack(anchor="w")

            entry = ctk.CTkEntry(
                col_frame, height=38, corner_radius=8,
                border_width=1, border_color="#cbd5e1",
                fg_color="#f8fafc", text_color="#111827",
                font=ctk.CTkFont(family=MYANMAR_FONT, size=13)
            )
            entry.pack(fill="x", pady=(2, 0))
            entry.bind("<Return>", self._on_enter_key)
            entry.bind("<KP_Enter>", self._on_enter_key)
            self.field_widgets[field["name"]] = entry

        for c in range(MAX_COLS):
            fields_container.grid_columnconfigure(c, weight=1)

        btn_frame = ctk.CTkFrame(input_card, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(5, 15))

        ctk.CTkButton(
            btn_frame, text="➕ Add", width=110, height=38,
            fg_color="#2563eb", hover_color="#1d4ed8",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=self.add_row
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame, text="✖ Clear", width=110, height=38,
            fg_color="#ef4444", hover_color="#dc2626",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=self.clear_entries
        ).pack(side="left")

        table_card = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=14)
        table_card.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        toolbar = ctk.CTkFrame(table_card, fg_color="transparent")
        toolbar.pack(fill="x", padx=20, pady=(15, 5))

        ctk.CTkLabel(
            toolbar, text="📊 Data Table",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=14, weight="bold"),
            text_color="#111827"
        ).pack(side="left")

        self.columns = ("Date", "Item Code", "Item Name", "Category", "Price", "Stock")

        ctk.CTkButton(
            toolbar, text="📥 Export to Excel", width=150, height=35,
            fg_color="#16a34a", hover_color="#15803d",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=lambda: export_to_excel(self.tree, "Item_List.xlsx")
        ).pack(side="right")

        ctk.CTkButton(
            toolbar, text="🗑 Delete Row", width=120, height=35,
            fg_color="#dc2626", hover_color="#b91c1c",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=self.delete_row
        ).pack(side="right", padx=(0, 8))

        ctk.CTkButton(
            toolbar, text="🔄 Refresh", width=110, height=35,
            fg_color="#7c3aed", hover_color="#6d28d9",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=self.refresh_table
        ).pack(side="right", padx=(0, 8))

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Treeview", background="#ffffff", foreground="#111827",
                        rowheight=35, fieldbackground="#ffffff", font=(MYANMAR_FONT, 11))
        style.map("Treeview", background=[("selected", "#dbeafe")])
        style.configure("Treeview.Heading", font=(MYANMAR_FONT, 11, "bold"),
                        background="#f3f4f6", foreground="#111827")

        self.tree_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        self.tree_frame.pack(fill="both", expand=True, padx=20, pady=(5, 20))

        self.tree = ttk.Treeview(self.tree_frame, columns=self.columns,
                                 show="headings", height=12)

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="center")

        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)

        self.all_records = []
        self._next_id = 1

        self._load_from_shared_store()

    def _on_enter_key(self, event=None):
        try:
            self.add_row()
        except Exception as e:
            print("Enter error:", e)
        return "break"

    def _load_from_shared_store(self):

        store = getattr(self.controller, "data_store", {})
        records = store.get("itemlist", [])

        if not records and hasattr(self.controller, "persistence"):
            records = self.controller.persistence.load_all("itemlist")
            store["itemlist"] = records

        self.all_records = list(records)

        max_id = 0
        for rec in self.all_records:
            try:
                rid = int(rec.get("_id", 0))
                if rid > max_id:
                    max_id = rid
            except Exception:
                pass

        self._next_id = max_id + 1
        self.refresh_table()

    def add_row(self):

        values = {}

        for field in self.input_fields:
            widget = self.field_widgets.get(field["name"])
            values[field["name"]] = widget.get().strip() if widget else ""

        try:
            values["Date"] = self.date_widget.get_date().strftime("%Y-%m-%d")
        except Exception:
            values["Date"] = datetime.now().strftime("%Y-%m-%d")

        item_code = values.get("Item Code", "")
        item_name = values.get("Item Name", "")

        if item_code == "":
            messagebox.showwarning("Missing", "Item Code ထည့်ပါ။")
            return

        if item_name == "":
            messagebox.showwarning("Missing", "Item Name ထည့်ပါ။")
            return

        category = values.get("Category", "")
        price_str = values.get("Price", "0")
        stock_str = values.get("Stock", "0")

        try:
            price = int(price_str) if price_str else 0
        except Exception:
            price = 0

        try:
            stock = int(stock_str) if stock_str else 0
        except Exception:
            stock = 0

        try:
            if self.dbWarehouse:
                existing = self.dbWarehouse.Search(item_code)
                if existing:
                    try:
                        current_qty = int(existing[0][4])
                        new_qty = current_qty + stock
                        self.dbWarehouse.updateQuantity(item_code, new_qty)
                    except Exception as e:
                        print("Update quantity error:", e)
                else:
                    self.dbWarehouse.insert(
                        item_code,
                        item_name,
                        category if category else "General",
                        stock,
                        price,
                        price,
                        values["Date"],
                        values["Date"]
                    )
        except Exception as e:
            messagebox.showerror(
                "Warehouse DB Error",
                f"Warehouse DB ထဲ ထည့်လို့မရပါ:\n{e}"
            )
            return

        record = {col: values.get(col, "") for col in self.columns}
        record["_id"] = self._next_id
        self._next_id += 1

        self.all_records.append(record)

        if hasattr(self.controller, "data_store"):
            self.controller.data_store.setdefault("itemlist", []).append(record)

        if hasattr(self.controller, "persistence"):
            self.controller.persistence.add("itemlist", record["_id"], record)

        self.refresh_table()
        self.clear_entries()

        messagebox.showinfo(
            "Saved",
            f"'{item_code}' ကို Warehouse DB ထဲ သိမ်းပြီးပါပြီ။\n"
            f"အခု Bill Page မှာ '{item_code}' ရိုက်ထည့်လို့ ရပါပြီ။"
        )

    def clear_entries(self):
        for field in self.input_fields:
            widget = self.field_widgets.get(field["name"])
            if widget:
                widget.delete(0, "end")

    def delete_row(self):

        selected = self.tree.selection()

        if not selected:
            messagebox.showwarning("No Selection",
                                   "ကျေးဇူးပြု၍ ဖျက်လိုသော Row ကို ရွေးပါ။")
            return

        ids_to_remove = set(int(iid) for iid in selected)
        removed = [r for r in self.all_records if r["_id"] in ids_to_remove]

        self.all_records = [r for r in self.all_records if r["_id"] not in ids_to_remove]

        if hasattr(self.controller, "data_store"):
            store = self.controller.data_store.setdefault("itemlist", [])
            for rec in removed:
                try:
                    store.remove(rec)
                except ValueError:
                    pass

        if hasattr(self.controller, "persistence"):
            for rec in removed:
                self.controller.persistence.delete("itemlist", rec["_id"])

        self.refresh_table()

    def refresh_table(self):

        for item in self.tree.get_children():
            self.tree.delete(item)

        for rec in self.all_records:
            values = [rec.get(col, "") for col in self.columns]
            self.tree.insert("", "end", iid=str(rec["_id"]), values=values)

        total = 0.0
        for rec in self.all_records:
            total += to_number(rec.get("Price", "0"))

        self.month_total_label.configure(text=f"📊 Month Total: {total:,.0f}")


# ============================================================
# BILL CONTENT
# ============================================================

class BillContent(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(parent, fg_color="#d6d9e5")
        self.controller = controller

        try:
            self.dbWarehouse = warehouse("real/Warehousedatabase.db")
            self.dbTotal = DatabaseTotal("real/Totaldatabase.db")
            self.dbSales = DatabaseSales("real/SalesReportdatabase.db")
        except Exception as e:
            print("DB init error (BillContent):", e)
            self.dbWarehouse = self.dbTotal = self.dbSales = None

        self.bill_number = str(random.randint(100000, 999999))
        self.cart_items = []
        self.total_amount = 0
        self.user_pay = 0
        self.cashback = 0
        self.bill_saved = False
        self.seller_name = ""

        def only_numbers(value):
            return value == "" or value.isdigit()

        number_command = self.register(only_numbers)

        topbar = ctk.CTkFrame(self, height=60, fg_color="white", corner_radius=0)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        ctk.CTkLabel(
            topbar, text="အရောင်းစာရင်း / Bill",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=23, weight="bold"),
            text_color="#111827"
        ).pack(side="left", padx=20)

        ctk.CTkLabel(
            topbar, text="King Mark Group",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            text_color="#64748b"
        ).pack(side="right", padx=20)

        content = ctk.CTkFrame(self, fg_color="#d6d9e5", corner_radius=0)
        content.pack(fill="both", expand=True, padx=15, pady=15)
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)

        left_area = ctk.CTkFrame(content, fg_color="transparent")
        left_area.grid(row=0, column=0, sticky="nsew", padx=(0, 7))

        sales_card = ctk.CTkFrame(left_area, fg_color="white", corner_radius=14)
        sales_card.pack(fill="x", pady=(0, 20))

        sales_title_row = ctk.CTkFrame(sales_card, fg_color="transparent")
        sales_title_row.pack(fill="x", padx=18, pady=(10, 5))

        ctk.CTkLabel(
            sales_title_row, text="Today's Sales",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=18, weight="bold"),
            text_color="#111827"
        ).pack(side="left")

        self.sales_date_label = ctk.CTkLabel(
            sales_title_row, text=datetime.now().strftime("%Y-%m-%d"),
            font=ctk.CTkFont(family=MYANMAR_FONT, size=11, weight="bold"),
            text_color="#64748b"
        )
        self.sales_date_label.pack(side="right")

        sales_value_frame = ctk.CTkFrame(sales_card, fg_color="#c0d8f7", corner_radius=10)
        sales_value_frame.pack(fill="x", padx=18, pady=(0, 12))

        self.today_sales_label = ctk.CTkLabel(
            sales_value_frame, text="0",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=25, weight="bold"),
            text_color="#2563eb"
        )
        self.today_sales_label.pack(pady=8)

        def update_today_sales():
            today = datetime.now().strftime("%Y-%m-%d")
            try:
                result = self.dbSales.Total(today) if self.dbSales else None
                total_sales = result[0] if result and result[0] is not None else 0
                self.today_sales_label.configure(text=f"{int(total_sales):,}")
            except Exception:
                self.today_sales_label.configure(text="0")

        bill_info = ctk.CTkFrame(left_area, fg_color="white", corner_radius=14)
        bill_info.pack(fill="x", pady=(50, 20))

        ctk.CTkLabel(
            bill_info, text="Bill Information",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=18, weight="bold"),
            text_color="#111827"
        ).pack(anchor="w", padx=18, pady=(12, 8))

        bill_info_row = ctk.CTkFrame(bill_info, fg_color="transparent")
        bill_info_row.pack(fill="x", padx=18, pady=(0, 14))
        bill_info_row.grid_columnconfigure(0, weight=1)
        bill_info_row.grid_columnconfigure(1, weight=1)

        bill_number_frame = ctk.CTkFrame(bill_info_row, fg_color="#a4c6f2", corner_radius=9)
        bill_number_frame.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        ctk.CTkLabel(
            bill_number_frame, text="Bill Number",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=11, weight="bold"),
            text_color="#64748b"
        ).pack(anchor="w", padx=12, pady=(7, 0))

        self.bill_number_label = ctk.CTkLabel(
            bill_number_frame, text="#" + self.bill_number,
            font=ctk.CTkFont(family=MYANMAR_FONT, size=18, weight="bold"),
            text_color="#2563eb"
        )
        self.bill_number_label.pack(anchor="w", padx=12, pady=(0, 7))

        seller_frame = ctk.CTkFrame(bill_info_row, fg_color="#ade9ee", corner_radius=9)
        seller_frame.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        ctk.CTkLabel(
            seller_frame, text="Seller",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=11, weight="bold"),
            text_color="#64748b"
        ).pack(anchor="w", padx=12, pady=(7, 0))

        self.seller_entry = ctk.CTkEntry(
            seller_frame, height=30, border_width=0, fg_color="transparent",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=14, weight="bold"),
            text_color="#111827"
        )
        self.seller_entry.pack(fill="x", padx=7, pady=(0, 5))
        self.seller_entry.insert(0, self.seller_name)
        self.seller_entry.configure(state="readonly")

        add_card = ctk.CTkFrame(left_area, fg_color="white", corner_radius=14)
        add_card.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            add_card, text="Add Item",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=18, weight="bold"),
            text_color="#111827"
        ).pack(anchor="w", padx=18, pady=(12, 8))

        item_row = ctk.CTkFrame(add_card, fg_color="transparent")
        item_row.pack(fill="x", padx=18, pady=(0, 14))
        item_row.grid_columnconfigure(0, weight=2)
        item_row.grid_columnconfigure(1, weight=1)
        item_row.grid_columnconfigure(2, weight=1)

        item_frame = ctk.CTkFrame(item_row, fg_color="#b4cde6", corner_radius=8)
        item_frame.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        ctk.CTkLabel(
            item_frame, text="Item Code",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=14, weight="bold"),
            text_color="#5e7088"
        ).pack(anchor="w", padx=10, pady=(5, 0))

        self.item_entry = ctk.CTkEntry(
            item_frame, height=40, border_width=0, fg_color="transparent",
            placeholder_text="Scan / Enter",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=18)
        )
        self.item_entry.pack(fill="x", padx=5, pady=(0, 4))

        qty_frame = ctk.CTkFrame(item_row, fg_color="#b4cde6", corner_radius=8)
        qty_frame.grid(row=0, column=1, sticky="ew", padx=5)

        ctk.CTkLabel(
            qty_frame, text="Quantity",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=14, weight="bold"),
            text_color="#5e7088"
        ).pack(anchor="w", padx=10, pady=(5, 0))

        self.qty_entry = ctk.CTkEntry(
            qty_frame, height=40, border_width=0, fg_color="transparent",
            placeholder_text="Qty",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=18),
            validate="key",
            validatecommand=(number_command, "%P")
        )
        self.qty_entry.pack(fill="x", padx=5, pady=(0, 4))
        self.qty_entry.insert(0, "1")

        self.add_item_button = ctk.CTkButton(
            item_row, text="+  Add Item", height=50, corner_radius=8,
            fg_color="#2563eb", hover_color="#1d4ed8",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold")
        )
        self.add_item_button.grid(row=0, column=2, sticky="ew", padx=(5, 0))

        payment_card = ctk.CTkFrame(left_area, fg_color="white", corner_radius=14)
        payment_card.pack(fill="x")

        ctk.CTkLabel(
            payment_card, text="Payment",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=18, weight="bold"),
            text_color="#111827"
        ).pack(anchor="w", padx=18, pady=(12, 8))

        payment_row = ctk.CTkFrame(payment_card, fg_color="transparent")
        payment_row.pack(fill="x", padx=18, pady=(0, 14))
        payment_row.grid_columnconfigure(0, weight=1)
        payment_row.grid_columnconfigure(1, weight=1)
        payment_row.grid_columnconfigure(2, weight=1)

        total_frame = ctk.CTkFrame(payment_row, fg_color="#b4f5cc", corner_radius=8)
        total_frame.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        ctk.CTkLabel(
            total_frame, text="TOTAL",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=10, weight="bold"),
            text_color="#64748b"
        ).pack(anchor="w", padx=10, pady=(5, 0))

        self.total_label = ctk.CTkLabel(
            total_frame, text="0",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=17, weight="bold"),
            text_color="#111827"
        )
        self.total_label.pack(anchor="w", padx=10, pady=(0, 6))

        pay_frame = ctk.CTkFrame(payment_row, fg_color="#b6d0f2", corner_radius=8)
        pay_frame.grid(row=0, column=1, sticky="ew", padx=5)

        ctk.CTkLabel(
            pay_frame, text="USER PAY",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=10, weight="bold"),
            text_color="#64748b"
        ).pack(anchor="w", padx=10, pady=(5, 0))

        self.pay_entry = ctk.CTkEntry(
            pay_frame, height=30, border_width=0, fg_color="transparent",
            placeholder_text="Payment",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=15, weight="bold"),
            validate="key",
            validatecommand=(number_command, "%P")
        )
        self.pay_entry.pack(fill="x", padx=5, pady=(0, 4))

        cashback_frame = ctk.CTkFrame(payment_row, fg_color="#e6ebab", corner_radius=8)
        cashback_frame.grid(row=0, column=2, sticky="ew", padx=(5, 0))

        ctk.CTkLabel(
            cashback_frame, text="CASHBACK",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=10, weight="bold"),
            text_color="#64748b"
        ).pack(anchor="w", padx=10, pady=(5, 0))

        self.cashback_label = ctk.CTkLabel(
            cashback_frame, text="0",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=17, weight="bold"),
            text_color="#16a34a"
        )
        self.cashback_label.pack(anchor="w", padx=10, pady=(0, 6))

        right_area = ctk.CTkFrame(content, fg_color="transparent")
        right_area.grid(row=0, column=1, sticky="nsew", padx=(7, 0))

        receipt_card = ctk.CTkFrame(right_area, fg_color="#cbd5e1", corner_radius=14)
        receipt_card.pack(fill="both", expand=True)

        ctk.CTkLabel(
            receipt_card, text="Receipt Preview • A4 size",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=15, weight="bold"),
            text_color="#334155"
        ).pack(pady=(8, 5))

        receipt_frame = ctk.CTkFrame(receipt_card, fg_color="white", corner_radius=3)
        receipt_frame.pack(fill="both", expand=True, padx=15, pady=(0, 8))

        preview_font = ctk.CTkFont(family="Courier New", size=11)
        preview_bold_font = ctk.CTkFont(family="Courier New", size=11, weight="bold")

        ctk.CTkLabel(
            receipt_frame, text="King Mark Group",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=22, weight="bold"),
            text_color="black"
        ).pack(pady=(10, 0))

        ctk.CTkLabel(
            receipt_frame, text="Ph:09777275950",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=11),
            text_color="black"
        ).pack(pady=(0, 2))

        self.receipt_bill_label = ctk.CTkLabel(
            receipt_frame, text="Bill Number: " + self.bill_number,
            font=preview_bold_font, text_color="black"
        )
        self.receipt_bill_label.pack()

        self.receipt_date_label = ctk.CTkLabel(
            receipt_frame, text="Date: " + datetime.now().strftime("%Y-%m-%d %H:%M"),
            font=preview_font, text_color="black"
        )
        self.receipt_date_label.pack(pady=(0, 3))

        ctk.CTkLabel(receipt_frame, text="-" * 48,
                     font=ctk.CTkFont(family="Courier New", size=10),
                     text_color="black").pack()

        header_row = ctk.CTkFrame(receipt_frame, fg_color="white")
        header_row.pack(fill="x", padx=8)

        for text, w, anchor in [("No", 30, "center"), ("Name", 105, "w"),
                                ("Qty", 35, "e"), ("Price", 55, "e"), ("Amount", 65, "e")]:
            ctk.CTkLabel(header_row, text=text, width=w, anchor=anchor,
                         font=preview_bold_font, text_color="black").pack(side="left")

        self.receipt_items = ctk.CTkScrollableFrame(receipt_frame, fg_color="white",
                                                    corner_radius=0)
        self.receipt_items.pack(fill="both", expand=True, padx=4, pady=2)

        ctk.CTkLabel(receipt_frame, text="-" * 48,
                     font=ctk.CTkFont(family="Courier New", size=10),
                     text_color="black").pack()

        total_receipt_row = ctk.CTkFrame(receipt_frame, fg_color="white")
        total_receipt_row.pack(fill="x", padx=12)

        ctk.CTkLabel(total_receipt_row, text="TOTAL", font=preview_bold_font,
                     text_color="black").pack(side="left")

        self.receipt_total_label = ctk.CTkLabel(total_receipt_row, text="0",
                                                font=preview_bold_font, text_color="black")
        self.receipt_total_label.pack(side="right")

        cashback_receipt_row = ctk.CTkFrame(receipt_frame, fg_color="white")
        cashback_receipt_row.pack(fill="x", padx=12)

        ctk.CTkLabel(cashback_receipt_row, text="Cashback", font=preview_font,
                     text_color="black").pack(side="left")

        self.receipt_cashback_label = ctk.CTkLabel(cashback_receipt_row, text="0",
                                                   font=preview_font, text_color="black")
        self.receipt_cashback_label.pack(side="right")

        ctk.CTkLabel(receipt_frame, text="Thank You!",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=11, weight="bold"),
                     text_color="black").pack(pady=(3, 5))

        button_row = ctk.CTkFrame(right_area, fg_color="transparent")
        button_row.pack(fill="x", pady=(8, 0))
        button_row.grid_columnconfigure(0, weight=1)
        button_row.grid_columnconfigure(1, weight=1)

        self.clear_button = ctk.CTkButton(
            button_row, text="Clear Bill", height=44, corner_radius=9,
            fg_color="#dc2626", hover_color="#b91c1c",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"))
        self.clear_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.print_button = ctk.CTkButton(
            button_row, text="Print", height=44, corner_radius=9,
            fg_color="#16a34a", hover_color="#15803d",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"))
        self.print_button.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        # BEHAVIOR
        def update_payment(event=None):
            try:
                self.user_pay = int(self.pay_entry.get()) if self.pay_entry.get() else 0
                self.cashback = self.user_pay - self.total_amount
                self.cashback_label.configure(text=f"{self.cashback:,}")
                self.receipt_cashback_label.configure(text=f"{self.cashback:,}")
            except Exception:
                self.user_pay = 0
                self.cashback = 0
                self.cashback_label.configure(text="0")
                self.receipt_cashback_label.configure(text="0")

        self.pay_entry.bind("<KeyRelease>", update_payment)

        def remove_item(item):
            if item in self.cart_items:
                self.cart_items.remove(item)
            self.calculate_total()
            refresh_receipt()
            self.bill_saved = False
            self.item_entry.focus()

        def refresh_receipt():
            for widget in self.receipt_items.winfo_children():
                widget.destroy()

            for index, item in enumerate(self.cart_items, start=1):
                row = ctk.CTkFrame(self.receipt_items, fg_color="white")
                row.pack(fill="x", padx=2, pady=2)

                ctk.CTkLabel(row, text=str(index), width=30, font=preview_font,
                             text_color="black").pack(side="left")

                name = str(item["name"])
                if len(name) > 16:
                    name = name[:16]

                ctk.CTkLabel(row, text=name, width=105, anchor="w",
                             font=preview_font, text_color="black").pack(side="left")
                ctk.CTkLabel(row, text=str(item["quantity"]), width=35, anchor="e",
                             font=preview_font, text_color="black").pack(side="left")
                ctk.CTkLabel(row, text=str(item["price"]), width=55, anchor="e",
                             font=preview_font, text_color="black").pack(side="left")
                ctk.CTkLabel(row, text=str(item["amount"]), width=65, anchor="e",
                             font=preview_font, text_color="black").pack(side="left")

                ctk.CTkButton(
                    row, text="X", width=30, height=25, corner_radius=5,
                    fg_color="#ef4444", hover_color="#dc2626",
                    font=ctk.CTkFont(family=MYANMAR_FONT, size=10, weight="bold"),
                    command=lambda ci=item: remove_item(ci)
                ).pack(side="left", padx=(5, 0))

            self.total_label.configure(text=f"{self.total_amount:,}")
            self.receipt_total_label.configure(text=f"{self.total_amount:,}")
            update_payment()

        def calculate_total():
            self.total_amount = sum(i["amount"] for i in self.cart_items)

        self.calculate_total = calculate_total

        def add_item(event=None):

            ItemCode = self.item_entry.get().strip()
            QtyText = self.qty_entry.get().strip()

            if ItemCode == "":
                messagebox.showwarning("Warning", "Please enter Item Code.")
                self.item_entry.focus()
                return

            if QtyText == "":
                QtyText = "1"

            Qty = int(QtyText)

            if Qty <= 0:
                messagebox.showwarning("Warning", "Quantity must be greater than 0.")
                self.qty_entry.focus()
                return

            rows = self.dbWarehouse.Search(ItemCode) if self.dbWarehouse else None

            if not rows:
                messagebox.showerror(
                    "Item Not Found",
                    f"'{ItemCode}' ကို Warehouse DB ထဲမှာ ရှာမတွေ့ပါ။\n\n"
                    f"ကျေးဇူးပြု၍ 'ပစ္စည်းစာရင်း' မှာ အရင်ထည့်ပါ။"
                )
                self.item_entry.focus()
                return

            item = rows[0]
            DBQuantity = int(item[4])
            Name = item[2]
            Price = int(item[6])

            already_in_cart = 0
            for ci in self.cart_items:
                if ci["itemcode"] == ItemCode:
                    already_in_cart = ci["quantity"]
                    break

            if already_in_cart + Qty > DBQuantity:
                messagebox.showwarning("Not Enough Stock",
                                       f"Available = {DBQuantity} ခု")
                return

            found = False
            for ci in self.cart_items:
                if ci["itemcode"] == ItemCode:
                    ci["quantity"] += Qty
                    ci["amount"] = ci["quantity"] * Price
                    found = True
                    break

            if not found:
                self.cart_items.append({
                    "itemcode": ItemCode, "name": Name,
                    "quantity": Qty, "price": Price,
                    "amount": Qty * Price
                })

            calculate_total()
            refresh_receipt()

            self.item_entry.delete(0, "end")
            self.qty_entry.delete(0, "end")
            self.qty_entry.insert(0, "1")
            self.item_entry.focus()
            self.bill_saved = False

        self.add_item_button.configure(command=add_item)
        self.item_entry.bind("<Return>", add_item)
        self.item_entry.bind("<KP_Enter>", add_item)
        self.qty_entry.bind("<Return>", add_item)
        self.qty_entry.bind("<KP_Enter>", add_item)

        # ====================================================
        # SYNC ITEMLIST STOCK (decrease after sale)
        # ====================================================

        def sync_itemlist_stock(sold_items):
            """
            For each sold item, decrease the ItemList's Stock column
            in data_store["itemlist"] and persistence.
            """
            store = getattr(controller, "data_store", None)
            if not store:
                return

            itemlist = store.get("itemlist", [])
            if not itemlist:
                return

            for sold in sold_items:
                code = sold["itemcode"]
                qty = sold["quantity"]

                for il_rec in itemlist:
                    if str(il_rec.get("Item Code", "")) == code:
                        try:
                            current_stock = to_number(il_rec.get("Stock", "0"))
                            new_stock = max(current_stock - qty, 0)
                            il_rec["Stock"] = str(int(new_stock))

                            if hasattr(controller, "persistence"):
                                controller.persistence.update(
                                    "itemlist", il_rec["_id"], il_rec
                                )
                        except Exception as e:
                            print(f"sync_itemlist_stock error: {e}")
                        break

        def save_bill():

            if len(self.cart_items) == 0:
                messagebox.showwarning("Warning", "Please add items before saving.")
                return False

            if self.bill_saved:
                messagebox.showinfo("Already Saved", "This bill has already been saved.")
                return True

            seller = self.seller_name or getattr(controller, "current_username", "")

            if seller == "":
                messagebox.showwarning("Seller", "Seller username is empty.")
                return False

            if self.pay_entry.get() == "":
                messagebox.showwarning("Payment", "Please enter User Pay.")
                self.pay_entry.focus()
                return False

            try:
                user_pay = int(self.pay_entry.get())
            except Exception:
                messagebox.showwarning("Payment", "Please enter a valid amount.")
                self.pay_entry.focus()
                return False

            if user_pay < self.total_amount:
                messagebox.showwarning("Payment Error", "User Pay is less than Total.")
                self.pay_entry.focus()
                return False

            self.user_pay = user_pay
            self.cashback = user_pay - self.total_amount
            update_payment()

            bill_date = datetime.now().strftime("%Y-%m-%d")

            try:
                for item in self.cart_items:

                    try:
                        robust_db_insert(
                            self.dbTotal,
                            item["itemcode"], item["name"], str(item["amount"])
                        )
                    except Exception as e:
                        print(f"dbTotal.insert failed: {e}")

                    try:
                        robust_db_insert(
                            self.dbTotal,
                            self.bill_number, item["itemcode"], item["name"],
                            item["quantity"], item["price"],
                            str(item["amount"]), seller
                        )
                    except Exception as e:
                        print(f"dbTotal.insertBill failed: {e}")

                    try:
                        robust_db_insert(
                            self.dbSales,
                            bill_date, item["amount"], self.bill_number, seller,
                            item["quantity"], item["itemcode"], item["name"]
                        )
                    except Exception as e:
                        print(f"dbSales.insert failed: {e}")

                    try:
                        current = self.dbWarehouse.getQuntity(item["itemcode"])
                        if current:
                            new_quantity = max(int(current[0]) - item["quantity"], 0)
                            self.dbWarehouse.updateQuantity(item["itemcode"], new_quantity)
                    except Exception as e:
                        print(f"Warehouse update failed: {e}")

                # ============ SYNC ITEMLIST STOCK ============
                sync_itemlist_stock(self.cart_items)

                self.bill_saved = True
                update_today_sales()
                messagebox.showinfo("Saved",
                                    "Bill saved successfully.\n\nYou can now press Print.")
                return True
            except Exception as e:
                messagebox.showerror("Database Error", str(e))
                return False

        self.pay_entry.bind("<Return>", lambda event: save_bill())
        self.pay_entry.bind("<KP_Enter>", lambda event: save_bill())

        def clear_bill():
            if len(self.cart_items) == 0 and not self.bill_saved:
                return
            if not messagebox.askyesno("Clear Bill",
                                       "Are you sure you want to clear this bill?"):
                return

            self.cart_items.clear()
            self.total_amount = 0
            self.user_pay = 0
            self.cashback = 0
            self.bill_saved = False

            self.item_entry.delete(0, "end")
            self.qty_entry.delete(0, "end")
            self.qty_entry.insert(0, "1")
            self.pay_entry.delete(0, "end")

            self.total_label.configure(text="0")
            self.cashback_label.configure(text="0")
            self.receipt_total_label.configure(text="0")
            self.receipt_cashback_label.configure(text="0")

            refresh_receipt()
            self.item_entry.focus()

        self.clear_button.configure(command=clear_bill)

        def print_bill():
            if len(self.cart_items) == 0:
                messagebox.showwarning("Warning", "There are no items to print.")
                return

            if not self.bill_saved:
                if not messagebox.askyesno("Bill Not Saved",
                                           "This bill has not been saved.\n\nSave it before printing?"):
                    return
                if not save_bill():
                    return

            seller = self.seller_name or getattr(controller, "current_username", "")
            bill_datetime = datetime.now().strftime("%Y-%m-%d %H:%M")

            receipt_lines = [
                "King Mark Group",
                "Ph:09777275950",
                "Bill Number: " + self.bill_number,
                "Date: " + bill_datetime,
                "-" * 48,
                "No  Name              Qty Price  Amount",
                "-" * 48,
            ]

            for index, item in enumerate(self.cart_items, start=1):
                name = str(item["name"])[:17]
                receipt_lines.append(
                    f"{index:<3}{name:<18}{item['quantity']:>4}{item['price']:>6}{item['amount']:>9}"
                )

            receipt_lines += [
                "-" * 48,
                f"{'TOTAL':<30}{self.total_amount:>18}",
                f"{'PAY':<30}{self.user_pay:>18}",
                f"{'CASHBACK':<30}{self.cashback:>18}",
                "-" * 48,
                "Seller: " + seller,
                "Thank You!",
            ]

            receipt_text = "\n".join(receipt_lines)
            receipt_file = os.path.join(tempfile.gettempdir(),
                                        "ABK_Bill_" + self.bill_number + ".txt")

            try:
                with open(receipt_file, "w", encoding="utf-8") as f:
                    f.write(receipt_text)
            except Exception as e:
                messagebox.showerror("File Error", str(e))
                return

            try:
                os.startfile(receipt_file, "print")
            except Exception as e:
                messagebox.showerror("Printer Error",
                                     "Windows could not send the receipt to printer.\n\n" + str(e))
                return

            messagebox.showinfo("Print", "Bill sent to printer.")

            self.cart_items.clear()
            self.total_amount = 0
            self.user_pay = 0
            self.cashback = 0
            self.bill_saved = False

            self.item_entry.delete(0, "end")
            self.qty_entry.delete(0, "end")
            self.qty_entry.insert(0, "1")
            self.pay_entry.delete(0, "end")

            self.bill_number = str(random.randint(100000, 999999))
            self.bill_number_label.configure(text="#" + self.bill_number)
            self.receipt_bill_label.configure(text="Bill Number: " + self.bill_number)
            self.receipt_date_label.configure(
                text="Date: " + datetime.now().strftime("%Y-%m-%d %H:%M"))

            self.total_label.configure(text="0")
            self.cashback_label.configure(text="0")
            self.receipt_total_label.configure(text="0")
            self.receipt_cashback_label.configure(text="0")

            refresh_receipt()
            update_today_sales()
            self.item_entry.focus()

        self.print_button.configure(command=print_bill)

        refresh_receipt()
        update_today_sales()
        self.item_entry.focus()

    def refresh_seller(self):
        username = getattr(self.controller, "current_username", "")
        self.seller_name = username
        try:
            self.seller_entry.configure(state="normal")
            self.seller_entry.delete(0, "end")
            if username:
                self.seller_entry.insert(0, username)
            self.seller_entry.configure(state="readonly")
        except Exception:
            pass


# ============================================================
# SALES PAGE WRAPPER
# ============================================================

class SalesPage(ctk.CTkFrame):

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="#f3f5f9")
        self.controller = controller
        self.bill = BillContent(self, controller)
        self.bill.pack(fill="both", expand=True)

    def refresh_seller(self):
        self.bill.refresh_seller()


# ============================================================
# PROFIT PAGE
# ============================================================

class ProfitPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(parent, fg_color="#f3f5f9")
        self.controller = controller

        try:
            self.dbSales = DatabaseSales("real/SalesReportdatabase.db")
        except Exception:
            self.dbSales = None

        scroll = ctk.CTkScrollableFrame(self, fg_color="#f3f5f9")
        scroll.pack(fill="both", expand=True)

        header = ctk.CTkFrame(scroll, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 10))

        header_left = ctk.CTkFrame(header, fg_color="transparent")
        header_left.pack(side="left")

        ctk.CTkLabel(
            header_left, text="အမြတ်စာရင်း",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=28, weight="bold"),
            text_color="#111827"
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_left, text="Profit & Loss Statement (Auto Calculated)",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
            text_color="#6b7280"
        ).pack(anchor="w")

        right_hdr = ctk.CTkFrame(header, fg_color="transparent")
        right_hdr.pack(side="right", pady=10)

        ctk.CTkLabel(right_hdr, text="📅 Month:",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
                     text_color="#374151").pack(side="left", padx=(0, 6))

        self.month_filter = ctk.CTkComboBox(
            right_hdr, values=get_recent_months(24),
            width=120, height=36, corner_radius=8,
            border_width=1, border_color="#cbd5e1",
            fg_color="#f8fafc", text_color="#111827",
            button_color="#7c3aed", button_hover_color="#6d28d9",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
            dropdown_font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
            state="readonly",
            command=lambda v: self.refresh()
        )
        self.month_filter.set(datetime.now().strftime("%Y-%m"))
        self.month_filter.pack(side="left")

        ctk.CTkButton(right_hdr, text="🔄 Refresh", width=100, height=36,
                      fg_color="#2563eb", hover_color="#1d4ed8",
                      font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
                      command=self.refresh).pack(side="left", padx=(8, 0))

        formula_card = ctk.CTkFrame(scroll, fg_color="#ffffff", corner_radius=14)
        formula_card.pack(fill="x", padx=15, pady=(0, 10))

        ctk.CTkLabel(formula_card, text="🧮 Calculation Formula",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=16, weight="bold"),
                     text_color="#111827").pack(anchor="w", padx=20, pady=(15, 5))

        ctk.CTkLabel(
            formula_card,
            text=(
                "ကုန်ပစ္စည်း(လစဉ်) + အထွေထွေကုန်ကျစရိတ်(လစဉ်) "
                "- အလုပ်သမားစာရင်း(လစဉ်) - အရောင်းစာရင်း(လစဉ်) = အမြတ် / အရှုံး"
            ),
            font=ctk.CTkFont(family=MYANMAR_FONT, size=12),
            text_color="#374151", justify="left"
        ).pack(anchor="w", padx=20, pady=(0, 15))

        self.cards_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        self.cards_frame.pack(fill="x", padx=15, pady=(0, 10))

        for i in range(2):
            self.cards_frame.grid_columnconfigure(i, weight=1)

        self.card_inventory = self._make_card(self.cards_frame, 0, 0,
            "ကုန်ပစ္စည်း (Inventory)", "0", "📦", "#dbeafe")
        self.card_expense = self._make_card(self.cards_frame, 0, 1,
            "အထွေထွေကုန်ကျစရိတ် (Expense)", "0", "💸", "#fef3c7")
        self.card_employee = self._make_card(self.cards_frame, 1, 0,
            "အလုပ်သမားစာရင်း (Employee)", "0", "👷", "#fde68a")
        self.card_sales = self._make_card(self.cards_frame, 1, 1,
            "အရောင်းစာရင်း (Sales)", "0", "💰", "#dcfce7")
        self.card_profit = self._make_card(self.cards_frame, 2, 0,
            "အမြတ် / အရှုံး (Net Profit)", "0", "📈", "#d1fae5")
        self.card_month = self._make_card(self.cards_frame, 2, 1,
            "လအတွင်း", "-", "📅", "#e0e7ff")

        self.refresh()

    def _make_card(self, parent, row, col, title, value, icon, icon_bg):

        card = ctk.CTkFrame(parent, fg_color="#ffffff", corner_radius=14, height=130)
        card.grid(row=row, column=col, sticky="ew", padx=7, pady=7)
        card.grid_propagate(False)

        icon_frame = ctk.CTkFrame(card, width=55, height=55, corner_radius=12,
                                  fg_color=icon_bg)
        icon_frame.place(x=18, y=35)
        icon_frame.pack_propagate(False)
        ctk.CTkLabel(icon_frame, text=icon, font=ctk.CTkFont(size=22)).pack(expand=True)

        ctk.CTkLabel(card, text=title,
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
                     text_color="#6b7280").place(x=90, y=25)

        v = ctk.CTkLabel(card, text=value,
                         font=ctk.CTkFont(family=MYANMAR_FONT, size=22, weight="bold"),
                         text_color="#111827")
        v.place(x=90, y=55)
        return v

    def refresh(self):

        month = self.month_filter.get()
        store = getattr(self.controller, "data_store", {})

        if not store.get("inventory") and hasattr(self.controller, "persistence"):
            store["inventory"] = self.controller.persistence.load_all("inventory")
        if not store.get("expense") and hasattr(self.controller, "persistence"):
            store["expense"] = self.controller.persistence.load_all("expense")
        if not store.get("employee") and hasattr(self.controller, "persistence"):
            store["employee"] = self.controller.persistence.load_all("employee")

        inventory_total = 0.0
        for rec in store.get("inventory", []):
            if str(rec.get("Date", "")).startswith(month):
                inventory_total += to_number(rec.get("Total", "0"))

        expense_total = 0.0
        for rec in store.get("expense", []):
            if str(rec.get("Date", "")).startswith(month):
                expense_total += to_number(rec.get("Amount", "0"))

        employee_total = 0.0
        for rec in store.get("employee", []):
            if str(rec.get("Date", "")).startswith(month):
                employee_total += to_number(rec.get("Salary", "0"))

        sales_total = 0.0
        try:
            if self.dbSales:
                self.dbSales.cur.execute(
                    "SELECT SUM(sales) FROM SalesReport WHERE strftime('%Y-%m', date)=?",
                    (month,)
                )
                result = self.dbSales.cur.fetchone()
                if result and result[0] is not None:
                    sales_total = float(result[0])
        except Exception as e:
            print("Sales Total Error:", e)

        # Inventory + Expense - Employee - Sales = Profit/Loss
        profit = inventory_total + expense_total - employee_total - sales_total

        self.card_inventory.configure(text=f"{inventory_total:,.0f}")
        self.card_expense.configure(text=f"{expense_total:,.0f}")
        self.card_employee.configure(text=f"{employee_total:,.0f}")
        self.card_sales.configure(text=f"{sales_total:,.0f}")
        self.card_profit.configure(text=f"{profit:,.0f}")
        self.card_month.configure(text=month)

        if profit >= 0:
            self.card_profit.configure(text_color="#15803d")
        else:
            self.card_profit.configure(text_color="#dc2626")

    def tkraise(self, aboveThis=None):
        result = super().tkraise(aboveThis)
        if hasattr(self, "card_profit") and self.card_profit.winfo_exists():
            self.refresh()
        return result


# ============================================================
# USER PAGE
# ============================================================

class UserPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(parent, fg_color="#d2daea")
        self.controller = controller

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#111827")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.pack(fill="x", padx=15, pady=(25, 20))

        ctk.CTkLabel(logo_frame, text="KM",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=36, weight="bold"),
                     text_color="#ffffff").pack()

        ctk.CTkLabel(logo_frame, text="King Mark Group",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=12),
                     text_color="#9ca3af").pack(pady=(0, 5))

        self.user_label = ctk.CTkLabel(logo_frame, text="User",
                                       font=ctk.CTkFont(family=MYANMAR_FONT, size=11),
                                       text_color="#60a5fa")
        self.user_label.pack(pady=(0, 5))

        nav_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        nav_frame.pack(fill="both", expand=True, padx=12)

        self.create_sidebar_button(nav_frame, "အရောင်းစာရင်း",
                                   lambda: self.show_content("sales"))
        self.create_sidebar_button(nav_frame, "Logout",
                                   lambda: controller.show_frame(FirstPage))

        bottom_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bottom_frame.pack(side="bottom", fill="x", padx=12, pady=15)

        ctk.CTkLabel(bottom_frame, text="King Mark Group POS",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=10),
                     text_color="#6b7280").pack()

        self.content_area = ctk.CTkFrame(self, fg_color="#99acd1")
        self.content_area.grid(row=0, column=1, sticky="nsew")
        self.content_area.grid_columnconfigure(0, weight=1)
        self.content_area.grid_rowconfigure(0, weight=1)

        self.bill_content = None
        self.show_content("sales")

    def create_sidebar_button(self, parent, text, command):

        button = ctk.CTkButton(
            parent, text=text, height=45, corner_radius=8,
            fg_color="transparent", hover_color="#1f2937",
            text_color="#d1d5db", anchor="w",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=command
        )
        button.pack(fill="x", pady=4)
        return button

    def show_content(self, page_name):

        for widget in self.content_area.winfo_children():
            widget.destroy()

        if page_name == "sales":
            self.bill_content = BillContent(self.content_area, self.controller)
            self.bill_content.pack(fill="both", expand=True)
            self.bill_content.refresh_seller()

    def set_username(self, username):
        self.user_label.configure(text=f"User: {username}")

    def tkraise(self, aboveThis=None):
        result = super().tkraise(aboveThis)
        if self.bill_content:
            try:
                self.bill_content.refresh_seller()
            except Exception:
                pass
        return result


# ============================================================
# FIRST PAGE - LOGIN
# ============================================================

class FirstPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(parent, fg_color="#1e293b")
        self.controller = controller

        self.bg_frame = ctk.CTkFrame(self, fg_color="#f3f5f9")
        self.bg_frame.pack(fill="both", expand=True)

        self.paned = tk.PanedWindow(
            self.bg_frame, orient=tk.HORIZONTAL,
            sashwidth=0, bd=0, relief="flat", bg="#f3f5f9"
        )
        self.paned.pack(fill="both", expand=True, padx=40, pady=40)

        self.left_panel = ctk.CTkFrame(self.paned, fg_color="#FEE504", corner_radius=25)
        self.paned.add(self.left_panel, minsize=400)

        try:
            logo_image = Image.open("Logo.png")
            logo_image = logo_image.resize((180, 180))
            self.logo_title = ctk.CTkLabel(
                self.left_panel, text="",
                image=ctk.CTkImage(light_image=logo_image,
                                   dark_image=logo_image, size=(180, 180))
            )
            self.logo_title.pack(pady=(120, 10))
        except Exception:
            pass

        ctk.CTkFrame(self.left_panel, width=100, height=5,
                     fg_color="#fbbf24", corner_radius=5).pack(pady=10)

        ctk.CTkLabel(self.left_panel, text="POS APPLICATION",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=22, weight="bold"),
                     text_color="#060606").pack(pady=(20, 5))

        ctk.CTkLabel(self.left_panel, text="Stock Management System",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=15),
                     text_color="#ff0000").pack(pady=5)

        ctk.CTkLabel(self.left_panel, text="Secure • Fast • Simple",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=14),
                     text_color="#312b3f").pack(pady=60)

        self.right_panel = ctk.CTkFrame(self.paned, fg_color="#ffffff", corner_radius=25)
        self.paned.add(self.right_panel, minsize=650)

        ctk.CTkLabel(self.right_panel, text="Welcome Back",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=38, weight="bold"),
                     text_color="#111827").pack(pady=(90, 5))

        ctk.CTkLabel(self.right_panel, text="Please login to continue",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=16),
                     text_color="#64748b").pack(pady=(0, 35))

        self.form_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.form_frame.pack(padx=80, fill="x")

        ctk.CTkLabel(self.form_frame, text="Username", anchor="w",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=15, weight="bold"),
                     text_color="#334155").pack(fill="x", pady=(0, 8))

        text1 = tk.StringVar()
        text2 = tk.StringVar()

        self.T1 = ctk.CTkEntry(self.form_frame, height=50, corner_radius=10,
                               border_width=1, border_color="#cbd5e1",
                               fg_color="#f8fafc", text_color="#111827",
                               font=ctk.CTkFont(family=MYANMAR_FONT, size=16),
                               textvariable=text1,
                               placeholder_text="Enter username")
        self.T1.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(self.form_frame, text="Password", anchor="w",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=15, weight="bold"),
                     text_color="#334155").pack(fill="x", pady=(0, 8))

        self.T2 = ctk.CTkEntry(self.form_frame, height=50, corner_radius=10,
                               border_width=1, border_color="#cbd5e1",
                               fg_color="#f8fafc", text_color="#111827",
                               font=ctk.CTkFont(family=MYANMAR_FONT, size=16),
                               textvariable=text2,
                               placeholder_text="Enter password", show="*")
        self.T2.pack(fill="x", pady=(0, 30))

        try:
            dbRegister = Registerdata("real/Registerdatabase.db")
            register = dbRegister.fetch()
            registerUser = dbRegister.fetchUser()
        except Exception:
            register = []
            registerUser = []

        def verify():
            if self.T1.get() == "admin" and self.T2.get() == "12345@":
                controller.current_username = self.T1.get()
                text1.set('')
                text2.set('')
                controller.show_frame(SecondPage)
                return
            try:
                i = 0
                for a in range(len(register)):
                    if (register[a][1] == self.T1.get()
                            and register[a][2] == self.T2.get()):
                        controller.current_username = register[a][1]
                        controller.show_frame(SecondPage)
                        i = 1
                        break
                if i == 0:
                    messagebox.showinfo("Error",
                                        "Please provide correct username and password!!")
            except Exception:
                messagebox.showinfo("Error",
                                    "Please provide correct username and password!!")

        def Userverify():
            if self.T1.get() == "admin" and self.T2.get() == "12345@":
                controller.current_username = self.T1.get()
                user_page = controller.frames[UserPage]
                user_page.set_username(self.T1.get())
                text1.set('')
                text2.set('')
                controller.show_frame(UserPage)
                return
            try:
                i = 0
                for a in range(len(registerUser)):
                    if (registerUser[a][1] == self.T1.get()
                            and registerUser[a][2] == self.T2.get()):
                        controller.current_username = registerUser[a][1]
                        user_page = controller.frames[UserPage]
                        user_page.set_username(registerUser[a][1])
                        text1.set('')
                        text2.set('')
                        controller.show_frame(UserPage)
                        i = 1
                        break
                if i == 0:
                    messagebox.showinfo("Error",
                                        "Please provide correct username and password!!")
            except Exception:
                messagebox.showinfo("Error",
                                    "Please provide correct username and password!!")

        self.button_frame = ctk.CTkFrame(self.form_frame, fg_color="transparent")
        self.button_frame.pack(fill="x")

        ctk.CTkButton(self.button_frame, text="Admin", height=48, corner_radius=10,
                      font=ctk.CTkFont(family=MYANMAR_FONT, size=15, weight="bold"),
                      fg_color="#ef4444", hover_color="#dc2626",
                      command=verify
                      ).pack(side="left", expand=True, fill="x", padx=(0, 7))

        ctk.CTkButton(self.button_frame, text="User", height=48, corner_radius=10,
                      font=ctk.CTkFont(family=MYANMAR_FONT, size=15, weight="bold"),
                      fg_color="#f59e0b", hover_color="#d97706",
                      command=Userverify
                      ).pack(side="left", expand=True, fill="x", padx=7)

        ctk.CTkButton(self.button_frame, text="Exit", height=48, corner_radius=10,
                      font=ctk.CTkFont(family=MYANMAR_FONT, size=15, weight="bold"),
                      fg_color="#10b981", hover_color="#059669",
                      command=self.controller.destroy
                      ).pack(side="left", expand=True, fill="x", padx=(7, 0))

        ctk.CTkLabel(
            self.right_panel,
            text="King Mark Group POS System created by Magway Software House © 2026, Contact - 09777275950",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=11),
            text_color="#94a3b8"
        ).pack(side="bottom", pady=25)


# ============================================================
# SECOND PAGE - ADMIN DASHBOARD
# ============================================================

class SecondPage(ctk.CTkFrame):

    def __init__(self, parent, controller):

        super().__init__(parent, fg_color="#d2daea")
        self.controller = controller

        try:
            self.dbSales = DatabaseSales("real/SalesReportdatabase.db")
            self.dbWarehouse = warehouse("real/Warehousedatabase.db")
        except Exception as error:
            print("DB init error:", error)
            self.dbSales = None
            self.dbWarehouse = None

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#111827")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.pack(fill="x", padx=15, pady=(25, 20))

        ctk.CTkLabel(logo_frame, text="KM",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=36, weight="bold"),
                     text_color="#ffffff").pack()

        ctk.CTkLabel(logo_frame, text="King Mark Group",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=12),
                     text_color="#9ca3af").pack(pady=(0, 5))

        nav_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        nav_frame.pack(fill="both", expand=True, padx=12)

        self.create_sidebar_button(nav_frame, "Dashboard",
                                   lambda: self.show_content("dashboard"))
        self.create_sidebar_button(nav_frame, "ကုန်ပစ္စည်း",
                                   lambda: self.show_content("inventory"))
        self.create_sidebar_button(nav_frame, "အလုပ်သမားစာရင်း",
                                   lambda: self.show_content("employee"))
        self.create_sidebar_button(nav_frame, "အထွေထွေကုန်ကျစရိတ်",
                                   lambda: self.show_content("expense"))
        self.create_sidebar_button(nav_frame, "ပစ္စည်းစာရင်း",
                                   lambda: self.show_content("itemlist"))
        self.create_sidebar_button(nav_frame, "အရောင်းစာရင်း",
                                   lambda: self.show_content("salesreport"))
        self.create_sidebar_button(nav_frame, "အမြတ်စာရင်း",
                                   lambda: self.show_content("profitreport"))
        self.create_sidebar_button(nav_frame, "Logout",
                                   lambda: controller.show_frame(FirstPage))

        bottom_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        bottom_frame.pack(side="bottom", fill="x", padx=12, pady=15)

        ctk.CTkLabel(bottom_frame, text="King Mark Group POS",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=10),
                     text_color="#6b7280").pack()

        self.content_area = ctk.CTkFrame(self, fg_color="#99acd1")
        self.content_area.grid(row=0, column=1, sticky="nsew")
        self.content_area.grid_columnconfigure(0, weight=1)
        self.content_area.grid_rowconfigure(0, weight=1)

        self.show_content("dashboard")

    def create_sidebar_button(self, parent, text, command):

        button = ctk.CTkButton(
            parent, text=text, height=45, corner_radius=8,
            fg_color="transparent", hover_color="#1f2937",
            text_color="#d1d5db", anchor="w",
            font=ctk.CTkFont(family=MYANMAR_FONT, size=13, weight="bold"),
            command=command
        )
        button.pack(fill="x", pady=4)
        return button

    def show_content(self, page_name):

        for widget in self.content_area.winfo_children():
            widget.destroy()

        if page_name == "dashboard":
            self.load_dashboard_content()
        elif page_name == "inventory":
            InventoryPage(self.content_area, self.controller).pack(fill="both", expand=True)
        elif page_name == "employee":
            EmployeePage(self.content_area, self.controller).pack(fill="both", expand=True)
        elif page_name == "expense":
            ExpensePage(self.content_area, self.controller).pack(fill="both", expand=True)
        elif page_name == "itemlist":
            ItemListPage(self.content_area, self.controller).pack(fill="both", expand=True)
        elif page_name == "salesreport":
            sp = SalesPage(self.content_area, self.controller)
            sp.pack(fill="both", expand=True)
            sp.refresh_seller()
        elif page_name == "profitreport":
            ProfitPage(self.content_area, self.controller).pack(fill="both", expand=True)

    def load_dashboard_content(self):

        self.main_scroll = ctk.CTkScrollableFrame(self.content_area, fg_color="#f3f5f9")
        self.main_scroll.pack(fill="both", expand=True)

        header = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 15))

        header_left = ctk.CTkFrame(header, fg_color="transparent")
        header_left.pack(side="left")

        ctk.CTkLabel(header_left, text="POS Dashboard",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=30, weight="bold"),
                     text_color="#111827").pack(anchor="w")

        ctk.CTkLabel(header_left, text="Sales overview and inventory status",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
                     text_color="#6b7280").pack(anchor="w")

        self.date_label = ctk.CTkLabel(
            header, text=datetime.now().strftime("%d %B %Y"),
            font=ctk.CTkFont(family=MYANMAR_FONT, size=14, weight="bold"),
            text_color="#374151")
        self.date_label.pack(side="right", pady=10)

        summary_frame = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        summary_frame.pack(fill="x", padx=15, pady=(0, 15))

        for i in range(3):
            summary_frame.grid_columnconfigure(i, weight=1)

        self.sales_card = self.create_summary_card(summary_frame, 0,
            "Today's Sales", "0", "💰", "#dcfce7")
        self.orders_card = self.create_summary_card(summary_frame, 1,
            "Today's Orders", "0", "🧾", "#dbeafe")
        self.items_card = self.create_summary_card(summary_frame, 2,
            "Total Items", "0", "📦", "#fef3c7")

        chart_container = ctk.CTkFrame(self.main_scroll, fg_color="transparent")
        chart_container.pack(fill="x", padx=15, pady=(0, 15))
        chart_container.grid_columnconfigure(0, weight=3)
        chart_container.grid_columnconfigure(1, weight=1)

        top_sales_card = ctk.CTkFrame(chart_container, fg_color="#ffffff", corner_radius=14)
        top_sales_card.grid(row=0, column=0, sticky="nsew", padx=(0, 7))

        ctk.CTkLabel(top_sales_card, text="Top 10 Most Sales Items",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=20, weight="bold"),
                     text_color="#111827").pack(anchor="w", padx=20, pady=(18, 2))

        self.top_sales_canvas = tk.Canvas(top_sales_card, height=300,
                                          bg="#ffffff", highlightthickness=0)
        self.top_sales_canvas.pack(fill="x", padx=15, pady=(0, 15))

        low_stock_card = ctk.CTkFrame(chart_container, fg_color="#ffffff", corner_radius=14)
        low_stock_card.grid(row=0, column=1, sticky="nsew", padx=(7, 0))

        ctk.CTkLabel(low_stock_card, text="Low Stock",
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=20, weight="bold"),
                     text_color="#111827").pack(anchor="w", padx=18, pady=(18, 2))

        self.low_stock_canvas = tk.Canvas(low_stock_card, height=300,
                                          bg="#ffffff", highlightthickness=0)
        self.low_stock_canvas.pack(fill="both", expand=True, padx=10, pady=(0, 15))

        self.refresh_dashboard()

    def create_summary_card(self, parent, column, title, value, icon, icon_bg):

        card = ctk.CTkFrame(parent, fg_color="#ffffff", corner_radius=14, height=130)
        card.grid(row=0, column=column, sticky="ew", padx=7)
        card.grid_propagate(False)

        icon_frame = ctk.CTkFrame(card, width=55, height=55,
                                  corner_radius=12, fg_color=icon_bg)
        icon_frame.place(x=18, y=25)
        icon_frame.pack_propagate(False)
        ctk.CTkLabel(icon_frame, text=icon, font=ctk.CTkFont(size=22)).pack(expand=True)

        ctk.CTkLabel(card, text=title,
                     font=ctk.CTkFont(family=MYANMAR_FONT, size=13),
                     text_color="#6b7280").place(x=90, y=25)

        v = ctk.CTkLabel(card, text=value,
                         font=ctk.CTkFont(family=MYANMAR_FONT, size=24, weight="bold"),
                         text_color="#111827")
        v.place(x=90, y=52)
        return v

    def refresh_dashboard(self):

        today = datetime.now().strftime("%Y-%m-%d")

        try:
            result = self.dbSales.Total(today) if self.dbSales else None
            value = (result[0] or 0) if result else 0
            self.sales_card.configure(text=f"{value:,.0f}")
        except Exception:
            self.sales_card.configure(text="0")

        try:
            result = self.dbSales.CountQuan(today) if self.dbSales else None
            value = (result[0] or 0) if result else 0
            self.orders_card.configure(text=str(value))
        except Exception:
            self.orders_card.configure(text="0")

        try:
            result = self.dbWarehouse.TotalQuantity() if self.dbWarehouse else None
            value = (result[0] or 0) if result else 0
            self.items_card.configure(text=str(value))
        except Exception:
            self.items_card.configure(text="0")

        self.load_top_sales_chart()
        self.load_low_stock_chart()

    # ================================================================
    # TOP SALES (robust — handles missing 'name' column)
    # ================================================================

    def load_top_sales_chart(self):

        canvas = self.top_sales_canvas
        canvas.delete("all")
        rows = []

        try:
            if self.dbSales:
                # Check what columns exist in SalesReport
                cols = get_table_columns(self.dbSales.cur, "SalesReport")

                # Prefer 'name' if available, else 'itemcode'
                if "name" in cols:
                    label_col = "name"
                elif "itemcode" in cols:
                    label_col = "itemcode"
                else:
                    label_col = None

                if label_col and "quantity" in cols:
                    self.dbSales.cur.execute(
                        f"""SELECT {label_col}, SUM(quantity) AS total_quantity
                           FROM SalesReport
                           WHERE {label_col} IS NOT NULL AND {label_col} != ''
                           GROUP BY {label_col}
                           ORDER BY total_quantity DESC
                           LIMIT 10"""
                    )
                    rows = self.dbSales.cur.fetchall()
                else:
                    print(f"Top Sales: missing columns. Available: {cols}")
        except Exception as e:
            print("Top Sales Error:", e)

        if not rows:
            canvas.create_text(300, 140, text="No sales data available",
                               font=(MYANMAR_FONT, 14), fill="#9ca3af")
            return

        max_value = max(row[1] or 0 for row in rows)
        chart_width, left, right, top, bar_height, gap = 700, 120, 40, 10, 20, 7

        for index, row in enumerate(rows):
            item_code = str(row[0])
            quantity = row[1] or 0
            y = top + index * (bar_height + gap)

            canvas.create_text(left - 10, y + bar_height / 2, text=item_code,
                               anchor="e", font=(MYANMAR_FONT, 10, "bold"), fill="#374151")
            canvas.create_rectangle(left, y, chart_width - right, y + bar_height,
                                    fill="#eef2f7", outline="")

            bar_width = (quantity / max_value) * (chart_width - left - right) if max_value > 0 else 0

            canvas.create_rectangle(left, y, left + bar_width, y + bar_height,
                                    fill="#2563eb", outline="")
            canvas.create_text(left + bar_width + 8, y + bar_height / 2,
                               text=str(quantity), anchor="w",
                               font=(MYANMAR_FONT, 10, "bold"), fill="#111827")

    def load_low_stock_chart(self):

        canvas = self.low_stock_canvas
        canvas.delete("all")
        rows = []

        try:
            if self.dbWarehouse:
                self.dbWarehouse.cur.execute(
                    "SELECT Item, Name, Quantity FROM Order2 WHERE Quantity < 2 ORDER BY Quantity ASC"
                )
                rows = self.dbWarehouse.cur.fetchall()
        except Exception as e:
            print("Low Stock Error:", e)

        if not rows:
            canvas.create_text(150, 140, text="No low-stock items",
                               font=(MYANMAR_FONT, 13), fill="#9ca3af")
            return

        y = 10
        for row in rows:
            item_code = str(row[0])
            name = str(row[1] or "")
            quantity = row[2] or 0

            canvas.create_rectangle(8, y, 330, y + 48, fill="#fff7ed", outline="")
            canvas.create_text(20, y + 16, text=item_code, anchor="w",
                               font=(MYANMAR_FONT, 11, "bold"), fill="#111827")
            canvas.create_text(20, y + 34, text=name[:25], anchor="w",
                               font=(MYANMAR_FONT, 9), fill="#6b7280")
            canvas.create_text(310, y + 24, text=str(quantity), anchor="e",
                               font=(MYANMAR_FONT, 12, "bold"), fill="#dc2626")

            y += 55
            if y > 270:
                break

    def tkraise(self, aboveThis=None):
        result = super().tkraise(aboveThis)
        if hasattr(self, 'refresh_dashboard'):
            if hasattr(self, 'sales_card') and self.sales_card.winfo_exists():
                self.refresh_dashboard()
        return result


# ============================================================
# APPLICATION
# ============================================================

class Application(ctk.CTk):

    def __init__(self, *args, **kwargs):

        super().__init__(*args, **kwargs)

        # ====================================================
        # SET GLOBAL TK FONTS (for ttk widgets)
        # ====================================================
        self._set_global_fonts()

        self.title("KM POS Application")

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        window_width = int(screen_width * 0.90)
        window_height = int(screen_height * 0.88)
        window_width = max(window_width, 1100)
        window_height = max(window_height, 650)
        window_width = min(window_width, screen_width)
        window_height = min(window_height, screen_height - 50)

        x = int((screen_width - window_width) / 2)
        y = int((screen_height - window_height) / 2)

        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.minsize(1000, 600)

        try:
            self.iconbitmap("logoicon.ico")
        except Exception:
            pass

        self.withdraw()

        self.persistence = DataPersistence("real/KMPOS_data.db")

        self.data_store = {
            "inventory": self.persistence.load_all("inventory"),
            "employee": self.persistence.load_all("employee"),
            "expense": self.persistence.load_all("expense"),
            "itemlist": self.persistence.load_all("itemlist"),
        }

        print(f"✅ Loaded: inventory={len(self.data_store['inventory'])}, "
              f"employee={len(self.data_store['employee'])}, "
              f"expense={len(self.data_store['expense'])}, "
              f"itemlist={len(self.data_store['itemlist'])}")

        self.current_username = ""

        window = ctk.CTkFrame(self, fg_color="#f3f5f9", corner_radius=0)
        window.pack(fill="both", expand=True)
        window.grid_rowconfigure(0, weight=1)
        window.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (FirstPage, SecondPage, UserPage):
            frame = F(window, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(FirstPage)
        self.deiconify()

    # ========================================================
    # GLOBAL FONT SETTER
    # ========================================================

    def _set_global_fonts(self):
        """
        Set Tk default fonts + ttk styles to use Myanmar font.
        This ensures ttk widgets (Treeview, Scrollbar, etc.)
        render Myanmar text correctly.
        """
        import tkinter.font as tkfont
        from tkinter import ttk

        # --- Tk named fonts ---
        for font_name in [
            "TkDefaultFont",
            "TkTextFont",
            "TkMenuFont",
            "TkHeadingFont",
            "TkCaptionFont",
            "TkSmallCaptionFont",
            "TkIconFont",
            "TkTooltipFont",
        ]:
            try:
                f = tkfont.nametofont(font_name)
                f.configure(family=MYANMAR_FONT, size=11)
            except Exception as e:
                print(f"Font set {font_name} skipped: {e}")

        # TkHeadingFont — bold
        try:
            heading = tkfont.nametofont("TkHeadingFont")
            heading.configure(family=MYANMAR_FONT, size=11, weight="bold")
        except Exception:
            pass

        # --- ttk Style defaults ---
        try:
            style = ttk.Style()
            try:
                style.theme_use("clam")
            except Exception:
                pass

            style.configure(".", font=(MYANMAR_FONT, 11))
            style.configure("TButton", font=(MYANMAR_FONT, 11))
            style.configure("TLabel", font=(MYANMAR_FONT, 11))
            style.configure("TEntry", font=(MYANMAR_FONT, 11))
            style.configure("TCombobox", font=(MYANMAR_FONT, 11))
            style.configure("Treeview",
                            font=(MYANMAR_FONT, 11),
                            rowheight=35)
            style.configure("Treeview.Heading",
                            font=(MYANMAR_FONT, 11, "bold"))
            style.configure("TNotebook", font=(MYANMAR_FONT, 11))
            style.configure("TNotebook.Tab", font=(MYANMAR_FONT, 11))
        except Exception as e:
            print(f"ttk style set skipped: {e}")

        print(f"✅ Global fonts set to: {MYANMAR_FONT}")

    # ========================================================
    # SHOW FRAME
    # ========================================================

    def show_frame(self, page):
        frame = self.frames[page]
        frame.tkraise()

# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    app = Application()
    app.mainloop()