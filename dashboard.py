"""
Super-Hobby.hu Promotion Tracker - Dashboard
=============================================
 
A local desktop dashboard (Tkinter) for the SQLite database produced by
tracker.py. Run tracker.py manually or via Windows Task Scheduler to keep
the database updated (see run_tracker.bat) - this dashboard only
reads/writes the same SQLite file, it does not scrape anything itself.
 
Features
--------
- Add / edit / delete promotions manually (delete supports multi-select:
  Ctrl+click or Shift+click rows in the table, then Törlés)
- Search & filter by brand, type (daily/weekly/monthly) and year
- Export the current filtered view to Excel (.xlsx) and CSV
- One-click SQLite backup into backups/
- Hungarian-style date display (year-qualified):
    daily   -> "2026 Augusztus 13., csütörtök"
    weekly  -> "2026/KW33"
    monthly -> "2026 Augusztus"
- Analysis panel for the current filter:
    number of promotions, counts by type, average / median interval
    between promotions, last promotion, days since last promotion,
    predicted next promotion date, max / average discount
 
Requirements
------------
    pip install openpyxl
 
Uses the same database schema/migration as tracker.py, so it's safe to
open this dashboard even on an older database file - it will upgrade
it the same way tracker.py does (adds the 'year' column and removes
any duplicate weekly/monthly rows created by the previous schema).
"""
 
import csv
import shutil
import sqlite3
import statistics
import tkinter as tk
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import ttk, messagebox, filedialog
 
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
 
 
# ============================================================
# CONFIGURATION (mirrors tracker.py so both scripts share one DB)
# ============================================================
 
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_FILE = DATA_DIR / "superhobby.db"
BACKUP_DIR = BASE_DIR / "backups"
EXPORT_DIR = BASE_DIR / "exports"
 
 
# ============================================================
# HUNGARIAN DATE FORMATTING
# ============================================================
 
HU_MONTHS = [
    "január", "február", "március", "április", "május", "június",
    "július", "augusztus", "szeptember", "október", "november", "december",
]
 
HU_DAYS = [
    "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat", "vasárnap",
]
 
TYPE_LABELS_HU = {"daily": "Napi", "weekly": "Heti", "monthly": "Havi"}
TYPE_LABELS_EN_TO_INTERNAL = {v: k for k, v in TYPE_LABELS_HU.items()}
ALL_LABEL = "Összes"
 
 
def format_long_date(date_obj, year=None):
    """e.g. '2026 Augusztus 13., csütörtök'"""
    y = year if year is not None else date_obj.year
    month = HU_MONTHS[date_obj.month - 1].capitalize()
    return f"{y} {month} {date_obj.day}., {HU_DAYS[date_obj.weekday()]}"
 
 
def format_month_name(date_obj, year=None):
    """e.g. '2026 Augusztus'"""
    y = year if year is not None else date_obj.year
    month = HU_MONTHS[date_obj.month - 1].capitalize()
    return f"{y} {month}"
 
 
def format_period_display(date_str, period, promotion_type, year):
    """Return the human-friendly, year-qualified period string."""
    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
 
    if promotion_type == "daily":
        return format_long_date(date_obj, year)
    elif promotion_type == "weekly":
        return f"{year}/{period}"  # e.g. "2026/KW33"
    elif promotion_type == "monthly":
        return format_month_name(date_obj, year)
    return period
 
 
def compute_period(date_obj, promotion_type):
    """Same logic as tracker.py's save_promotions(), for manual entries."""
    if promotion_type == "daily":
        return date_obj.strftime("%Y-%m-%d")
    elif promotion_type == "weekly":
        return f"KW{date_obj.isocalendar().week:02d}"
    elif promotion_type == "monthly":
        return date_obj.strftime("%Y-%m")
    return date_obj.strftime("%Y-%m-%d")
 
 
def compute_year(date_obj, promotion_type):
    """Same logic as tracker.py: weekly uses the ISO calendar year so
    edge-of-year weeks are filed under the correct year."""
    if promotion_type == "weekly":
        return date_obj.isocalendar().year
    return date_obj.year
 
 
# ============================================================
# DATABASE LAYER (schema + migration mirrored from tracker.py)
# ============================================================
 
def get_table_columns(conn):
    if not conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='promotions'"
    ).fetchone():
        return []
    return [row[1] for row in conn.execute("PRAGMA table_info(promotions)").fetchall()]
 
 
def get_unique_index_column_sets(conn):
    """Reads SQLite's own schema introspection (rather than parsing the
    CREATE TABLE text) to find every UNIQUE index's column set."""
    result = []
    for idx in conn.execute("PRAGMA index_list(promotions)").fetchall():
        is_unique = idx[2]
        index_name = idx[1]
        if not is_unique:
            continue
        cols = [info[2] for info in conn.execute(f"PRAGMA index_info('{index_name}')").fetchall()]
        result.append(set(cols))
    return result
 
 
def table_needs_migration(conn):
    columns = get_table_columns(conn)
    if not columns:
        return False  # table doesn't exist yet
 
    if "year" not in columns:
        return True
 
    target = {"year", "period", "promotion_type", "brand"}
    return not any(cols == target for cols in get_unique_index_column_sets(conn))
 
 
def migrate_database(conn):
    if not table_needs_migration(conn):
        return
 
    existing_cols = get_table_columns(conn)
    if "year" not in existing_cols:
        conn.execute("ALTER TABLE promotions ADD COLUMN year INTEGER")
        conn.commit()
 
    conn.execute("""
        UPDATE promotions
        SET year = CAST(substr(date, 1, 4) AS INTEGER)
        WHERE year IS NULL
    """)
    conn.commit()
 
    conn.execute("""
        CREATE TABLE promotions_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            year INTEGER NOT NULL,
            period TEXT NOT NULL,
            promotion_type TEXT NOT NULL,
            brand TEXT NOT NULL,
            discount INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(year, period, promotion_type, brand)
        )
    """)
 
    conn.execute("""
        INSERT INTO promotions_new
            (id, date, year, period, promotion_type, brand, discount, created_at)
        SELECT id, date, year, period, promotion_type, brand, discount, created_at
        FROM promotions
        WHERE id IN (
            SELECT MIN(id)
            FROM promotions
            GROUP BY year, period, promotion_type, brand
        )
    """)
 
    conn.execute("DROP TABLE promotions")
    conn.execute("ALTER TABLE promotions_new RENAME TO promotions")
    conn.commit()
 
 
def get_connection():
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS promotions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            year INTEGER NOT NULL,
            period TEXT NOT NULL,
            promotion_type TEXT NOT NULL,
            brand TEXT NOT NULL,
            discount INTEGER NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(year, period, promotion_type, brand)
        )
    """)
    conn.commit()
    migrate_database(conn)
 
    # Belt-and-suspenders: never leave a NULL year in the table.
    conn.execute("""
        UPDATE promotions
        SET year = CAST(substr(date, 1, 4) AS INTEGER)
        WHERE year IS NULL
    """)
    conn.commit()
 
    return conn
 
 
# Row tuple layout used throughout this file:
# (id, date, year, period, promotion_type, brand, discount, created_at)
 
def fetch_promotions(conn, brand=None, promotion_type=None, year=None):
    query = (
        "SELECT id, date, year, period, promotion_type, brand, discount, created_at "
        "FROM promotions WHERE 1=1"
    )
    params = []
 
    if brand and brand != ALL_LABEL:
        query += " AND brand = ?"
        params.append(brand)
 
    if promotion_type and promotion_type != ALL_LABEL:
        query += " AND promotion_type = ?"
        params.append(promotion_type)
 
    if year and year != ALL_LABEL:
        query += " AND year = ?"
        params.append(int(year))
 
    query += " ORDER BY date DESC, promotion_type, brand"
 
    return conn.execute(query, params).fetchall()
 
 
def get_distinct_brands(conn):
    rows = conn.execute("SELECT DISTINCT brand FROM promotions ORDER BY brand").fetchall()
    return [r[0] for r in rows]
 
 
def get_distinct_years(conn):
    rows = conn.execute("SELECT DISTINCT year FROM promotions ORDER BY year DESC").fetchall()
    return [str(r[0]) for r in rows]
 
 
def insert_promotion(conn, date_obj, promotion_type, brand, discount):
    period = compute_period(date_obj, promotion_type)
    year = compute_year(date_obj, promotion_type)
    created_at = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        """
        INSERT INTO promotions (date, year, period, promotion_type, brand, discount, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (date_obj.strftime("%Y-%m-%d"), year, period, promotion_type, brand, discount, created_at),
    )
    conn.commit()
 
 
def update_promotion(conn, row_id, date_obj, promotion_type, brand, discount):
    period = compute_period(date_obj, promotion_type)
    year = compute_year(date_obj, promotion_type)
    conn.execute(
        """
        UPDATE promotions
        SET date = ?, year = ?, period = ?, promotion_type = ?, brand = ?, discount = ?
        WHERE id = ?
        """,
        (date_obj.strftime("%Y-%m-%d"), year, period, promotion_type, brand, discount, row_id),
    )
    conn.commit()
 
 
def delete_promotion(conn, row_id):
    conn.execute("DELETE FROM promotions WHERE id = ?", (row_id,))
    conn.commit()
 
 
# ============================================================
# ANALYSIS
# ============================================================
 
def compute_stats(rows):
    """rows: list of (id, date, year, period, promotion_type, brand, discount, created_at)"""
    stats = {
        "count": len(rows),
        "by_type": Counter(),
        "avg_interval": None,
        "median_interval": None,
        "last_date": None,
        "days_since_last": None,
        "predicted_next": None,
        "max_discount": None,
        "avg_discount": None,
    }
 
    if not rows:
        return stats
 
    stats["by_type"] = Counter(r[4] for r in rows)
 
    dates = sorted({datetime.strptime(r[1], "%Y-%m-%d").date() for r in rows})
 
    if len(dates) >= 2:
        diffs = [(dates[i + 1] - dates[i]).days for i in range(len(dates) - 1)]
        stats["avg_interval"] = statistics.mean(diffs)
        stats["median_interval"] = statistics.median(diffs)
 
    last_date = dates[-1]
    stats["last_date"] = last_date
    stats["days_since_last"] = (datetime.now().date() - last_date).days
 
    if stats["avg_interval"] is not None:
        stats["predicted_next"] = last_date + timedelta(days=round(stats["avg_interval"]))
 
    discounts = [r[6] for r in rows]
    stats["max_discount"] = max(discounts)
    stats["avg_discount"] = statistics.mean(discounts)
 
    return stats
 
 
# ============================================================
# EXPORT
# ============================================================
 
def export_csv(rows, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["ID", "Dátum", "Időszak", "Típus", "Márka", "Kedvezmény (%)"])
        for r in rows:
            row_id, date_str, year, period, promotion_type, brand, discount, _ = r
            writer.writerow([
                row_id,
                date_str,
                format_period_display(date_str, period, promotion_type, year),
                TYPE_LABELS_HU.get(promotion_type, promotion_type),
                brand,
                discount,
            ])
 
 
def export_excel(rows, path):
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl nincs telepítve. Telepítés: pip install openpyxl")
 
    wb = Workbook()
    ws = wb.active
    ws.title = "Promóciók"
 
    headers = ["ID", "Dátum", "Időszak", "Típus", "Márka", "Kedvezmény (%)"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
 
    for r in rows:
        row_id, date_str, year, period, promotion_type, brand, discount, _ = r
        ws.append([
            row_id,
            date_str,
            format_period_display(date_str, period, promotion_type, year),
            TYPE_LABELS_HU.get(promotion_type, promotion_type),
            brand,
            discount,
        ])
 
    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[col_cells[0].column_letter].width = max_len + 3
 
    wb.save(path)
 
 
def backup_database():
    BACKUP_DIR.mkdir(exist_ok=True)
    if not DB_FILE.exists():
        raise FileNotFoundError("Az adatbázis fájl még nem létezik.")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"superhobby_backup_{timestamp}.db"
    shutil.copy2(DB_FILE, backup_path)
    return backup_path
 
 
# ============================================================
# ADD / EDIT DIALOG
# ============================================================
 
class PromotionDialog(tk.Toplevel):
    def __init__(self, parent, title, existing=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.grab_set()
 
        # existing row layout: (id, date, year, period, promotion_type, brand, discount, created_at)
        existing_date = existing[1] if existing else datetime.now().strftime("%Y-%m-%d")
        existing_type = existing[4] if existing else "daily"
        existing_brand = existing[5] if existing else ""
        existing_discount = existing[6] if existing else ""
 
        frame = ttk.Frame(self, padding=15)
        frame.grid(row=0, column=0)
 
        ttk.Label(frame, text="Dátum (ÉÉÉÉ-HH-NN):").grid(row=0, column=0, sticky="w", pady=4)
        self.date_var = tk.StringVar(value=existing_date)
        ttk.Entry(frame, textvariable=self.date_var, width=20).grid(row=0, column=1, pady=4)
 
        ttk.Label(frame, text="Típus:").grid(row=1, column=0, sticky="w", pady=4)
        self.type_var = tk.StringVar(value=TYPE_LABELS_HU.get(existing_type, "Napi"))
        ttk.Combobox(
            frame, textvariable=self.type_var, values=list(TYPE_LABELS_HU.values()),
            state="readonly", width=18,
        ).grid(row=1, column=1, pady=4)
 
        ttk.Label(frame, text="Márka:").grid(row=2, column=0, sticky="w", pady=4)
        self.brand_var = tk.StringVar(value=existing_brand)
        ttk.Entry(frame, textvariable=self.brand_var, width=20).grid(row=2, column=1, pady=4)
 
        ttk.Label(frame, text="Kedvezmény (%):").grid(row=3, column=0, sticky="w", pady=4)
        self.discount_var = tk.StringVar(value=str(existing_discount))
        ttk.Entry(frame, textvariable=self.discount_var, width=20).grid(row=3, column=1, pady=4)
 
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(btn_frame, text="Mentés", command=self._on_save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Mégse", command=self.destroy).pack(side="left", padx=5)
 
        self.bind("<Return>", lambda e: self._on_save())
 
    def _on_save(self):
        try:
            date_obj = datetime.strptime(self.date_var.get().strip(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Hiba", "A dátum formátuma: ÉÉÉÉ-HH-NN (pl. 2026-08-13)")
            return
 
        promotion_type = TYPE_LABELS_EN_TO_INTERNAL.get(self.type_var.get())
        if not promotion_type:
            messagebox.showerror("Hiba", "Válassz típust.")
            return
 
        brand = self.brand_var.get().strip()
        if not brand:
            messagebox.showerror("Hiba", "A márka mező nem lehet üres.")
            return
 
        try:
            discount = int(self.discount_var.get().strip())
            if not (0 <= discount <= 100):
                raise ValueError
        except ValueError:
            messagebox.showerror("Hiba", "A kedvezménynek 0 és 100 közötti egész számnak kell lennie.")
            return
 
        self.result = (date_obj, promotion_type, brand, discount)
        self.destroy()
 
 
# ============================================================
# MAIN DASHBOARD WINDOW
# ============================================================
 
class Dashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Super-Hobby.hu - Promóció Dashboard")
        self.geometry("1150x680")
        self.minsize(950, 600)
 
        self.conn = get_connection()
        self.current_rows = []
 
        self._build_filter_bar()
        self._build_table()
        self._build_action_bar()
        self._build_stats_panel()
 
        self.refresh()
 
    # --------------------------------------------------------
    # UI construction
    # --------------------------------------------------------
 
    def _build_filter_bar(self):
        bar = ttk.Frame(self, padding=(10, 10, 10, 0))
        bar.pack(fill="x")
 
        ttk.Label(bar, text="Márka:").pack(side="left")
        self.brand_filter = tk.StringVar(value=ALL_LABEL)
        self.brand_combo = ttk.Combobox(bar, textvariable=self.brand_filter, state="readonly", width=20)
        self.brand_combo.pack(side="left", padx=(4, 15))
 
        ttk.Label(bar, text="Típus:").pack(side="left")
        self.type_filter = tk.StringVar(value=ALL_LABEL)
        self.type_combo = ttk.Combobox(
            bar, textvariable=self.type_filter, state="readonly", width=12,
            values=[ALL_LABEL] + list(TYPE_LABELS_HU.values()),
        )
        self.type_combo.pack(side="left", padx=(4, 15))
 
        ttk.Label(bar, text="Év:").pack(side="left")
        self.year_filter = tk.StringVar(value=ALL_LABEL)
        self.year_combo = ttk.Combobox(bar, textvariable=self.year_filter, state="readonly", width=10)
        self.year_combo.pack(side="left", padx=(4, 15))
 
        ttk.Button(bar, text="Szűrés", command=self.refresh).pack(side="left", padx=4)
        ttk.Button(bar, text="Szűrők törlése", command=self._reset_filters).pack(side="left")
 
    def _build_table(self):
        table_frame = ttk.Frame(self, padding=10)
        table_frame.pack(fill="both", expand=True)
 
        columns = ("id", "date_display", "type", "brand", "discount")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
 
        headings = {
            "id": ("ID", 50),
            "date_display": ("Dátum / Időszak", 300),
            "type": ("Típus", 90),
            "brand": ("Márka", 260),
            "discount": ("Kedvezmény", 100),
        }
        for col, (text, width) in headings.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="w")
 
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
 
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
 
        # id (str) -> full DB row, for edit/delete lookups
        self.row_lookup = {}
 
    def _build_action_bar(self):
        bar = ttk.Frame(self, padding=(10, 0, 10, 10))
        bar.pack(fill="x")
 
        ttk.Button(bar, text="+ Hozzáadás", command=self.add_promotion).pack(side="left", padx=3)
        ttk.Button(bar, text="Szerkesztés", command=self.edit_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="Törlés (Ctrl/Shift = több)", command=self.delete_selected).pack(side="left", padx=3)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(bar, text="Export CSV", command=self.do_export_csv).pack(side="left", padx=3)
        ttk.Button(bar, text="Export Excel", command=self.do_export_excel).pack(side="left", padx=3)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(bar, text="Adatbázis mentése (backup)", command=self.do_backup).pack(side="left", padx=3)
        ttk.Button(bar, text="Frissítés", command=self.refresh).pack(side="right", padx=3)
 
    def _build_stats_panel(self):
        panel = ttk.LabelFrame(self, text="Elemzés (a jelenlegi szűrésre)", padding=10)
        panel.pack(fill="x", padx=10, pady=(0, 10))
 
        self.stats_labels = {}
        fields = [
            ("count", "Promóciók száma:"),
            ("by_type", "Napi / Heti / Havi:"),
            ("avg_interval", "Átlagos időköz (nap):"),
            ("median_interval", "Medián időköz (nap):"),
            ("last_date", "Utolsó promóció:"),
            ("days_since_last", "Napok az utolsó óta:"),
            ("predicted_next", "Következő becsült dátum:"),
            ("max_discount", "Maximális kedvezmény:"),
            ("avg_discount", "Átlagos kedvezmény:"),
        ]
 
        for i, (key, label) in enumerate(fields):
            row, col = divmod(i, 3)
            cell = ttk.Frame(panel)
            cell.grid(row=row, column=col, sticky="w", padx=15, pady=3)
            ttk.Label(cell, text=label, font=("", 9, "bold")).pack(anchor="w")
            value_lbl = ttk.Label(cell, text="-")
            value_lbl.pack(anchor="w")
            self.stats_labels[key] = value_lbl
 
    # --------------------------------------------------------
    # Data refresh
    # --------------------------------------------------------
 
    def refresh(self):
        brands = [ALL_LABEL] + get_distinct_brands(self.conn)
        years = [ALL_LABEL] + get_distinct_years(self.conn)
        self.brand_combo["values"] = brands
        self.year_combo["values"] = years
        if self.brand_filter.get() not in brands:
            self.brand_filter.set(ALL_LABEL)
        if self.year_filter.get() not in years:
            self.year_filter.set(ALL_LABEL)
 
        promotion_type = TYPE_LABELS_EN_TO_INTERNAL.get(self.type_filter.get())
 
        rows = fetch_promotions(
            self.conn,
            brand=self.brand_filter.get(),
            promotion_type=promotion_type,
            year=self.year_filter.get(),
        )
        self.current_rows = rows
 
        self.tree.delete(*self.tree.get_children())
        self.row_lookup = {}
        for r in rows:
            row_id, date_str, year, period, promotion_type, brand, discount, _ = r
            display_date = format_period_display(date_str, period, promotion_type, year)
            self.tree.insert(
                "", "end", iid=str(row_id),
                values=(row_id, display_date, TYPE_LABELS_HU.get(promotion_type, promotion_type), brand, f"{discount}%"),
            )
            self.row_lookup[str(row_id)] = r
 
        self._update_stats(rows)
 
    def _reset_filters(self):
        self.brand_filter.set(ALL_LABEL)
        self.type_filter.set(ALL_LABEL)
        self.year_filter.set(ALL_LABEL)
        self.refresh()
 
    def _update_stats(self, rows):
        stats = compute_stats(rows)
        self.stats_labels["count"].config(text=str(stats["count"]))
 
        by_type = stats["by_type"]
        self.stats_labels["by_type"].config(
            text=f"{by_type.get('daily', 0)} / {by_type.get('weekly', 0)} / {by_type.get('monthly', 0)}"
        )
 
        self.stats_labels["avg_interval"].config(
            text=f"{stats['avg_interval']:.1f}" if stats["avg_interval"] is not None else "-"
        )
        self.stats_labels["median_interval"].config(
            text=f"{stats['median_interval']:.1f}" if stats["median_interval"] is not None else "-"
        )
        self.stats_labels["last_date"].config(
            text=format_long_date(stats["last_date"]) if stats["last_date"] else "-"
        )
        self.stats_labels["days_since_last"].config(
            text=str(stats["days_since_last"]) if stats["days_since_last"] is not None else "-"
        )
        self.stats_labels["predicted_next"].config(
            text=format_long_date(stats["predicted_next"]) if stats["predicted_next"] else "-"
        )
        self.stats_labels["max_discount"].config(
            text=f"{stats['max_discount']}%" if stats["max_discount"] is not None else "-"
        )
        self.stats_labels["avg_discount"].config(
            text=f"{stats['avg_discount']:.1f}%" if stats["avg_discount"] is not None else "-"
        )
 
    # --------------------------------------------------------
    # CRUD actions
    # --------------------------------------------------------
 
    def _get_selected_rows(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Nincs kijelölés", "Előbb válassz ki legalább egy sort a táblázatban.")
            return []
        return [self.row_lookup[iid] for iid in selection if iid in self.row_lookup]
 
    def add_promotion(self):
        dialog = PromotionDialog(self, "Promóció hozzáadása")
        self.wait_window(dialog)
        if dialog.result:
            date_obj, promotion_type, brand, discount = dialog.result
            try:
                insert_promotion(self.conn, date_obj, promotion_type, brand, discount)
            except sqlite3.IntegrityError:
                messagebox.showerror("Hiba", "Ez a promóció (év + időszak + típus + márka) már létezik.")
                return
            self.refresh()
 
    def edit_selected(self):
        rows = self._get_selected_rows()
        if not rows:
            return
        if len(rows) > 1:
            messagebox.showinfo("Több sor kijelölve", "Szerkesztéshez csak egy sort jelölj ki.")
            return
        row = rows[0]
        dialog = PromotionDialog(self, "Promóció szerkesztése", existing=row)
        self.wait_window(dialog)
        if dialog.result:
            date_obj, promotion_type, brand, discount = dialog.result
            try:
                update_promotion(self.conn, row[0], date_obj, promotion_type, brand, discount)
            except sqlite3.IntegrityError:
                messagebox.showerror("Hiba", "Ez a promóció (év + időszak + típus + márka) már létezik.")
                return
            self.refresh()
 
    def delete_selected(self):
        rows = self._get_selected_rows()
        if not rows:
            return
 
        if len(rows) == 1:
            row = rows[0]
            message = f"Biztosan törlöd ezt a bejegyzést?\n\n{row[5]} - {row[6]}% ({row[1]})"
        else:
            preview = "\n".join(f"• {r[5]} - {r[6]}% ({r[1]})" for r in rows[:10])
            more = f"\n... és még {len(rows) - 10} bejegyzés" if len(rows) > 10 else ""
            message = f"Biztosan törlöd ezt a {len(rows)} bejegyzést?\n\n{preview}{more}"
 
        confirm = messagebox.askyesno("Megerősítés", message)
        if confirm:
            for row in rows:
                delete_promotion(self.conn, row[0])
            self.refresh()
 
    # --------------------------------------------------------
    # Export / backup actions
    # --------------------------------------------------------
 
    def do_export_csv(self):
        if not self.current_rows:
            messagebox.showinfo("Nincs adat", "Nincs exportálható adat a jelenlegi szűréssel.")
            return
        EXPORT_DIR.mkdir(exist_ok=True)
        default_name = f"promociok_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        path = filedialog.asksaveasfilename(
            initialdir=EXPORT_DIR, initialfile=default_name,
            defaultextension=".csv", filetypes=[("CSV fájl", "*.csv")],
        )
        if not path:
            return
        export_csv(self.current_rows, path)
        messagebox.showinfo("Kész", f"Exportálva:\n{path}")
 
    def do_export_excel(self):
        if not self.current_rows:
            messagebox.showinfo("Nincs adat", "Nincs exportálható adat a jelenlegi szűréssel.")
            return
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Hiányzó függőség", "Az Excel exporthoz telepítsd: pip install openpyxl")
            return
        EXPORT_DIR.mkdir(exist_ok=True)
        default_name = f"promociok_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(
            initialdir=EXPORT_DIR, initialfile=default_name,
            defaultextension=".xlsx", filetypes=[("Excel fájl", "*.xlsx")],
        )
        if not path:
            return
        export_excel(self.current_rows, path)
        messagebox.showinfo("Kész", f"Exportálva:\n{path}")
 
    def do_backup(self):
        try:
            path = backup_database()
        except FileNotFoundError as e:
            messagebox.showerror("Hiba", str(e))
            return
        messagebox.showinfo("Mentés kész", f"Adatbázis mentve ide:\n{path}")
 
 
# ============================================================
# ENTRY POINT
# ============================================================
 
def main():
    app = Dashboard()
    app.mainloop()
 
 
if __name__ == "__main__":
    main()