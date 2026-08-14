"""
Super-Hobby.hu Promotion Dashboard

A local desktop dashboard (Tkinter) for the SQLite database produced by
tracker.py.

The dashboard lets you browse for and select the SQLite database.
The selected database path is remembered between launches.

The database can therefore be stored anywhere, including a Google Drive
folder, while the EXE remains completely independent of the database location.

Features:
- Browse/select database on first launch
- Remember selected database
- Change database at any time
- Add / edit / delete promotions manually
- Delete supports multi-select:
  Ctrl+click or Shift+click rows in the table, then Törlés
- Search/filter by brand, type (daily/weekly/monthly) and year
- Export current filtered view to Excel (.xlsx) and CSV
- One-click SQLite backup
- Hungarian-style date display
- Analysis panel for current filter

Requirements:
    pip install openpyxl
"""

import csv
import json
import shutil
import sqlite3
import statistics
import tkinter as tk
from calendar import monthrange

from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import ttk, messagebox, filedialog


# ============================================================
# OPTIONAL EXCEL SUPPORT
# ============================================================

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ============================================================
# APPLICATION PATHS / CONFIGURATION
# ============================================================

# Folder containing the Python script or EXE.
BASE_DIR = Path(__file__).resolve().parent

# These remain relative to the application.
EXPORT_DIR = BASE_DIR / "exports"
BACKUP_DIR = BASE_DIR / "backups"

# ------------------------------------------------------------
# Persistent application configuration
#
# We store this in the user's AppData folder rather than beside
# the EXE. This avoids permission problems if the EXE is located
# in Program Files or another protected directory.
# ------------------------------------------------------------

APP_DATA_DIR = Path.home() / "AppData" / "Roaming" / "SuperHobbyDashboard"
CONFIG_FILE = APP_DATA_DIR / "config.json"


# ============================================================
# DATABASE SELECTION / CONFIGURATION
# ============================================================

def load_saved_database_path():
    """
    Load the previously selected database path.

    Returns:
        Path or None
    """

    try:
        if not CONFIG_FILE.exists():
            return None

        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)

        db_path = config.get("database_path")

        if not db_path:
            return None

        return Path(db_path)

    except (OSError, json.JSONDecodeError):
        return None


def save_database_path(db_path):
    """
    Save the selected database path for future launches.
    """

    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)

    config = {
        "database_path": str(Path(db_path).resolve())
    }

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)


def choose_database(parent=None, initial_dir=None):
    """
    Show a file browser and let the user select a SQLite database.

    Returns:
        Path or None if cancelled.
    """

    if initial_dir is None:
        initial_dir = str(BASE_DIR)

    selected = filedialog.askopenfilename(
        parent=parent,
        title="Válaszd ki a Super-Hobby adatbázist",
        initialdir=initial_dir,
        filetypes=[
            ("SQLite adatbázis", "*.db *.sqlite *.sqlite3"),
            ("Minden fájl", "*.*"),
        ],
    )

    if not selected:
        return None

    return Path(selected).resolve()


def get_database_path(parent=None):
    """
    Determine which database to use.

    1. Try previously saved database.
    2. If it exists, use it.
    3. Otherwise ask the user to select one.
    """

    saved_path = load_saved_database_path()

    if saved_path and saved_path.exists():
        return saved_path

    if saved_path:
        messagebox.showwarning(
            "Adatbázis nem található",
            "A korábban kiválasztott adatbázis nem található:\n\n"
            f"{saved_path}\n\n"
            "Válassz egy új adatbázist."
        )

    selected = choose_database(parent)

    if selected is None:
        return None

    save_database_path(selected)

    return selected


# ============================================================
# HUNGARIAN DATE FORMATTING
# ============================================================

HU_MONTHS = [
    "január",
    "február",
    "március",
    "április",
    "május",
    "június",
    "július",
    "augusztus",
    "szeptember",
    "október",
    "november",
    "december",
]

HU_DAYS = [
    "hétfő",
    "kedd",
    "szerda",
    "csütörtök",
    "péntek",
    "szombat",
    "vasárnap",
]

TYPE_LABELS_HU = {
    "daily": "Napi",
    "weekly": "Heti",
    "monthly": "Havi",
}

TYPE_LABELS_EN_TO_INTERNAL = {
    v: k for k, v in TYPE_LABELS_HU.items()
}

ALL_LABEL = "Összes"


def strip_count_suffix(display_text):
    """'Revell (2)' -> 'Revell'. Used to recover the raw brand name
    (or ALL_LABEL) from what's shown in the brand dropdown, which is
    labelled with an entry count."""

    if display_text.endswith(")") and " (" in display_text:
        base, _, suffix = display_text.rpartition(" (")
        if suffix[:-1].isdigit():
            return base
    return display_text


def format_long_date(date_obj, year=None):
    """Example: '2026 Augusztus 13., csütörtök'"""

    y = year if year is not None else date_obj.year

    month = HU_MONTHS[date_obj.month - 1].capitalize()

    return f"{y} {month} {date_obj.day}., {HU_DAYS[date_obj.weekday()]}"


def format_month_name(date_obj, year=None):
    """Example: '2026 Augusztus'"""

    y = year if year is not None else date_obj.year

    month = HU_MONTHS[date_obj.month - 1].capitalize()

    return f"{y} {month}"


def format_period_display(date_str, period, promotion_type, year):
    """Return human-friendly, year-qualified period string."""

    date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()

    if promotion_type == "daily":
        return format_long_date(date_obj, year)

    elif promotion_type == "weekly":
        return f"{year}/{period}"

    elif promotion_type == "monthly":
        return format_month_name(date_obj, year)

    return period


def format_date_span(start_date, end_date):
    """Example: '1 év 2 hónap 6 nap (összesen 432 nap)'

    Counts INCLUSIVELY - if the database has entries dated Aug 12,
    13 and 14, that's 3 days of data, not 2 (Aug 14 minus Aug 12).
    So the end date is treated as if it were one day later for the
    purposes of this calculation."""

    effective_end = end_date + timedelta(days=1)

    total_days = (effective_end - start_date).days

    years = effective_end.year - start_date.year
    months = effective_end.month - start_date.month
    days = effective_end.day - start_date.day

    if days < 0:
        months -= 1
        prev_month = effective_end.month - 1
        prev_year = effective_end.year
        if prev_month == 0:
            prev_month = 12
            prev_year -= 1
        days += monthrange(prev_year, prev_month)[1]

    if months < 0:
        years -= 1
        months += 12

    parts = []
    if years:
        parts.append(f"{years} év")
    if months:
        parts.append(f"{months} hónap")
    parts.append(f"{days} nap")

    return f"{' '.join(parts)} (összesen {total_days} nap)"


# ============================================================
# PERIOD CALCULATIONS
# ============================================================

def compute_period(date_obj, promotion_type):
    """Same logic as tracker.py."""

    if promotion_type == "daily":
        return date_obj.strftime("%Y-%m-%d")

    elif promotion_type == "weekly":
        return f"KW{date_obj.isocalendar().week:02d}"

    elif promotion_type == "monthly":
        return date_obj.strftime("%Y-%m")

    return date_obj.strftime("%Y-%m-%d")


def compute_year(date_obj, promotion_type):
    """
    Weekly promotions use the ISO calendar year.
    """

    if promotion_type == "weekly":
        return date_obj.isocalendar().year

    return date_obj.year


# ============================================================
# DATABASE LAYER
# ============================================================

def get_table_columns(conn):
    if not conn.execute(
        "SELECT name FROM sqlite_master "
        "WHERE type='table' AND name='promotions'"
    ).fetchone():
        return []

    return [
        row[1]
        for row in conn.execute(
            "PRAGMA table_info(promotions)"
        ).fetchall()
    ]


def get_unique_index_column_sets(conn):
    """
    Reads SQLite schema information to find UNIQUE indexes.
    """

    result = []

    for idx in conn.execute(
        "PRAGMA index_list(promotions)"
    ).fetchall():

        is_unique = idx[2]
        index_name = idx[1]

        if not is_unique:
            continue

        cols = [
            info[2]
            for info in conn.execute(
                f"PRAGMA index_info('{index_name}')"
            ).fetchall()
        ]

        result.append(set(cols))

    return result


def table_needs_migration(conn):

    columns = get_table_columns(conn)

    if not columns:
        return False

    if "year" not in columns:
        return True

    target = {
        "year",
        "period",
        "promotion_type",
        "brand",
    }

    return not any(
        cols == target
        for cols in get_unique_index_column_sets(conn)
    )


def migrate_database(conn):

    if not table_needs_migration(conn):
        return

    existing_cols = get_table_columns(conn)

    if "year" not in existing_cols:

        conn.execute(
            "ALTER TABLE promotions ADD COLUMN year INTEGER"
        )

        conn.commit()

    conn.execute(
        """
        UPDATE promotions
        SET year = CAST(substr(date, 1, 4) AS INTEGER)
        WHERE year IS NULL
        """
    )

    conn.commit()

    conn.execute(
        """
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
        """
    )

    conn.execute(
        """
        INSERT INTO promotions_new
            (id, date, year, period, promotion_type,
             brand, discount, created_at)

        SELECT
            id,
            date,
            year,
            period,
            promotion_type,
            brand,
            discount,
            created_at

        FROM promotions

        WHERE id IN (
            SELECT MIN(id)
            FROM promotions
            GROUP BY year, period, promotion_type, brand
        )
        """
    )

    conn.execute("DROP TABLE promotions")

    conn.execute(
        "ALTER TABLE promotions_new RENAME TO promotions"
    )

    conn.commit()


def get_connection(db_path):
    """
    Open the selected database.

    IMPORTANT:
    This function no longer silently creates a new database file
    at a hard-coded location.

    The database path must already exist.
    """

    db_path = Path(db_path)

    if not db_path.exists():
        raise FileNotFoundError(
            f"Az adatbázis nem található:\n{db_path}"
        )

    if not db_path.is_file():
        raise FileNotFoundError(
            f"A kiválasztott útvonal nem fájl:\n{db_path}"
        )

    conn = sqlite3.connect(str(db_path))

    # Create the table only if the selected database is a valid
    # SQLite database but does not yet contain the promotions table.
    conn.execute(
        """
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
        """
    )

    conn.commit()

    migrate_database(conn)

    # Belt-and-suspenders:
    # never leave a NULL year in the table.
    conn.execute(
        """
        UPDATE promotions
        SET year = CAST(substr(date, 1, 4) AS INTEGER)
        WHERE year IS NULL
        """
    )

    conn.commit()

    return conn


# ============================================================
# DATABASE QUERIES
# ============================================================

def fetch_promotions(
    conn,
    brand=None,
    promotion_type=None,
    year=None
):

    query = (
        "SELECT id, date, year, period, promotion_type, "
        "brand, discount, created_at "
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

    query += (
        " ORDER BY date DESC, promotion_type, brand"
    )

    return conn.execute(query, params).fetchall()


def get_distinct_brands(conn):

    rows = conn.execute(
        "SELECT DISTINCT brand "
        "FROM promotions "
        "ORDER BY brand"
    ).fetchall()

    return [r[0] for r in rows]


def get_brand_counts(conn):
    """List of (brand, count) across the whole database, e.g.
    [('Revell', 2), ('Vallejo', 5)] - used to label the brand
    dropdown with how many entries each brand has."""

    rows = conn.execute(
        "SELECT brand, COUNT(*) "
        "FROM promotions "
        "GROUP BY brand "
        "ORDER BY brand"
    ).fetchall()

    return [(r[0], r[1]) for r in rows]


def get_database_span_text(conn):
    """Human-readable coverage span (oldest to newest date) across
    the ENTIRE database - deliberately not affected by the current
    brand/type/year filter, since 'how many days are tracked' is a
    property of the whole dataset, not of whatever slice is shown."""

    row = conn.execute(
        "SELECT MIN(date), MAX(date) FROM promotions"
    ).fetchone()

    if not row or row[0] is None or row[1] is None:
        return None

    start_date = datetime.strptime(row[0], "%Y-%m-%d").date()
    end_date = datetime.strptime(row[1], "%Y-%m-%d").date()

    return format_date_span(start_date, end_date)


def get_distinct_years(conn):

    rows = conn.execute(
        "SELECT DISTINCT year "
        "FROM promotions "
        "ORDER BY year DESC"
    ).fetchall()

    return [str(r[0]) for r in rows]


def insert_promotion(
    conn,
    date_obj,
    promotion_type,
    brand,
    discount
):

    period = compute_period(
        date_obj,
        promotion_type
    )

    year = compute_year(
        date_obj,
        promotion_type
    )

    created_at = datetime.now().isoformat(
        timespec="seconds"
    )

    conn.execute(
        """
        INSERT INTO promotions
            (
                date,
                year,
                period,
                promotion_type,
                brand,
                discount,
                created_at
            )

        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            date_obj.strftime("%Y-%m-%d"),
            year,
            period,
            promotion_type,
            brand,
            discount,
            created_at,
        ),
    )

    conn.commit()


def update_promotion(
    conn,
    row_id,
    date_obj,
    promotion_type,
    brand,
    discount
):

    period = compute_period(
        date_obj,
        promotion_type
    )

    year = compute_year(
        date_obj,
        promotion_type
    )

    conn.execute(
        """
        UPDATE promotions

        SET
            date = ?,
            year = ?,
            period = ?,
            promotion_type = ?,
            brand = ?,
            discount = ?

        WHERE id = ?
        """,
        (
            date_obj.strftime("%Y-%m-%d"),
            year,
            period,
            promotion_type,
            brand,
            discount,
            row_id,
        ),
    )

    conn.commit()


def delete_promotion(conn, row_id):

    conn.execute(
        "DELETE FROM promotions WHERE id = ?",
        (row_id,)
    )

    conn.commit()


# ============================================================
# ANALYSIS
# ============================================================

def compute_stats(rows):

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

    stats["by_type"] = Counter(
        r[4] for r in rows
    )

    dates = sorted(
        {
            datetime.strptime(
                r[1],
                "%Y-%m-%d"
            ).date()
            for r in rows
        }
    )

    if len(dates) >= 2:

        diffs = [
            (dates[i + 1] - dates[i]).days
            for i in range(len(dates) - 1)
        ]

        stats["avg_interval"] = statistics.mean(
            diffs
        )

        stats["median_interval"] = statistics.median(
            diffs
        )

    last_date = dates[-1]

    stats["last_date"] = last_date

    stats["days_since_last"] = (
        datetime.now().date() - last_date
    ).days

    if stats["avg_interval"] is not None:

        stats["predicted_next"] = (
            last_date
            + timedelta(
                days=round(
                    stats["avg_interval"]
                )
            )
        )

    discounts = [
        r[6]
        for r in rows
    ]

    stats["max_discount"] = max(discounts)

    stats["avg_discount"] = statistics.mean(
        discounts
    )

    return stats


# ============================================================
# EXPORT
# ============================================================

def export_csv(rows, path):

    with open(
        path,
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as f:

        writer = csv.writer(
            f,
            delimiter=";"
        )

        writer.writerow(
            [
                "ID",
                "Dátum",
                "Időszak",
                "Típus",
                "Márka",
                "Kedvezmény (%)",
            ]
        )

        for r in rows:

            (
                row_id,
                date_str,
                year,
                period,
                promotion_type,
                brand,
                discount,
                _,
            ) = r

            writer.writerow(
                [
                    row_id,
                    date_str,
                    format_period_display(
                        date_str,
                        period,
                        promotion_type,
                        year,
                    ),
                    TYPE_LABELS_HU.get(
                        promotion_type,
                        promotion_type,
                    ),
                    brand,
                    discount,
                ]
            )


def export_excel(rows, path):

    if not OPENPYXL_AVAILABLE:
        raise RuntimeError(
            "openpyxl nincs telepítve. "
            "Telepítés: pip install openpyxl"
        )

    wb = Workbook()

    ws = wb.active
    ws.title = "Promóciók"

    headers = [
        "ID",
        "Dátum",
        "Időszak",
        "Típus",
        "Márka",
        "Kedvezmény (%)",
    ]

    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for r in rows:

        (
            row_id,
            date_str,
            year,
            period,
            promotion_type,
            brand,
            discount,
            _,
        ) = r

        ws.append(
            [
                row_id,
                date_str,
                format_period_display(
                    date_str,
                    period,
                    promotion_type,
                    year,
                ),
                TYPE_LABELS_HU.get(
                    promotion_type,
                    promotion_type,
                ),
                brand,
                discount,
            ]
        )

    for col_cells in ws.columns:

        max_len = max(
            (
                len(str(c.value))
                for c in col_cells
                if c.value is not None
            ),
            default=10,
        )

        ws.column_dimensions[
            col_cells[0].column_letter
        ].width = max_len + 3

    wb.save(path)


# ============================================================
# BACKUP
# ============================================================

def backup_database(db_file):

    BACKUP_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    db_file = Path(db_file)

    if not db_file.exists():
        raise FileNotFoundError(
            "Az adatbázis fájl nem található."
        )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup_path = (
        BACKUP_DIR
        / f"superhobby_backup_{timestamp}.db"
    )

    shutil.copy2(
        db_file,
        backup_path
    )

    return backup_path


# ============================================================
# ADD / EDIT DIALOG
# ============================================================

class PromotionDialog(tk.Toplevel):

    def __init__(
        self,
        parent,
        title,
        existing=None
    ):

        super().__init__(parent)

        self.title(title)
        self.resizable(False, False)
        self.result = None

        self.grab_set()

        existing_date = (
            existing[1]
            if existing
            else datetime.now().strftime(
                "%Y-%m-%d"
            )
        )

        existing_type = (
            existing[4]
            if existing
            else "daily"
        )

        existing_brand = (
            existing[5]
            if existing
            else ""
        )

        existing_discount = (
            existing[6]
            if existing
            else ""
        )

        frame = ttk.Frame(
            self,
            padding=15
        )

        frame.grid(
            row=0,
            column=0
        )

        ttk.Label(
            frame,
            text="Dátum (ÉÉÉÉ-HH-NN):"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            pady=4
        )

        self.date_var = tk.StringVar(
            value=existing_date
        )

        ttk.Entry(
            frame,
            textvariable=self.date_var,
            width=20
        ).grid(
            row=0,
            column=1,
            pady=4
        )

        ttk.Label(
            frame,
            text="Típus:"
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=4
        )

        self.type_var = tk.StringVar(
            value=TYPE_LABELS_HU.get(
                existing_type,
                "Napi"
            )
        )

        ttk.Combobox(
            frame,
            textvariable=self.type_var,
            values=list(
                TYPE_LABELS_HU.values()
            ),
            state="readonly",
            width=18,
        ).grid(
            row=1,
            column=1,
            pady=4
        )

        ttk.Label(
            frame,
            text="Márka:"
        ).grid(
            row=2,
            column=0,
            sticky="w",
            pady=4
        )

        self.brand_var = tk.StringVar(
            value=existing_brand
        )

        ttk.Entry(
            frame,
            textvariable=self.brand_var,
            width=20
        ).grid(
            row=2,
            column=1,
            pady=4
        )

        ttk.Label(
            frame,
            text="Kedvezmény (%):"
        ).grid(
            row=3,
            column=0,
            sticky="w",
            pady=4
        )

        self.discount_var = tk.StringVar(
            value=str(existing_discount)
        )

        ttk.Entry(
            frame,
            textvariable=self.discount_var,
            width=20
        ).grid(
            row=3,
            column=1,
            pady=4
        )

        btn_frame = ttk.Frame(frame)

        btn_frame.grid(
            row=4,
            column=0,
            columnspan=2,
            pady=(12, 0)
        )

        ttk.Button(
            btn_frame,
            text="Mentés",
            command=self._on_save
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            btn_frame,
            text="Mégse",
            command=self.destroy
        ).pack(
            side="left",
            padx=5
        )

        self.bind(
            "<Return>",
            lambda e: self._on_save()
        )

    def _on_save(self):

        try:

            date_obj = datetime.strptime(
                self.date_var.get().strip(),
                "%Y-%m-%d"
            ).date()

        except ValueError:

            messagebox.showerror(
                "Hiba",
                "A dátum formátuma: "
                "ÉÉÉÉ-HH-NN "
                "(pl. 2026-08-13)"
            )

            return

        promotion_type = (
            TYPE_LABELS_EN_TO_INTERNAL.get(
                self.type_var.get()
            )
        )

        if not promotion_type:

            messagebox.showerror(
                "Hiba",
                "Válassz típust."
            )

            return

        brand = self.brand_var.get().strip()

        if not brand:

            messagebox.showerror(
                "Hiba",
                "A márka mező nem lehet üres."
            )

            return

        try:

            discount = int(
                self.discount_var.get().strip()
            )

            if not (0 <= discount <= 100):
                raise ValueError

        except ValueError:

            messagebox.showerror(
                "Hiba",
                "A kedvezménynek 0 és 100 "
                "közötti egész számnak kell lennie."
            )

            return

        self.result = (
            date_obj,
            promotion_type,
            brand,
            discount,
        )

        self.destroy()


# ============================================================
# MAIN DASHBOARD WINDOW
# ============================================================

class Dashboard:

    def __init__(self, root, db_file):

        self.root = root

        self.db_file = Path(db_file)

        self.root.title(
            "Super-Hobby.hu - Promóció Dashboard"
        )

        self.root.geometry("1150x680")

        self.root.minsize(
            950,
            600
        )

        # ----------------------------------------------------
        # DATABASE CONNECTION
        # ----------------------------------------------------

        try:

            self.conn = get_connection(
                self.db_file
            )

        except Exception as e:

            messagebox.showerror(
                "Adatbázis hiba",
                f"Nem sikerült megnyitni az adatbázist:\n\n"
                f"{self.db_file}\n\n"
                f"Hiba:\n{e}"
            )

            self.root.destroy()

            raise

        self.current_rows = []

        # ----------------------------------------------------
        # BUILD UI
        # ----------------------------------------------------

        self._build_database_bar()
        self._build_filter_bar()
        self._build_table()
        self._build_action_bar()
        self._build_stats_panel()

        # Close database properly when window closes.
        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        self.refresh()

    # ========================================================
    # DATABASE BAR
    # ========================================================

    def _build_database_bar(self):

        bar = ttk.Frame(
            self.root,
            padding=(10, 10, 10, 0)
        )

        bar.pack(
            fill="x"
        )

        ttk.Label(
            bar,
            text="Adatbázis:",
            font=("", 9, "bold")
        ).pack(
            side="left"
        )

        self.database_label = ttk.Label(
            bar,
            text=self.db_file.name
        )

        self.database_label.pack(
            side="left",
            padx=(5, 15)
        )

        ttk.Button(
            bar,
            text="Adatbázis kiválasztása",
            command=self.change_database
        ).pack(
            side="left"
        )

        # Show a shortened path on the right.
        self.database_path_label = ttk.Label(
            bar,
            text=str(self.db_file),
            foreground="gray"
        )

        self.database_path_label.pack(
            side="right",
            padx=5
        )

    # ========================================================
    # DATABASE SWITCHING
    # ========================================================

    def change_database(self):

        current_dir = (
            str(self.db_file.parent)
            if self.db_file.exists()
            else str(BASE_DIR)
        )

        selected = choose_database(
            parent=self.root,
            initial_dir=current_dir
        )

        if selected is None:
            return

        # If the user selected the currently active database.
        if selected == self.db_file:
            return

        try:

            new_conn = get_connection(
                selected
            )

        except Exception as e:

            messagebox.showerror(
                "Adatbázis hiba",
                f"Nem sikerült megnyitni a kiválasztott "
                f"adatbázist:\n\n"
                f"{selected}\n\n"
                f"Hiba:\n{e}"
            )

            return

        # Close old connection only after the new database
        # has been successfully opened.
        try:
            self.conn.close()
        except Exception:
            pass

        self.conn = new_conn
        self.db_file = selected

        # Remember selection.
        save_database_path(
            selected
        )

        self.database_label.config(
            text=self.db_file.name
        )

        self.database_path_label.config(
            text=str(self.db_file)
        )

        # Reset filters.
        self.brand_filter.set(
            ALL_LABEL
        )

        self.type_filter.set(
            ALL_LABEL
        )

        self.year_filter.set(
            ALL_LABEL
        )

        self.refresh()

        messagebox.showinfo(
            "Adatbázis kiválasztva",
            f"Az új adatbázis használata aktív:\n\n"
            f"{self.db_file}"
        )

    # ========================================================
    # FILTER BAR
    # ========================================================

    def _build_filter_bar(self):

        bar = ttk.Frame(
            self.root,
            padding=(10, 10, 10, 0)
        )

        bar.pack(
            fill="x"
        )

        ttk.Label(
            bar,
            text="Márka:"
        ).pack(
            side="left"
        )

        self.brand_filter = tk.StringVar(
            value=ALL_LABEL
        )

        self.brand_combo = ttk.Combobox(
            bar,
            textvariable=self.brand_filter,
            state="readonly",
            width=26
        )

        self.brand_combo.pack(
            side="left",
            padx=(4, 15)
        )

        ttk.Label(
            bar,
            text="Típus:"
        ).pack(
            side="left"
        )

        self.type_filter = tk.StringVar(
            value=ALL_LABEL
        )

        self.type_combo = ttk.Combobox(
            bar,
            textvariable=self.type_filter,
            state="readonly",
            width=12,
            values=[
                ALL_LABEL
            ] + list(
                TYPE_LABELS_HU.values()
            ),
        )

        self.type_combo.pack(
            side="left",
            padx=(4, 15)
        )

        ttk.Label(
            bar,
            text="Év:"
        ).pack(
            side="left"
        )

        self.year_filter = tk.StringVar(
            value=ALL_LABEL
        )

        self.year_combo = ttk.Combobox(
            bar,
            textvariable=self.year_filter,
            state="readonly",
            width=10
        )

        self.year_combo.pack(
            side="left",
            padx=(4, 15)
        )

        ttk.Button(
            bar,
            text="Szűrés",
            command=self.refresh
        ).pack(
            side="left",
            padx=4
        )

        ttk.Button(
            bar,
            text="Szűrők törlése",
            command=self._reset_filters
        ).pack(
            side="left"
        )

    # ========================================================
    # TABLE
    # ========================================================

    def _build_table(self):

        table_frame = ttk.Frame(
            self.root,
            padding=10
        )

        table_frame.pack(
            fill="both",
            expand=True
        )

        columns = (
            "id",
            "date_display",
            "type",
            "brand",
            "discount",
        )

        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            selectmode="extended"
        )

        headings = {
            "id": ("ID", 50),
            "date_display": (
                "Dátum / Időszak",
                300
            ),
            "type": (
                "Típus",
                90
            ),
            "brand": (
                "Márka",
                260
            ),
            "discount": (
                "Kedvezmény",
                100
            ),
        }

        for col, (text, width) in headings.items():

            self.tree.heading(
                col,
                text=text
            )

            self.tree.column(
                col,
                width=width,
                anchor="w"
            )

        vsb = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        self.tree.configure(
            yscrollcommand=vsb.set
        )

        self.tree.pack(
            side="left",
            fill="both",
            expand=True
        )

        vsb.pack(
            side="left",
            fill="y"
        )

        self.tree.bind(
            "<Double-1>",
            lambda e: self.edit_selected()
        )

        # ID -> full DB row
        self.row_lookup = {}

    # ========================================================
    # ACTION BAR
    # ========================================================

    def _build_action_bar(self):

        bar = ttk.Frame(
            self.root,
            padding=(10, 0, 10, 10)
        )

        bar.pack(
            fill="x"
        )

        ttk.Button(
            bar,
            text="+ Hozzáadás",
            command=self.add_promotion
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            bar,
            text="Szerkesztés",
            command=self.edit_selected
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            bar,
            text="Törlés (Ctrl/Shift = több)",
            command=self.delete_selected
        ).pack(
            side="left",
            padx=3
        )

        ttk.Separator(
            bar,
            orient="vertical"
        ).pack(
            side="left",
            fill="y",
            padx=10
        )

        ttk.Button(
            bar,
            text="Export CSV",
            command=self.do_export_csv
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            bar,
            text="Export Excel",
            command=self.do_export_excel
        ).pack(
            side="left",
            padx=3
        )

        ttk.Separator(
            bar,
            orient="vertical"
        ).pack(
            side="left",
            fill="y",
            padx=10
        )

        ttk.Button(
            bar,
            text="Adatbázis mentése (backup)",
            command=self.do_backup
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            bar,
            text="Frissítés",
            command=self.refresh
        ).pack(
            side="right",
            padx=3
        )

    # ========================================================
    # STATS PANEL
    # ========================================================

    def _build_stats_panel(self):

        panel = ttk.LabelFrame(
            self.root,
            text="Elemzés (a jelenlegi szűrésre)",
            padding=10
        )

        panel.pack(
            fill="x",
            padx=10,
            pady=(0, 10)
        )

        self.stats_labels = {}

        fields = [
            (
                "count",
                "Promóciók száma:"
            ),
            (
                "by_type",
                "Napi / Heti / Havi:"
            ),
            (
                "avg_interval",
                "Átlagos időköz (nap):"
            ),
            (
                "median_interval",
                "Medián időköz (nap):"
            ),
            (
                "span_text",
                "Lefedett időszak (teljes adatbázis):"
            ),
            (
                "last_date",
                "Utolsó promóció:"
            ),
            (
                "days_since_last",
                "Napok az utolsó óta:"
            ),
            (
                "predicted_next",
                "Következő becsült dátum:"
            ),
            (
                "max_discount",
                "Maximális kedvezmény:"
            ),
            (
                "avg_discount",
                "Átlagos kedvezmény:"
            ),
        ]

        for i, (key, label) in enumerate(fields):

            row, col = divmod(
                i,
                3
            )

            cell = ttk.Frame(panel)

            cell.grid(
                row=row,
                column=col,
                sticky="w",
                padx=15,
                pady=3
            )

            ttk.Label(
                cell,
                text=label,
                font=("", 9, "bold")
            ).pack(
                anchor="w"
            )

            value_lbl = ttk.Label(
                cell,
                text="-"
            )

            value_lbl.pack(
                anchor="w"
            )

            self.stats_labels[key] = value_lbl

    # ========================================================
    # DATA REFRESH
    # ========================================================

    def refresh(self):

        brand_counts = get_brand_counts(
            self.conn
        )

        total_count = sum(
            c for _, c in brand_counts
        )

        # raw brand name / ALL_LABEL -> its current "Name (N)" label
        label_for_raw = {
            ALL_LABEL: f"{ALL_LABEL} ({total_count})"
        }
        for brand, count in brand_counts:
            label_for_raw[brand] = f"{brand} ({count})"

        brands = [
            label_for_raw[ALL_LABEL]
        ] + [
            label_for_raw[brand]
            for brand, _ in brand_counts
        ]

        years = [
            ALL_LABEL
        ] + get_distinct_years(
            self.conn
        )

        self.brand_combo["values"] = brands
        self.year_combo["values"] = years

        # Keep whichever brand was selected, just with a refreshed
        # count - or fall back to "Összes" if it no longer exists.
        raw_selected_brand = strip_count_suffix(
            self.brand_filter.get()
        )
        self.brand_filter.set(
            label_for_raw.get(
                raw_selected_brand,
                label_for_raw[ALL_LABEL]
            )
        )

        if self.year_filter.get() not in years:
            self.year_filter.set(
                ALL_LABEL
            )

        promotion_type = (
            TYPE_LABELS_EN_TO_INTERNAL.get(
                self.type_filter.get()
            )
        )

        rows = fetch_promotions(
            self.conn,
            brand=strip_count_suffix(self.brand_filter.get()),
            promotion_type=promotion_type,
            year=self.year_filter.get(),
        )

        self.current_rows = rows

        self.tree.delete(
            *self.tree.get_children()
        )

        self.row_lookup = {}

        for r in rows:

            (
                row_id,
                date_str,
                year,
                period,
                promotion_type,
                brand,
                discount,
                _,
            ) = r

            display_date = (
                format_period_display(
                    date_str,
                    period,
                    promotion_type,
                    year
                )
            )

            self.tree.insert(
                "",
                "end",
                iid=str(row_id),
                values=(
                    row_id,
                    display_date,
                    TYPE_LABELS_HU.get(
                        promotion_type,
                        promotion_type,
                    ),
                    brand,
                    f"{discount}%",
                ),
            )

            self.row_lookup[
                str(row_id)
            ] = r

        self._update_stats(
            rows
        )

    def _reset_filters(self):

        self.brand_filter.set(
            ALL_LABEL
        )

        self.type_filter.set(
            ALL_LABEL
        )

        self.year_filter.set(
            ALL_LABEL
        )

        self.refresh()

    # ========================================================
    # STATISTICS
    # ========================================================

    def _update_stats(self, rows):

        stats = compute_stats(
            rows
        )

        self.stats_labels[
            "count"
        ].config(
            text=str(
                stats["count"]
            )
        )

        by_type = stats["by_type"]

        self.stats_labels[
            "by_type"
        ].config(
            text=(
                f"{by_type.get('daily', 0)} / "
                f"{by_type.get('weekly', 0)} / "
                f"{by_type.get('monthly', 0)}"
            )
        )

        self.stats_labels[
            "avg_interval"
        ].config(
            text=(
                f"{stats['avg_interval']:.1f}"
                if stats["avg_interval"] is not None
                else "-"
            )
        )

        self.stats_labels[
            "median_interval"
        ].config(
            text=(
                f"{stats['median_interval']:.1f}"
                if stats["median_interval"] is not None
                else "-"
            )
        )

        self.stats_labels[
            "span_text"
        ].config(
            text=(
                get_database_span_text(self.conn)
                or "-"
            )
        )

        self.stats_labels[
            "last_date"
        ].config(
            text=(
                format_long_date(
                    stats["last_date"]
                )
                if stats["last_date"]
                else "-"
            )
        )

        self.stats_labels[
            "days_since_last"
        ].config(
            text=(
                str(
                    stats["days_since_last"]
                )
                if stats["days_since_last"] is not None
                else "-"
            )
        )

        self.stats_labels[
            "predicted_next"
        ].config(
            text=(
                format_long_date(
                    stats["predicted_next"]
                )
                if stats["predicted_next"]
                else "-"
            )
        )

        self.stats_labels[
            "max_discount"
        ].config(
            text=(
                f"{stats['max_discount']}%"
                if stats["max_discount"] is not None
                else "-"
            )
        )

        self.stats_labels[
            "avg_discount"
        ].config(
            text=(
                f"{stats['avg_discount']:.1f}%"
                if stats["avg_discount"] is not None
                else "-"
            )
        )

    # ========================================================
    # CRUD
    # ========================================================

    def _get_selected_rows(self):

        selection = self.tree.selection()

        if not selection:

            messagebox.showinfo(
                "Nincs kijelölés",
                "Előbb válassz ki legalább egy sort "
                "a táblázatban."
            )

            return []

        return [
            self.row_lookup[iid]
            for iid in selection
            if iid in self.row_lookup
        ]

    def add_promotion(self):

        dialog = PromotionDialog(
            self.root,
            "Promóció hozzáadása"
        )

        self.root.wait_window(
            dialog
        )

        if dialog.result:

            (
                date_obj,
                promotion_type,
                brand,
                discount,
            ) = dialog.result

            try:

                insert_promotion(
                    self.conn,
                    date_obj,
                    promotion_type,
                    brand,
                    discount
                )

            except sqlite3.IntegrityError:

                messagebox.showerror(
                    "Hiba",
                    "Ez a promóció "
                    "(év + időszak + típus + márka) "
                    "már létezik."
                )

                return

            self.refresh()

    def edit_selected(self):

        rows = self._get_selected_rows()

        if not rows:
            return

        if len(rows) > 1:

            messagebox.showinfo(
                "Több sor kijelölve",
                "Szerkesztéshez csak egy sort jelölj ki."
            )

            return

        row = rows[0]

        dialog = PromotionDialog(
            self.root,
            "Promóció szerkesztése",
            existing=row
        )

        self.root.wait_window(
            dialog
        )

        if dialog.result:

            (
                date_obj,
                promotion_type,
                brand,
                discount,
            ) = dialog.result

            try:

                update_promotion(
                    self.conn,
                    row[0],
                    date_obj,
                    promotion_type,
                    brand,
                    discount
                )

            except sqlite3.IntegrityError:

                messagebox.showerror(
                    "Hiba",
                    "Ez a promóció "
                    "(év + időszak + típus + márka) "
                    "már létezik."
                )

                return

            self.refresh()

    def delete_selected(self):

        rows = self._get_selected_rows()

        if not rows:
            return

        if len(rows) == 1:

            row = rows[0]

            message = (
                "Biztosan törlöd ezt a bejegyzést?\n\n"
                f"{row[5]} - {row[6]}% ({row[1]})"
            )

        else:

            preview = "\n".join(
                f"• {r[5]} - {r[6]}% ({r[1]})"
                for r in rows[:10]
            )

            more = (
                f"\n... és még "
                f"{len(rows) - 10} bejegyzés"
                if len(rows) > 10
                else ""
            )

            message = (
                f"Biztosan törlöd ezt a "
                f"{len(rows)} bejegyzést?\n\n"
                f"{preview}{more}"
            )

        confirm = messagebox.askyesno(
            "Megerősítés",
            message
        )

        if confirm:

            for row in rows:

                delete_promotion(
                    self.conn,
                    row[0]
                )

            self.refresh()

    # ========================================================
    # EXPORT / BACKUP
    # ========================================================

    def do_export_csv(self):

        if not self.current_rows:

            messagebox.showinfo(
                "Nincs adat",
                "Nincs exportálható adat "
                "a jelenlegi szűréssel."
            )

            return

        EXPORT_DIR.mkdir(
            parents=True,
            exist_ok=True
        )

        default_name = (
            f"promociok_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            f".csv"
        )

        path = filedialog.asksaveasfilename(
            initialdir=EXPORT_DIR,
            initialfile=default_name,
            defaultextension=".csv",
            filetypes=[
                (
                    "CSV fájl",
                    "*.csv"
                )
            ],
        )

        if not path:
            return

        export_csv(
            self.current_rows,
            path
        )

        messagebox.showinfo(
            "Kész",
            f"Exportálva:\n{path}"
        )

    def do_export_excel(self):

        if not self.current_rows:

            messagebox.showinfo(
                "Nincs adat",
                "Nincs exportálható adat "
                "a jelenlegi szűréssel."
            )

            return

        if not OPENPYXL_AVAILABLE:

            messagebox.showerror(
                "Hiányzó függőség",
                "Az Excel exporthoz telepítsd:\n\n"
                "pip install openpyxl"
            )

            return

        EXPORT_DIR.mkdir(
            parents=True,
            exist_ok=True
        )

        default_name = (
            f"promociok_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            f".xlsx"
        )

        path = filedialog.asksaveasfilename(
            initialdir=EXPORT_DIR,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[
                (
                    "Excel fájl",
                    "*.xlsx"
                )
            ],
        )

        if not path:
            return

        export_excel(
            self.current_rows,
            path
        )

        messagebox.showinfo(
            "Kész",
            f"Exportálva:\n{path}"
        )

    def do_backup(self):

        try:

            path = backup_database(
                self.db_file
            )

        except FileNotFoundError as e:

            messagebox.showerror(
                "Hiba",
                str(e)
            )

            return

        except Exception as e:

            messagebox.showerror(
                "Backup hiba",
                f"Nem sikerült elkészíteni a mentést:\n\n"
                f"{e}"
            )

            return

        messagebox.showinfo(
            "Mentés kész",
            f"Adatbázis mentve ide:\n{path}"
        )

    # ========================================================
    # CLOSE
    # ========================================================

    def on_close(self):

        try:
            self.conn.close()
        except Exception:
            pass

        self.root.destroy()


# ============================================================
# ENTRY POINT
# ============================================================

def main():

    # A single Tk root is used for the entire application lifetime -
    # both the initial database-picker dialog and the dashboard
    # itself. Creating a second Tk() instance later (e.g. destroying
    # this one and building a new Tk() for the dashboard) confuses
    # ttk's theme engine and causes a
    # "can't invoke event command: application has been destroyed"
    # error on exit.
    root = tk.Tk()

    root.withdraw()

    db_path = get_database_path(
        parent=root
    )

    if db_path is None:

        messagebox.showinfo(
            "Kilépés",
            "Nem választottál adatbázist.",
            parent=root
        )

        root.destroy()

        return

    root.deiconify()

    # Start actual dashboard on the same root window.
    try:
        Dashboard(
            root,
            db_path
        )
    except Exception:
        # Dashboard.__init__ already showed an error dialog and
        # destroyed the root window in this case.
        return

    root.mainloop()


if __name__ == "__main__":
    main()
