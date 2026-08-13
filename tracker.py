import json
import re
import sqlite3
import sys
from pathlib import Path
from datetime import datetime
 
# Windows' cmd.exe often uses a non-UTF-8 console codepage, which can
# make print() crash on accented Hungarian characters (é, á, ó, ...).
# Force UTF-8 output so the script never dies on a print() call.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
 
import requests
from bs4 import BeautifulSoup
 
import urllib3
 
urllib3.disable_warnings(
    urllib3.exceptions.InsecureRequestWarning
)
 
 
# ============================================================
# CONFIGURATION
# ============================================================
 
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_FILE = DATA_DIR / "superhobby.db"
CONFIG_FILE = BASE_DIR / "config.json"
 
URL = "https://www.super-hobby.hu/"
 
 
# ============================================================
# DATABASE
# ============================================================
 
def get_table_columns(conn):
    if not conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='promotions'"
    ).fetchone():
        return []
    return [row[1] for row in conn.execute("PRAGMA table_info(promotions)").fetchall()]
 
 
def get_unique_index_column_sets(conn):
    """All sets of columns covered by a UNIQUE index/constraint on
    'promotions' (this reads SQLite's own schema introspection instead
    of parsing the CREATE TABLE text, which is more reliable)."""
 
    result = []
 
    for idx in conn.execute("PRAGMA index_list(promotions)").fetchall():
        # idx columns: (seq, name, unique, origin, partial)
        is_unique = idx[2]
        index_name = idx[1]
 
        if not is_unique:
            continue
 
        cols = [
            info[2]
            for info in conn.execute(f"PRAGMA index_info('{index_name}')").fetchall()
        ]
        result.append(set(cols))
 
    return result
 
 
def table_needs_migration(conn):
    """True if the promotions table exists but is missing the 'year'
    column, or its UNIQUE constraint is not yet (year, period,
    promotion_type, brand). That old constraint - (date, promotion_type,
    brand) - let a weekly/monthly promotion be re-inserted every day
    the scraper ran, since 'date' changed daily even though the period
    (KW.. / YYYY-MM) stayed the same."""
 
    columns = get_table_columns(conn)
 
    if not columns:
        return False  # table doesn't exist yet - nothing to migrate
 
    if "year" not in columns:
        return True
 
    target = {"year", "period", "promotion_type", "brand"}
    unique_sets = get_unique_index_column_sets(conn)
 
    return not any(cols == target for cols in unique_sets)
 
 
def migrate_database(conn):
    if not table_needs_migration(conn):
        return
 
    print("Adatbázis frissítése (migráció)...")
 
    existing_cols = get_table_columns(conn)
 
    if "year" not in existing_cols:
        conn.execute("ALTER TABLE promotions ADD COLUMN year INTEGER")
        conn.commit()
 
    # Backfill any missing year values from the date column. This runs
    # unconditionally (not just right after adding the column) so it
    # also repairs a table that already had a 'year' column with NULLs
    # left in it from an earlier, incomplete migration.
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
 
    # Keep only the earliest row of each (year, period, type, brand)
    # group - this removes the duplicate monthly/weekly rows that the
    # old schema allowed to accumulate.
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
 
    row_count = conn.execute("SELECT COUNT(*) FROM promotions").fetchone()[0]
    null_years = conn.execute("SELECT COUNT(*) FROM promotions WHERE year IS NULL").fetchone()[0]
 
    print(f"Migráció kész - {row_count} bejegyzés, ebből {null_years} hiányzó évvel.")
 
 
def init_database():
 
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
 
    # Belt-and-suspenders: guarantee there is never a NULL year in the
    # table, regardless of how a row got in (manual SQL, older code,
    # a future edit path, etc.)
    conn.execute("""
        UPDATE promotions
        SET year = CAST(substr(date, 1, 4) AS INTEGER)
        WHERE year IS NULL
    """)
    conn.commit()
 
    return conn
 
 
# ============================================================
# DOWNLOAD WEBSITE
# ============================================================
 
def download_page():
 
    headers = {
        "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/150.0 Safari/537.36",
 
        "Accept-Language":
            "hu-HU,hu;q=0.9,en;q=0.8"
    }
 
    response = requests.get(
    URL,
    headers=headers,
    timeout=30,
    verify=False
    )
 
    response.raise_for_status()
 
    return response.text
 
 
# ============================================================
# TEXT CLEANING
# ============================================================
 
def clean_text(text):
 
    return re.sub(
        r"\s+",
        " ",
        text
    ).strip()
 
 
# ============================================================
# FIND PROMOTIONS
# ============================================================
 
def extract_promotions(html):
 
    soup = BeautifulSoup(
        html,
        "lxml"
    )
 
    # Remove things we don't need
    for element in soup(
        ["script", "style", "noscript"]
    ):
        element.decompose()
 
    lines = [
        clean_text(line)
        for line in soup.get_text("\n").splitlines()
        if clean_text(line)
    ]
 
    promotions = []
 
    current_type = None
 
    section_patterns = {
 
        "daily": [
            "Napi kedvezmény",
            "Promotion of the day",
            "Oferta dnia"
        ],
 
        "weekly": [
            "Heti kedvezmény",
            "Promotion of the week",
            "Oferta tygodnia"
        ],
 
        "monthly": [
            "Havi kedvezmény",
            "Promotion of the month",
            "Oferta miesiąca"
        ]
    }
 
 
    for line in lines:
 
        # ----------------------------------------------------
        # Detect section
        # ----------------------------------------------------
 
        found_section = None
 
        for section, patterns in section_patterns.items():
 
            for pattern in patterns:
 
                if pattern.lower() in line.lower():
 
                    found_section = section
                    break
 
            if found_section:
                break
 
 
        if found_section:
 
            current_type = found_section
 
            continue
 
 
        if current_type is None:
            continue
 
 
        # ----------------------------------------------------
        # Detect:
        #
        # MiniArt - 12%
        # Sword - 12%
        # ----------------------------------------------------
 
        match = re.match(
            r"^(.+?)\s*[-\u2013]\s*(\d+)\s*%\s*$",
            line
        )
 
 
        if not match:
            continue
 
 
        brand = clean_text(
            match.group(1)
        )
 
        discount = int(
            match.group(2)
        )
 
 
        # ----------------------------------------------------
        # Only accept expected discount levels
        # ----------------------------------------------------
 
        expected_discount = {
 
            "daily": 12,
            "weekly": 8,
            "monthly": 5
 
        }[current_type]
 
 
        if discount != expected_discount:
            continue
 
 
        # ----------------------------------------------------
        # Add promotion
        # ----------------------------------------------------
 
        promotions.append({
 
            "type": current_type,
 
            "brand": brand,
 
            "discount": discount
 
        })
 
 
        # ----------------------------------------------------
        # Super-Hobby has 3 brands per section
        # ----------------------------------------------------
 
        count = sum(
            1
            for p in promotions
            if p["type"] == current_type
        )
 
 
        if count >= 3:
 
            current_type = None
 
 
    return promotions
 
 
# ============================================================
# ISO WEEK / YEAR
# ============================================================
 
def get_iso_week(date):
 
    return date.isocalendar().week
 
 
def compute_year(date, promotion_type):
    """The year an entry belongs to. Weekly promotions use the ISO
    calendar year (so e.g. a week that starts in late December but
    belongs to ISO week 1 is filed under the correct following year);
    daily/monthly just use the plain calendar year."""
 
    if promotion_type == "weekly":
        return date.isocalendar().year
 
    return date.year
 
 
# ============================================================
# SAVE PROMOTIONS
# ============================================================
 
def save_promotions(
    conn,
    promotions,
    date
):
 
    date_string = date.strftime(
        "%Y-%m-%d"
    )
 
    created_at = datetime.now().isoformat(
        timespec="seconds"
    )
 
 
    for promotion in promotions:
 
        promotion_type = promotion["type"]
 
        brand = promotion["brand"]
 
        discount = promotion["discount"]
 
        year = compute_year(date, promotion_type)
 
 
        # ----------------------------------------------------
        # Period displayed to the user
        # ----------------------------------------------------
 
        if promotion_type == "daily":
 
            period = date_string
 
 
        elif promotion_type == "weekly":
 
            period = f"KW{get_iso_week(date):02d}"
 
 
        elif promotion_type == "monthly":
 
            # Hungarian month name is generated when exporting.
            # Internally we store YYYY-MM.
            period = date.strftime("%Y-%m")
 
 
        else:
 
            period = date_string
 
 
        # ----------------------------------------------------
        # INSERT OR IGNORE now keys off (year, period, type, brand)
        # instead of the raw date - so a weekly/monthly promotion
        # that is still current the next time the scraper runs
        # (same KW.. or same YYYY-MM) is not inserted again.
        # ----------------------------------------------------
 
        cursor = conn.execute("""
 
            INSERT OR IGNORE INTO promotions
 
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
 
        """, (
 
            date_string,
            year,
            period,
            promotion_type,
            brand,
            discount,
            created_at
 
        ))
 
        promotion["year"] = year
        promotion["_inserted"] = cursor.rowcount > 0
 
 
    conn.commit()
 
 
# ============================================================
# SANITY CHECK
# ============================================================
 
def validate_promotions(promotions):
 
    daily = [
        p for p in promotions
        if p["type"] == "daily"
    ]
 
    weekly = [
        p for p in promotions
        if p["type"] == "weekly"
    ]
 
    monthly = [
        p for p in promotions
        if p["type"] == "monthly"
    ]
 
 
    if len(daily) != 3:
 
        raise RuntimeError(
            f"Expected 3 daily promotions, "
            f"found {len(daily)}"
        )
 
 
    if len(weekly) != 3:
 
        raise RuntimeError(
            f"Expected 3 weekly promotions, "
            f"found {len(weekly)}"
        )
 
 
    if len(monthly) != 3:
 
        raise RuntimeError(
            f"Expected 3 monthly promotions, "
            f"found {len(monthly)}"
        )
 
 
# ============================================================
# DISPLAY
# ============================================================
 
def print_promotions(
    promotions,
    date
):
 
    print()
 
    print(
        date.strftime(
            "%Y-%m-%d (%A)"
        )
    )
 
    print(
        "-" * 40
    )
 
 
    for promotion in promotions:
 
        year = promotion.get("year", compute_year(date, promotion["type"]))
        status = "új" if promotion.get("_inserted", True) else "már létezett"
 
        print(
            f"{promotion['type']:8} | "
            f"{promotion['brand']} - "
            f"{promotion['discount']}%  "
            f"[év={year}, {status}]"
        )
 
 
# ============================================================
# MAIN
# ============================================================
 
def main():
 
    print(
        "Super-Hobby.hu promotion tracker"
    )
 
    print(
        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    )
 
    print(
        f"Adatbázis: {DB_FILE}"
    )
 
    print()
 
 
    conn = None
 
 
    try:
 
        # ----------------------------------------------------
        # Database (creates the table and/or migrates it)
        # ----------------------------------------------------
 
        conn = init_database()
 
 
        # ----------------------------------------------------
        # Today's date
        # ----------------------------------------------------
 
        today = datetime.now().date()
 
 
        # ----------------------------------------------------
        # Download
        # ----------------------------------------------------
 
        print(
            "Downloading Super-Hobby..."
        )
 
        html = download_page()
 
 
        # ----------------------------------------------------
        # Extract
        # ----------------------------------------------------
 
        print(
            "Reading promotions..."
        )
 
        promotions = extract_promotions(
            html
        )
 
 
        # ----------------------------------------------------
        # Validate
        # ----------------------------------------------------
 
        validate_promotions(
            promotions
        )
 
 
        # ----------------------------------------------------
        # Save
        # ----------------------------------------------------
 
        save_promotions(
            conn,
            promotions,
            today
        )
 
 
        # ----------------------------------------------------
        # Display
        # ----------------------------------------------------
 
        print(
            "Successfully saved."
        )
 
        print_promotions(
            promotions,
            today
        )
 
 
    except Exception as e:
 
        print()
 
        print(
            "ERROR:"
        )
 
        print(
            str(e)
        )
 
 
    finally:
 
        if conn is not None:
            conn.close()
 
 
# ============================================================
 
if __name__ == "__main__":
 
    main()