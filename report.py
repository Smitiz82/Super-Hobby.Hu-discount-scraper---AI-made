import sqlite3
import sys
from pathlib import Path
from statistics import mean, median
from datetime import datetime


BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "data" / "superhobby.db"


def get_connection():
    return sqlite3.connect(DB_FILE)


def brand_history(brand):

    conn = get_connection()

    rows = conn.execute("""
        SELECT date, promotion_type, discount
        FROM promotions
        WHERE LOWER(brand) = LOWER(?)
        ORDER BY date
    """, (brand,)).fetchall()

    conn.close()

    return rows


def date_difference(date1, date2):

    d1 = datetime.strptime(date1, "%Y-%m-%d")
    d2 = datetime.strptime(date2, "%Y-%m-%d")

    return (d2 - d1).days


def report(brand):

    rows = brand_history(brand)

    print()
    print("=" * 60)
    print(f"{brand.upper()} - SUPER-HOBBY DISCOUNT HISTORY")
    print("=" * 60)

    if not rows:
        print("No records found.")
        return

    print()

    for date, promotion_type, discount in rows:

        print(
            f"{date}   "
            f"{promotion_type:8} "
            f"- {discount}%"
        )

    # Statistics

    dates = [row[0] for row in rows]

    intervals = []

    for i in range(1, len(dates)):
        intervals.append(
            date_difference(
                dates[i - 1],
                dates[i]
            )
        )

    print()
    print("-" * 60)

    print(f"Number of recorded promotions: {len(rows)}")

    daily = sum(
        1 for row in rows
        if row[1] == "daily"
    )

    weekly = sum(
        1 for row in rows
        if row[1] == "weekly"
    )

    monthly = sum(
        1 for row in rows
        if row[1] == "monthly"
    )

    print(f"Daily:   {daily}")
    print(f"Weekly:  {weekly}")
    print(f"Monthly: {monthly}")

    discounts = [
        row[2]
        for row in rows
    ]

    print(
        f"Average discount: "
        f"{mean(discounts):.1f}%"
    )

    if intervals:

        print(
            f"Average interval: "
            f"{mean(intervals):.1f} days"
        )

        print(
            f"Median interval: "
            f"{median(intervals):.1f} days"
        )

        print(
            f"Shortest interval: "
            f"{min(intervals)} days"
        )

        print(
            f"Longest interval: "
            f"{max(intervals)} days"
        )

    print()
    print(f"Last promotion: {dates[-1]}")

    if len(dates) >= 2:

        previous = dates[-2]

        print(
            f"Previous promotion: {previous}"
        )

        print(
            f"Days since previous: "
            f"{date_difference(previous, dates[-1])}"
        )


if __name__ == "__main__":

    if len(sys.argv) < 2:

        print(
            "Usage:\n"
            "  python report.py Vallejo"
        )

        sys.exit(1)

    brand = " ".join(sys.argv[1:])

    report(brand)