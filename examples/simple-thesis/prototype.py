"""Reproducible reservation prototype and the functional cases in Chapter 5."""
from __future__ import annotations

from datetime import date
import json
from pathlib import Path
import platform
import sqlite3

ROOT = Path(__file__).resolve().parent
DAY = "2026-10-09"


def database() -> sqlite3.Connection:
    db = sqlite3.connect(":memory:", isolation_level=None)
    db.execute("PRAGMA foreign_keys = ON")
    db.execute("PRAGMA busy_timeout = 5000")
    db.executescript("""
        CREATE TABLE users (id INTEGER PRIMARY KEY);
        CREATE TABLE seats (id INTEGER PRIMARY KEY);
        CREATE TABLE bookings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id),
            seat_id INTEGER NOT NULL REFERENCES seats(id),
            day TEXT NOT NULL,
            start INTEGER NOT NULL,
            end INTEGER NOT NULL CHECK (end > start),
            status TEXT NOT NULL DEFAULT 'active'
                CHECK (status IN ('active', 'cancelled'))
        );
        INSERT INTO users VALUES (1), (2);
        INSERT INTO seats VALUES (1), (2), (3);
    """)
    return db


def reserve(db, user, seat, day, start, end):
    try:
        if date.fromisoformat(day).isoformat() != day:
            return "invalid", None
    except (ValueError, TypeError):
        return "invalid", None
    if type(start) is not int or type(end) is not int:
        return "invalid", None
    if not 480 <= start < end <= 1320:
        return "invalid", None

    # reservebegin
    db.execute("BEGIN IMMEDIATE")
    try:
        if not db.execute("SELECT 1 FROM users WHERE id=?", (user,)).fetchone():
            raise ValueError("unknown-user")
        if not db.execute("SELECT 1 FROM seats WHERE id=?", (seat,)).fetchone():
            raise ValueError("unknown-seat")
        overlap = "day=? AND status='active' AND start<? AND end>?"
        query = "SELECT 1 FROM bookings WHERE "
        if db.execute(query + "seat_id=? AND " + overlap,
                      (seat, day, end, start)).fetchone():
            raise ValueError("seat-conflict")
        if db.execute(query + "user_id=? AND " + overlap,
                      (user, day, end, start)).fetchone():
            raise ValueError("user-conflict")
        cursor = db.execute(
            "INSERT INTO bookings(user_id,seat_id,day,start,end) "
            "VALUES (?,?,?,?,?)", (user, seat, day, start, end))
        db.execute("COMMIT")
        return "created", cursor.lastrowid
    except ValueError as error:
        db.execute("ROLLBACK")
        return str(error), None
    except BaseException:
        db.execute("ROLLBACK")
        raise
    # reserveend


def cancel(db, user, booking):
    cursor = db.execute(
        "UPDATE bookings SET status='cancelled' "
        "WHERE id=? AND user_id=? AND status='active'", (booking, user))
    return "cancelled" if cursor.rowcount == 1 else "not-owner-or-inactive"


def statistics(db, day):
    return db.execute(
        "SELECT COUNT(*), COALESCE(SUM(end-start),0) FROM bookings "
        "WHERE day=? AND status='active'", (day,)).fetchone()


def main():
    cases = []

    def check(identifier, name, actual, expected):
        passed = actual == expected
        cases.append(dict(id=identifier, name=name, actual=actual,
                          expected=expected, passed=passed))
        print(f"{identifier}: {'PASS' if passed else 'FAIL'} - {name}")
        if not passed:
            raise AssertionError(f"{identifier}: {actual!r} != {expected!r}")

    db = database()
    try:
        result, first = reserve(db, 1, 1, DAY, 540, 600)
        check("T01", "valid reservation", result, "created")
        check("T02", "seat overlap", reserve(db, 2, 1, DAY, 570, 630)[0], "seat-conflict")
        check("T03", "adjacent periods", reserve(db, 2, 1, DAY, 600, 660)[0], "created")
        check("T04", "user overlap", reserve(db, 1, 2, DAY, 570, 630)[0], "user-conflict")
        check("T05", "reversed interval", reserve(db, 2, 2, DAY, 720, 660)[0], "invalid")
        check("T06", "unknown seat", reserve(db, 2, 99, DAY, 660, 720)[0], "unknown-seat")
        check("T07", "another user's cancellation", cancel(db, 2, first), "not-owner-or-inactive")
        check("T08", "owner's cancellation", cancel(db, 1, first), "cancelled")
        check("T09", "released period", reserve(db, 2, 1, DAY, 540, 600)[0], "created")
        check("T10", "repeated cancellation", cancel(db, 1, first), "not-owner-or-inactive")
        check("T11", "different day", reserve(db, 1, 1, "2026-10-10", 540, 600)[0], "created")
        count, minutes = statistics(db, DAY)
        check("T12", "daily statistics", (count, minutes), (2, 120))
        report = dict(python=platform.python_version(), sqlite=sqlite3.sqlite_version,
                      cases=cases, passed=len(cases), total=len(cases), day=DAY,
                      active_bookings=count, booked_minutes=minutes,
                      available_seats=3, opening_minutes=840,
                      utilization_percent=round(minutes / (3 * 840) * 100, 2))
    finally:
        db.close()

    assets = ROOT / "assets"
    assets.mkdir(exist_ok=True)
    (assets / "test-results.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    commands = {
        "SampleTestCount": report["total"],
        "SamplePassedCount": report["passed"],
        "SamplePythonVersion": report["python"],
        "SampleSQLiteVersion": report["sqlite"],
        "SampleUtilization": f'{report["utilization_percent"]:.2f}',
    }
    (assets / "test-summary.tex").write_text(
        "".join(f"\\newcommand{{\\{name}}}{{{value}}}\n" for name, value in commands.items()),
        encoding="utf-8")
    source = Path(__file__).read_text(encoding="utf-8")
    core = source.split("    # reservebegin\n", 1)[1].split("    # reserveend\n", 1)[0]
    header = next(line for line in source.splitlines() if line.startswith("def reserve("))
    (assets / "reservation-core.py").write_text(header + "\n" + core, encoding="utf-8")
    print(f"{report['passed']}/{report['total']} passed; utilization {report['utilization_percent']:.2f}%")


if __name__ == "__main__":
    main()
