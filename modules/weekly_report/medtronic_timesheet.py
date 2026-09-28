"""Builds/maintains the single, continuously-updated Medtronic daily
timesheet workbook (output/Medtronic_Time_Sheet.xlsx) from a flat JSON store
of daily entries (data/medtronic_timesheet_entries.json). Both are persisted
in Vercel Blob (see blob_store.py) rather than local disk - Vercel's
deployment filesystem is read-only, and even /tmp is wiped between deploys
and isn't shared across serverless instances, so anything written there
could vanish or be invisible to the very next request. Every save rebuilds
the whole workbook from the JSON store rather than surgically editing the
previous .xlsx - far less error-prone than shifting merged cell ranges by
hand, and the result is identical either way since the store is the single
source of truth.

Layout mirrors the "check" reference sheet in
template/Medtronic_Time_Sheet.xlsx (that sheet is only ever read, never
written to): a title row, then one block per week - a "Reporting Date
(...)" banner, a column header row, one merged Date/Total-Hrs row-group per
day (one row per person who logged hours that day), and a closing "Total
hours" row. Weeks are stacked newest-first; days within a week are
chronological, matching the Medtronic weekly status docx's "newest week on
top" convention."""

import io
import json
import threading
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from . import blob_store
from .llm_consolidate import consolidate_notes

STORE_PATHNAME = "data/medtronic_timesheet_entries.json"
WORKBOOK_FILENAME = "Medtronic_Time_Sheet.xlsx"
XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SHEET_NAME = "Medtronic Time Sheet"

_lock = threading.Lock()

TITLE_FILL = PatternFill("solid", fgColor="FF1F5D91")
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFFFF")
WEEK_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="FFD9E6F2")
HEADER_FONT = Font(name="Calibri", size=11, bold=True)
DATA_FONT = Font(name="Calibri", size=11)
TOTAL_FONT = Font(name="Calibri", size=11, bold=True)

_THIN = Side(style="thin")
BOX = Border(top=_THIN, bottom=_THIN, left=_THIN, right=_THIN)

CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(vertical="center")
LEFT_WRAP = Alignment(vertical="center", wrap_text=True)

COLUMN_WIDTHS = {"A": 16, "B": 27, "C": 92, "D": 9, "E": 14}
HEADERS = ["Date", "Responsible Person", "Description", "Hrs", "Total Hrs"]


def _load_entries() -> list:
    raw = blob_store.get(STORE_PATHNAME)
    return json.loads(raw) if raw else []


def _save_entries(entries: list) -> None:
    blob_store.put(
        STORE_PATHNAME,
        json.dumps(entries, indent=2, ensure_ascii=False).encode("utf-8"),
        content_type="application/json",
    )


def _week_bounds(date_iso: str):
    """Mon-Fri work week containing date_iso, matching the reference
    sheet's 09/07/2026 (Mon) - 09/11/2026 (Fri) example."""
    d = datetime.strptime(date_iso, "%Y-%m-%d").date()
    monday = d - timedelta(days=d.weekday())
    friday = monday + timedelta(days=4)
    return monday, friday


def _fmt(d) -> str:
    return d.strftime("%m/%d/%Y")


def _group_by_week(entries: list) -> dict:
    """{week_start_iso: {"end": week_end_iso, "dates": {date_iso: [entry, ...]}}}"""
    weeks = {}
    for entry in entries:
        monday, friday = _week_bounds(entry["date"])
        week = weeks.setdefault(monday.isoformat(), {"end": friday.isoformat(), "dates": {}})
        week["dates"].setdefault(entry["date"], []).append(entry)
    return weeks


def _write_banner(ws, row: int, text: str, font):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = font
    cell.fill = TITLE_FILL
    cell.alignment = CENTER
    ws.row_dimensions[row].height = 18 if row == 1 else 15.6


def _write_header(ws, row: int):
    for col, text in enumerate(HEADERS, start=1):
        cell = ws.cell(row=row, column=col, value=text)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BOX
        cell.alignment = CENTER_WRAP


def _style_row(ws, row: int):
    for col in range(1, 6):
        cell = ws.cell(row=row, column=col)
        cell.font = DATA_FONT
        cell.border = BOX


def _write_day_groups(ws, start_row: int, dates: dict) -> tuple:
    """Writes one merged Date/Total-Hrs row-group per date in dates (a
    {date_iso: [entry, ...]} dict), in chronological order, starting at
    start_row. Returns (next_free_row, total_hrs_written)."""
    row = start_row
    total = 0.0
    for date_iso in sorted(dates.keys()):
        day_entries = dates[date_iso]
        day_start_row = row
        day_total = sum(float(e["hrs"]) for e in day_entries)
        total += day_total

        for entry in day_entries:
            ws.cell(row=row, column=2, value=entry["person"]).alignment = LEFT
            ws.cell(row=row, column=3, value=entry["description"]).alignment = LEFT_WRAP
            ws.cell(row=row, column=4, value=float(entry["hrs"])).alignment = CENTER
            _style_row(ws, row)
            row += 1

        day_end_row = row - 1
        day_date = datetime.strptime(date_iso, "%Y-%m-%d").date()
        date_cell = ws.cell(row=day_start_row, column=1, value=_fmt(day_date))
        date_cell.alignment = CENTER
        total_cell = ws.cell(row=day_start_row, column=5, value=day_total)
        total_cell.alignment = CENTER
        total_cell.font = TOTAL_FONT
        if day_end_row > day_start_row:
            ws.merge_cells(start_row=day_start_row, start_column=1, end_row=day_end_row, end_column=1)
            ws.merge_cells(start_row=day_start_row, start_column=5, end_row=day_end_row, end_column=5)

    return row, total


def _write_total_row(ws, row: int, total: float, label: str = "Total hours") -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    for col in range(1, 5):
        cell = ws.cell(row=row, column=col)
        cell.border = BOX
        cell.font = TOTAL_FONT
        cell.alignment = CENTER
    ws.cell(row=row, column=1, value=label)
    total_value = ws.cell(row=row, column=5, value=total)
    total_value.font = TOTAL_FONT
    total_value.alignment = CENTER
    total_value.border = BOX


def _new_workbook(ws_title: str):
    wb = Workbook()
    ws = wb.active
    ws.title = ws_title
    for col, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width
    return wb, ws


def _rebuild_workbook(entries: list) -> None:
    """Rebuilds the one live, continuously-updated workbook: every week
    gets its own 'Reporting Date (...)' banner + header + total, stacked
    newest week first."""
    wb, ws = _new_workbook(SHEET_NAME)

    _write_banner(ws, 1, "Medtronic Time Sheet", TITLE_FONT)
    row = 2

    weeks = _group_by_week(entries)
    for week_start in sorted(weeks.keys(), reverse=True):
        week = weeks[week_start]
        start_d = datetime.strptime(week_start, "%Y-%m-%d").date()
        end_d = datetime.strptime(week["end"], "%Y-%m-%d").date()

        _write_banner(ws, row, f"Reporting Date ({_fmt(start_d)} - {_fmt(end_d)})", WEEK_FONT)
        row += 1
        _write_header(ws, row)
        row += 1

        row, week_total = _write_day_groups(ws, row, week["dates"])
        _write_total_row(ws, row, week_total)
        row += 1

    buf = io.BytesIO()
    wb.save(buf)
    blob_store.put(f"output/{WORKBOOK_FILENAME}", buf.getvalue(), content_type=XLSX_CONTENT_TYPE)


def _parse_range(start_iso: str, end_iso: str):
    start_d = datetime.strptime(start_iso, "%Y-%m-%d").date()
    end_d = datetime.strptime(end_iso, "%Y-%m-%d").date()
    if end_d < start_d:
        start_d, end_d = end_d, start_d
    return start_d, end_d


def _entries_in_range(entries: list, start_d, end_d) -> list:
    return [e for e in entries if start_d <= datetime.strptime(e["date"], "%Y-%m-%d").date() <= end_d]


def list_entries_in_range(start_iso: str, end_iso: str) -> list:
    """Every logged entry whose date falls within [start_iso, end_iso]
    inclusive, sorted chronologically - used to feed the "Weekly Update"
    LLM consolidation from what was already logged via Daily Update."""
    start_d, end_d = _parse_range(start_iso, end_iso)
    with _lock:
        entries = _load_entries()
    return sorted(_entries_in_range(entries, start_d, end_d), key=lambda e: (e["date"], e["person"]))


def build_range_workbook(start_iso: str, end_iso: str) -> str:
    """"Custom Download": a one-off workbook (never touches the live one)
    holding every entry whose date falls within [start_iso, end_iso]
    inclusive, laid out as a single continuous table - one range banner up
    top, then every date in the range one after another with no per-week
    'Reporting Date' banners in between, and one grand total at the bottom.
    Returns the generated file's basename in OUTPUT_DIR."""
    start_d, end_d = _parse_range(start_iso, end_iso)

    with _lock:
        entries = _load_entries()

    by_date = {}
    for e in _entries_in_range(entries, start_d, end_d):
        by_date.setdefault(e["date"], []).append(e)

    wb, ws = _new_workbook(SHEET_NAME)
    _write_banner(ws, 1, "Medtronic Time Sheet", TITLE_FONT)
    _write_banner(ws, 2, f"Reporting Date ({_fmt(start_d)} - {_fmt(end_d)})", WEEK_FONT)
    row = 3
    _write_header(ws, row)
    row += 1

    row, total = _write_day_groups(ws, row, by_date)
    _write_total_row(ws, row, total)

    filename = f"Medtronic_Time_Sheet_{start_d.isoformat()}_to_{end_d.isoformat()}.xlsx"
    buf = io.BytesIO()
    wb.save(buf)
    blob_store.put(f"output/{filename}", buf.getvalue(), content_type=XLSX_CONTENT_TYPE)
    return filename


def _day_preview(entries: list, date_iso: str) -> dict:
    day_entries = [e for e in entries if e["date"] == date_iso]
    monday, friday = _week_bounds(date_iso)
    week_total = sum(float(e["hrs"]) for e in entries if _week_bounds(e["date"])[0] == monday)
    return {
        "date": date_iso,
        "rows": [{"person": e["person"], "description": e["description"], "hrs": e["hrs"]} for e in day_entries],
        "day_total": sum(float(e["hrs"]) for e in day_entries),
        "week_start": monday.isoformat(),
        "week_end": friday.isoformat(),
        "week_total": week_total,
    }


def get_day(date_iso: str) -> dict:
    """Raises ValueError for a malformed date_iso."""
    datetime.strptime(date_iso, "%Y-%m-%d")
    with _lock:
        return _day_preview(_load_entries(), date_iso)


def add_daily_entry(date_iso: str, person: str, description: str, hrs: float) -> dict:
    """Adds a new (date, person) row, or - if that person already has a row
    for that date - merges into it (Hrs added, Description appended), per
    the "merge same-day resubmits" rule. Rebuilds the live workbook from the
    updated store and returns this date's rows/day-total/week-total."""
    datetime.strptime(date_iso, "%Y-%m-%d")
    with _lock:
        entries = _load_entries()
        existing = next((e for e in entries if e["date"] == date_iso and e["person"] == person), None)
        now = datetime.utcnow().isoformat()

        if existing:
            existing["hrs"] = float(existing["hrs"]) + float(hrs)
            existing["description"] = f"{existing['description']} | {description}" if existing["description"] else description
            existing["updated_at"] = now
        else:
            entries.append({
                "date": date_iso,
                "person": person,
                "description": description,
                "hrs": float(hrs),
                "saved_at": now,
            })

        _save_entries(entries)
        _rebuild_workbook(entries)
        return _day_preview(entries, date_iso)


def remove_daily_entry(date_iso: str, person: str) -> dict:
    """Removes the (date, person) row - the daily-entry equivalent of
    updates_store.remove_entry. Rebuilds the live workbook from the updated
    store and returns this date's rows/day-total/week-total."""
    datetime.strptime(date_iso, "%Y-%m-%d")
    with _lock:
        entries = _load_entries()
        entries = [e for e in entries if not (e["date"] == date_iso and e["person"] == person)]
        _save_entries(entries)
        _rebuild_workbook(entries)
        return _day_preview(entries, date_iso)


def build_weekly_report_from_timesheet(start_iso: str, end_iso: str) -> dict:
    """"Weekly Update" mode: instead of typing a narrative by hand, analyze
    whatever was already logged via Daily Update for [start_iso, end_iso]
    and consolidate it with the same LLM pipeline used for the manual
    narrative flow (consolidate_notes) - one raw-notes block per logged
    entry, "<person>: <description>". Only "Activities Completed This Week"
    (completed_tasks) and "Activities In Process" (milestones) are filled in
    from that analysis; "Next Action" and "Activities To Be Started Next
    Week" are deliberately left blank for the reviewer to fill in by hand.
    Raises ValueError if nothing was logged in that range."""
    entries = list_entries_in_range(start_iso, end_iso)
    if not entries:
        raise ValueError("No daily updates were logged for these dates yet.")

    raw_notes = "\n\n".join(f"{e['person']}: {e['description']}" for e in entries)
    report = consolidate_notes(raw_notes)
    report["next_action"] = []
    report["upcoming_activities"] = []
    return report
