"""
Monthly Attendance Report Generator
-------------------------------------
Builds the office's monthly attendance sheet (same layout/colors as the
HYD_Monthly_sales template) from each employee's raw attendance export.

Raw data (one block per employee, tab-separated, exactly as exported):
    S.NO#  STATUS  DATE  DAY  DISTRIBUTOR NAME (CHECK-IN)  CHECK-IN  ...
    1      PRESENT 2026-09-01 Tuesday ...

Rules applied to every employee, for every day of the month:
  * Holiday days (Friday by default, plus any extra dates you give):
        ABSENT  -> "Holiday"
        PRESENT -> "P"   (they actually worked)
  * Other days:
        PRESENT -> "P"
        ABSENT  -> the first N absences of the month (N = 3 by default,
                   in date order) are "CL" (Casual Leave); the rest are "A".
  * Days that are missing from the raw data are left blank.

Employee master (name, designation, DOJ, base town, CNIC, contact) lives in
ma_data/employees.json -- edit that file when someone joins/leaves.

Usage:
    python generate_monthly_attendance.py Report.xlsx raw_emp1.txt raw_emp2.txt ...
    (raw files are matched to employees.json in order)
"""

import os
import re
import sys
import json
import calendar
from collections import Counter
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ma_data")

HOLIDAY_WEEKDAYS = {4}          # Monday=0 ... Friday=4
CL_ALLOWED = 3
FIRST_DATE_COL = 8              # column H
WEEKDAY_ABBR = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

DATE_FORMATS = ("%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y")


# ----------------------------------------------------------------- parsing --
def _parse_date(s):
    s = (s or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_person_raw(raw_text):
    """Return ({date: 'PRESENT'|'ABSENT'}, [warnings]) from one employee's export."""
    days = {}
    warnings = []
    for line in raw_text.splitlines():
        if not line.strip():
            continue
        fields = [f.strip() for f in line.split("\t")]
        status = None
        d = None
        for f in fields:
            up = f.upper()
            if status is None and up in ("PRESENT", "ABSENT"):
                status = up
            if d is None:
                d = _parse_date(f)
        if d is None:
            continue                      # header / junk line
        if status is None:
            warnings.append(f"{d}: status samajh nahi aaya, ABSENT maana gaya")
            status = "ABSENT"
        if days.get(d) == "PRESENT":
            continue                      # PRESENT always wins on duplicates
        days[d] = status
    return days, warnings


# ------------------------------------------------------------------ logic --
def month_days(year, month):
    n = calendar.monthrange(year, month)[1]
    return [date(year, month, d) for d in range(1, n + 1)]


def detect_month(all_days_maps):
    c = Counter()
    for m in all_days_maps:
        for d in m:
            c[(d.year, d.month)] += 1
    if not c:
        return None
    return c.most_common(1)[0][0]


def classify_month(days_map, year, month, extra_holidays=(), cl_allowed=CL_ALLOWED,
                   holiday_weekdays=HOLIDAY_WEEKDAYS):
    """Return list of (date, code) for each day of the month.
    code in 'P', 'A', 'CL', 'Holiday', or None (no data)."""
    extra = set(extra_holidays)
    cl_used = 0
    out = []
    for d in month_days(year, month):
        status = days_map.get(d)
        is_holiday = d.weekday() in holiday_weekdays or d in extra
        if status is None:
            code = None                   # no data for this day -> blank
        elif status == "PRESENT":
            code = "P"
        else:  # ABSENT
            if is_holiday:
                code = "Holiday"
            elif cl_used < cl_allowed:
                code = "CL"
                cl_used += 1
            else:
                code = "A"
        out.append((d, code))
    return out


# ---------------------------------------------------------------- styling --
DARK = PatternFill("solid", fgColor="632523")       # Accent2, darker 50%
LIGHT = PatternFill("solid", fgColor="F2DCDB")      # Accent2, lighter 80%
MEDIUM = PatternFill("solid", fgColor="E6B8B7")     # Accent2, lighter 60% (holidays)
GREY = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="000000")
MED = Side(style="medium", color="000000")
BOX_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOX_MED = Border(left=MED, right=MED, top=MED, bottom=MED)
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_workbook(employees, region, raw_by_index, out_path, year=None, month=None,
                   extra_holidays=(), cl_allowed=CL_ALLOWED):
    """employees: list of dicts from employees.json
    raw_by_index: {employee_index: raw_text}   (missing/empty -> row stays blank)
    Returns (summary_rows, warnings)."""
    parsed = {}
    warnings = []
    for i, txt in raw_by_index.items():
        if txt and txt.strip():
            m, w = parse_person_raw(txt)
            parsed[i] = m
            warnings += [f"{employees[i]['name']}: {x}" for x in w]

    if year is None or month is None:
        dm = detect_month(parsed.values())
        if dm is None:
            raise ValueError("Kisi bhi data mein koi valid date nahi mili.")
        year, month = dm

    days = month_days(year, month)
    n_days = len(days)
    last_date_col = FIRST_DATE_COL + n_days - 1
    sum_col = last_date_col + 2                      # one gap column, then summary
    labels = ["P.Days", "Absent", "CL", "Holidays"]
    last_col = sum_col + len(labels) - 1

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Title rows
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_date_col)
    c = ws.cell(row=1, column=1, value="Popular group -Dairy division")
    c.font = Font(name="Arial", size=18, bold=True, color="1F497D")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_date_col)
    c2 = ws.cell(row=2, column=1, value=f"Attendance Sheet - {year}")
    c2.font = Font(name="Arial", size=14, bold=True)
    for r in (1, 2):
        for col in range(1, last_date_col + 1):
            cell = ws.cell(row=r, column=col)
            cell.fill = GREY
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, last_date_col + 1):
        ws.cell(row=1, column=col).border = Border(top=MED, bottom=None,
                                                    left=MED if col == 1 else None,
                                                    right=MED if col == last_date_col else None)
    ws.row_dimensions[1].height = 23.25
    ws.row_dimensions[2].height = 18.75

    # Header row 4 (+ weekday row 5)
    head = ["Sr.", "Name", "Designation", "DOJ", "Base Town", "CNIC", "Contact #"]
    for col, h in enumerate(head, 1):
        ws.merge_cells(start_row=4, start_column=col, end_row=5, end_column=col)
        c = ws.cell(row=4, column=col, value=h)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = DARK
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=5, column=col).fill = DARK
        for r in (4, 5):
            ws.cell(row=r, column=col).border = Border(left=THIN, right=THIN, top=MED if r == 4 else None,
                                                       bottom=MED if r == 5 else None)
    for i, d in enumerate(days):
        col = FIRST_DATE_COL + i
        c = ws.cell(row=4, column=col, value=d)
        c.number_format = "d-mmm"
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = DARK
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=THIN, right=THIN, top=MED)
        w = ws.cell(row=5, column=col, value=WEEKDAY_ABBR[d.weekday()])
        w.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        w.fill = DARK
        w.alignment = Alignment(horizontal="center", vertical="center")
        w.border = Border(left=THIN, right=THIN, bottom=MED)
    ws.row_dimensions[4].height = 30.75
    ws.row_dimensions[5].height = 15.75

    # Region row 6
    ws.merge_cells(start_row=6, start_column=2, end_row=6, end_column=4)
    c = ws.cell(row=6, column=2, value=region)
    c.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    for col in range(2, 5):
        ws.cell(row=6, column=col).fill = DARK
        ws.cell(row=6, column=col).border = BOX_MED
    ws.cell(row=6, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[6].height = 19.5
    for j, lab in enumerate(labels):
        c = ws.cell(row=6, column=sum_col + j, value=lab)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = DARK
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX_MED

    # Employee rows
    first_row = 7
    summary_rows = []
    for idx, emp in enumerate(employees):
        r = first_row + idx
        ws.row_dimensions[r].height = 16.5
        a = ws.cell(row=r, column=1, value=emp["sr"])
        a.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        a.fill = DARK
        a.alignment = Alignment(horizontal="center", vertical="center")
        a.border = BOX_THIN

        vals = [
            (2, emp["name"], Font(name="Calibri", size=12, bold=True, color="0070C0"), "left"),
            (3, emp["designation"], Font(name="Calibri", size=12, bold=True, color="7030A0"), "center"),
            (4, emp["doj"], Font(name="Calibri", size=9, bold=True), "center"),
            (5, emp["base_town"], Font(name="Calibri", size=9, bold=True), "center"),
            (6, emp["cnic"], Font(name="Calibri", size=9, bold=True), "center"),
            (7, emp["contact"], Font(name="Calibri", size=9, bold=True), "center"),
        ]
        for col, v, font, al in vals:
            c = ws.cell(row=r, column=col, value=v)
            c.font = font
            c.fill = GREY
            c.alignment = Alignment(horizontal=al, vertical="center")
            c.border = BOX_THIN

        codes = classify_month(parsed.get(idx, {}), year, month, extra_holidays, cl_allowed) \
            if idx in parsed else [(d, None) for d in days]
        for i, (d, code) in enumerate(codes):
            col = FIRST_DATE_COL + i
            c = ws.cell(row=r, column=col, value=code)
            c.font = Font(name="Arial", size=11)
            c.fill = MEDIUM if code == "Holiday" else LIGHT
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = CELL_BORDER

        rng = f"{get_column_letter(FIRST_DATE_COL)}{r}:{get_column_letter(last_date_col)}{r}"
        formulas = [f'=COUNTIF({rng},"P")', f'=COUNTIF({rng},"A")',
                    f'=COUNTIF({rng},"CL")', f'=COUNTIF({rng},"Holiday")']
        for j, fm in enumerate(formulas):
            c = ws.cell(row=r, column=sum_col + j, value=fm)
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = DARK
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BOX_MED

        cnt = Counter(code for _, code in codes if code)
        summary_rows.append({"name": emp["name"], "P": cnt["P"], "A": cnt["A"],
                             "CL": cnt["CL"], "Holiday": cnt["Holiday"],
                             "has_data": idx in parsed})

    # Legend
    last_emp_row = first_row + len(employees) - 1
    all_rng = f"{get_column_letter(FIRST_DATE_COL)}{first_row}:{get_column_letter(last_date_col)}{last_emp_row}"
    legend = [
        ("A", "Absent", "FF0000", "A"),
        ("C. L", "Casual Leave", "0070C0", "CL"),
        ("S. L", "Sick Leave", "00B050", "SL"),
        ("A. L.", "Annual Leave", "974706", "AL"),
    ]
    lr = last_emp_row + 3
    for k, (short, long_, color, code) in enumerate(legend):
        r = lr + k
        fill = PatternFill("solid", fgColor=color)
        for col, v in ((3, short), (4, long_), (5, f'=COUNTIF({all_rng},"{code}")'), (6, code)):
            c = ws.cell(row=r, column=col, value=v)
            c.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
            c.fill = fill
            c.alignment = Alignment(horizontal="left" if col == 4 else "center", vertical="center")
        ws.row_dimensions[r].height = 15.75

    # Widths / view
    widths = {1: 6, 2: 20, 3: 12, 4: 14, 5: 12, 6: 16, 7: 13}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    for i in range(n_days):
        ws.column_dimensions[get_column_letter(FIRST_DATE_COL + i)].width = 8.6
    ws.column_dimensions[get_column_letter(last_date_col + 1)].width = 3
    for j in range(len(labels)):
        ws.column_dimensions[get_column_letter(sum_col + j)].width = 9.5
    ws.freeze_panes = "F7"
    ws.sheet_view.showGridLines = True

    wb.save(out_path)
    return summary_rows, warnings, (year, month)


def load_employees():
    base = os.path.dirname(os.path.abspath(__file__))
    candidates = [os.path.join(DATA_DIR, "employees.json"), os.path.join(base, "employees.json")]
    for path in candidates:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            return data["region"], data["employees"]
    raise FileNotFoundError(
        "employees.json nahi mili. App wale folder mein 'ma_data' naam ka folder banayein "
        "aur employees.json usme rakhein (ya employees.json ko seedha app.py ke sath rakh dein).")


if __name__ == "__main__":
    out = sys.argv[1]
    files = sys.argv[2:]
    region, emps = load_employees()
    raw = {}
    for i, p in enumerate(files):
        with open(p, encoding="utf-8") as f:
            raw[i] = f.read()
    rows, warns, ym = build_workbook(emps, region, raw, out)
    print("Month:", ym)
    for r in rows:
        print(r)
    for w in warns:
        print("WARNING:", w)