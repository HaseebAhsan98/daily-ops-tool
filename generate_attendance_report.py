"""
Daily Attendance Report Generator
-----------------------------------
Reads the raw "Order Summary" export from the S&D software (per-distributor
rows) and consolidates it into one Present/Absent row per employee.

Logic:
- Each employee can appear in many rows (one per distributor they cover).
- If ANY row for that employee shows PRESENT -> employee is PRESENT overall.
- If ALL rows for that employee show ABSENT  -> employee is ABSENT overall.
- Report should be generated at the fixed daily cutoff time (e.g. 10:30 AM);
  whatever the software shows at that moment is what counts.

Usage:
    python generate_attendance_report.py "Order_Summary_14-Sep-2026.xls" "Attendance_Report_14-Sep-2026.xlsx"
"""

import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def load_raw_export(path):
    """The S&D software exports .xls files that are actually HTML tables."""
    tables = pd.read_html(path)
    raw = tables[0]
    header = raw.iloc[0]
    data = raw[1:].copy()
    data.columns = header
    return data


def consolidate_attendance(data):
    """One row per employee: Present if any distributor-row is Present, else Absent.
    For a Present employee, Check-in Time and Distributor are taken from the first
    row where they were marked Present (their earliest check-in of the day)."""

    def summarize_employee(g):
        present_rows = g[g["Status"] == "PRESENT"]
        if len(present_rows):
            first = present_rows.iloc[0]
            status = "PRESENT"
            check_in = first["Check in"]
            distributor = first["Distributor"]
        else:
            status = "ABSENT"
            check_in = "--"
            distributor = "--"
        return pd.Series({
            "Final_Status": status,
            "Check_In": check_in,
            "Distributor_Location": distributor,
        })

    summary = (
        data.groupby(["Emp ID", "Emp Name", "Designation"], sort=False)
        .apply(summarize_employee, include_groups=False)
        .reset_index()
    )
    # Absent first, Present last -- within each designation
    summary["sort_key"] = summary["Final_Status"].map({"ABSENT": 0, "PRESENT": 1})
    summary = summary.sort_values(["Designation", "sort_key", "Emp Name"]).drop(columns="sort_key")
    return summary


def write_report(summary, out_path, report_date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"

    title_font = Font(name="Arial", size=14, bold=True)
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F5597")
    absent_fill = PatternFill("solid", fgColor="F8CBAD")
    present_fill = PatternFill("solid", fgColor="C6E0B4")
    body_font = Font(name="Arial", size=11)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:G1")
    ws["A1"] = f"Daily Attendance Report - {report_date}"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    headers = ["Emp ID", "Emp Name", "Designation", "Status", "Check-in Time", "Distributor Location", "Distributors Covered"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border

    row_idx = 4
    for _, r in summary.iterrows():
        ws.cell(row=row_idx, column=1, value=r["Emp ID"]).border = border
        ws.cell(row=row_idx, column=2, value=r["Emp Name"]).border = border
        ws.cell(row=row_idx, column=3, value=r["Designation"]).border = border
        status_cell = ws.cell(row=row_idx, column=4, value=r["Final_Status"])
        status_cell.border = border
        status_cell.alignment = center
        status_cell.font = Font(name="Arial", size=11, bold=True)
        status_cell.fill = absent_fill if r["Final_Status"] == "ABSENT" else present_fill
        ws.cell(row=row_idx, column=5, value=r["Check_In"]).border = border
        ws.cell(row=row_idx, column=6, value=r["Distributor_Location"]).border = border
        ws.cell(row=row_idx, column=7, value=r.get("Distributors", "")).border = border
        for col in (1, 2, 3, 5, 6, 7):
            ws.cell(row=row_idx, column=col).font = body_font
        row_idx += 1

    # Summary counts at the bottom
    total = len(summary)
    absent_n = (summary["Final_Status"] == "ABSENT").sum()
    present_n = total - absent_n
    row_idx += 1
    for label, val in (("Total", total), ("Absent", int(absent_n)), ("Present", int(present_n))):
        ws.cell(row=row_idx, column=1, value=label).font = Font(name="Arial", bold=True)
        ws.cell(row=row_idx, column=2, value=val).font = Font(name="Arial", bold=True)
        row_idx += 1

    widths = [10, 22, 16, 12, 14, 40, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(out_path)


if __name__ == "__main__":
    in_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else "Attendance_Report.xlsx"
    report_date = sys.argv[3] if len(sys.argv) > 3 else pd.Timestamp.today().strftime("%d-%b-%Y")

    raw = load_raw_export(in_path)
    summary = consolidate_attendance(raw)
    # attach distributor counts for reference
    dist_counts = raw.groupby("Emp ID")["Distributor"].count()
    summary["Distributors"] = summary["Emp ID"].map(dist_counts)
    write_report(summary, out_path, report_date)
    print(f"Report written to {out_path}")