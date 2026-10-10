"""
KPK Target vs Achievement Report Generator
--------------------------------------------
Builds the exact KPK Target vs Achievement workbook matching the office's
master Excel sheet:

  1. Summary                   -- Cumulative raw dispatch log for KPK.
  2. KPK Target vs Achievement -- ALL brands on ONE single sheet:
                                  today TW / Maza Chaika, Yazoo-FM, Milker-FM,
                                  Maaza, and Maza/Polly Juices Pet Bottle,
                                  complete with target trays, liter conversions,
                                  achievements, live Excel SUM formulas, and
                                  percentage formulas.
  3. Category                  -- Master product catalog with SKU Name, Category, and Size.

Reference & Config files under kpk_data/:
  - kpk_data/targets.json      -- KPK 9 distributors targets for each brand/size.
  - kpk_data/partname.json     -- KPK distributor name normalization lookup.
  - ach_data/catgery.json      -- Shared master catalog for SKU categorization.
"""

import os
import json
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from generate_dispatch_report import parse_raw_dispatch

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ACH_DATA_DIR = os.path.join(BASE_DIR, "ach_data")
KPK_DATA_DIR = os.path.join(BASE_DIR, "kpk_data")

# Excel Styling
FONT_NAME = "Calibri"
BLUE_TITLE_FILL = PatternFill("solid", fgColor="00B0F0")
HDR_PEACH_FILL = PatternFill("solid", fgColor="F4B183")
DATA_ROW_FILL = PatternFill("solid", fgColor="FCE4D6")
TOTAL_YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
GRAND_TOTAL_FILL = PatternFill("solid", fgColor="FFC000")
SUMMARY_HDR_FILL = PatternFill("solid", fgColor="1F4E79")

THIN_SIDE = Side(style="thin", color="BFBFBF")
BOX_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


def load_reference_data():
    """Load shared SKU catalog, KPK partname lookup, and KPK targets."""
    with open(os.path.join(ACH_DATA_DIR, "catgery.json"), encoding="utf-8") as f:
        catgery = json.load(f)
    with open(os.path.join(KPK_DATA_DIR, "partname.json"), encoding="utf-8") as f:
        partname = json.load(f)
    with open(os.path.join(KPK_DATA_DIR, "targets.json"), encoding="utf-8") as f:
        targets = json.load(f)

    sku_lookup = {
        row["sku"].strip().upper(): (row["category"].strip(), row["size"].strip())
        for row in catgery
    }
    party_lookup = {}
    for row in partname:
        party_lookup[row["raw"].strip().lower()] = (row["normalized"] or row["raw"]).strip()

    return sku_lookup, party_lookup, targets


def save_targets(targets):
    """Save updated KPK monthly targets to kpk_data/targets.json."""
    path = os.path.join(KPK_DATA_DIR, "targets.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(targets, f, indent=2)


def normalize_party(raw_party, party_lookup):
    p = raw_party.strip()
    low = p.lower()
    if low in party_lookup:
        return party_lookup[low]
    
    # Fuzzy matching for KPK 9 distributors
    if "barkatullah" in low:
        return "Haji Barkatullah & Co. - Peshawar (Mega Distributor)"
    if "hasnain" in low and "kohat" in low:
        return "Hasnain Traders -Kohat (Mega Distributor)"
    if "badshah" in low or "mardan" in low:
        return "Badshah Traders -Mardan (Mega Distributor)"
    if "haris" in low:
        return "Haris Traders & Commision Agent -Karak (Mega Distributor)"
    if "mingora" in low or "swat" in low or "tariq" in low:
        return "Tariq Traders - Mingora Swat (Mega Distributor)"
    if "hafiz" in low or "d.i" in low or "di khan" in low:
        return "Hafiz Traders - D.I Khan (Mega Distributor)"
    if "ahmad" in low and "kohat" in low:
        return "Ahmad Traders -Kohat (Mega Distributor)"
    if "ahmad" in low and "peshawar" in low:
        return "Ahmad Traders -Peshawar (Mega Distributor)"
    if "mohaimin" in low:
        return "Abdul Mohaimin Traders -Karak (Mega Distributor)"

    return p


def extract_date_range_label(raw_text):
    """Extract 'Month of  01 TO D-MON -YYYY' from dispatch header."""
    m = re.search(r"Date:\s*(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})", raw_text)
    if m:
        try:
            d2 = datetime.strptime(m.group(2), "%Y-%m-%d")
            return f"Month of  01 TO {d2.day}-{d2.strftime('%b').upper()} -{d2.year}"
        except Exception:
            pass
    return "Month of  01 TO " + datetime.today().strftime("%d-%b-%Y").upper()


def flatten_dispatch_groups(groups):
    """Flatten DO groups for Summary sheet and aggregation."""
    rows = []
    for g in groups:
        for name, case_count, qty in g["items"]:
            mo = ""
            if g.get("dispatch_date"):
                try:
                    mo = datetime.strptime(g["dispatch_date"], "%Y-%m-%d").strftime("%b-%y")
                except Exception:
                    mo = ""
            rows.append({
                "party": g["party"],
                "sku": name,
                "vehicle": g["vehicle"],
                "bilty": g["bilty"],
                "qty": qty,
                "date": g["dispatch_date"],
                "month": mo,
                "freight": g.get("freight", 0.0),
                "ob_user": g.get("ob_user", ""),
                "do": g.get("do", ""),
                "dn": g.get("dn", ""),
            })
    return rows


def load_cumulative_log(log_path):
    if os.path.exists(log_path):
        with open(log_path, encoding="utf-8") as f:
            return json.load(f)
    return []


def save_cumulative_log(log_path, rows):
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2)


def set_month_log(log_path, raw_text):
    """Replace cumulative log with newly pasted month-to-date data."""
    groups = parse_raw_dispatch(raw_text)
    rows = flatten_dispatch_groups(groups)
    save_cumulative_log(log_path, rows)
    return rows, len(rows), len(rows)


def add_day_to_log(log_path, raw_text):
    """Append day's dispatch to cumulative log with deduplication."""
    groups = parse_raw_dispatch(raw_text)
    new_rows = flatten_dispatch_groups(groups)

    existing = load_cumulative_log(log_path)
    existing_keys = {
        (r["party"], r["sku"], r["vehicle"], r["bilty"], r["qty"], r["date"]) for r in existing
    }
    added = 0
    for r in new_rows:
        key = (r["party"], r["sku"], r["vehicle"], r["bilty"], r["qty"], r["date"])
        if key not in existing_keys:
            existing.append(r)
            existing_keys.add(key)
            added += 1

    save_cumulative_log(log_path, existing)
    return existing, added, len(new_rows)


def aggregate_achievements(rows, sku_lookup, party_lookup):
    """Aggregate total trays shipped per normalized distributor and (category, size)."""
    dist_ach = {}
    unmatched_skus = set()
    for r in rows:
        norm_dist = normalize_party(r["party"], party_lookup)
        sku_clean = r["sku"].strip().upper()
        mapping = sku_lookup.get(sku_clean)
        if not mapping:
            unmatched_skus.add(r["sku"])
            continue
        cat, sz = mapping
        if norm_dist not in dist_ach:
            dist_ach[norm_dist] = {}
        dist_ach[norm_dist][(cat, sz)] = dist_ach[norm_dist].get((cat, sz), 0.0) + float(r["qty"])

    return dist_ach, unmatched_skus


def build_summary_sheet(wb, rows):
    """Build Sheet 1: Raw dispatch cumulative rows for KPK."""
    ws = wb.create_sheet("Summary")
    headers = ["Customer :", "SKU NAME", "Vehicle#", "Bilty#", "Sum of Total Tray", "Date", "Month", "Freight", "OB  user"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        c.fill = SUMMARY_HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX_BORDER
    ws.row_dimensions[1].height = 24

    for r_idx, r in enumerate(rows, start=2):
        ws.cell(row=r_idx, column=1, value=r["party"])
        ws.cell(row=r_idx, column=2, value=r["sku"])
        ws.cell(row=r_idx, column=3, value=r["vehicle"])
        ws.cell(row=r_idx, column=4, value=r["bilty"])
        qty_c = ws.cell(row=r_idx, column=5, value=r["qty"])
        qty_c.number_format = "#,##0"
        ws.cell(row=r_idx, column=6, value=r["date"])
        ws.cell(row=r_idx, column=7, value=r["month"])
        fr_c = ws.cell(row=r_idx, column=8, value=r["freight"])
        fr_c.number_format = "#,##0"
        ws.cell(row=r_idx, column=9, value=r["ob_user"])
        for c in range(1, 10):
            ws.cell(row=r_idx, column=c).font = Font(name=FONT_NAME, size=10)
            ws.cell(row=r_idx, column=c).border = BOX_BORDER

    widths = [35, 32, 12, 10, 16, 12, 10, 10, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = True
    return ws


def build_kpk_sheet(wb, month_label, targets, dist_ach):
    """
    Build Sheet 2: ALL brands side-by-side in ONE SINGLE SHEET:
    today TW / Maza Chaika, Yazoo-FM, Milker-FM, Maaza, and Juices Pet Bottle.
    """
    ws = wb.create_sheet("KPK Target vs Achievement")

    # Row 1: Title
    ws.merge_cells("B1:AM1")
    t = ws.cell(row=1, column=2, value="KPK TARGET VS ACHIEVEMENT - Oct-26")
    t.font = Font(name=FONT_NAME, size=16, bold=True)
    t.fill = BLUE_TITLE_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Row 2: Subtitle
    ws.merge_cells("B2:AM2")
    m = ws.cell(row=2, column=2, value=month_label)
    m.font = Font(name=FONT_NAME, size=12, bold=True)
    m.fill = BLUE_TITLE_FILL
    m.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Row 3 & 4: Top headers
    for c_idx, h_text in [(2, "S.No"), (3, "Distributor Name"), (4, "Region")]:
        ws.merge_cells(start_row=3, start_column=c_idx, end_row=4, end_column=c_idx)
        cell = ws.cell(row=3, column=c_idx, value=h_text)
        cell.font = Font(name=FONT_NAME, size=10, bold=True)
        cell.fill = HDR_PEACH_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 1: today TW/Maza Chaika: Col E to M (5 to 13)
    ws.merge_cells("E3:M3")
    c_today = ws.cell(row=3, column=5, value="today TW/Maza Chaika  (Target & Ach)")
    c_today.font = Font(name=FONT_NAME, size=11, bold=True)
    c_today.fill = HDR_PEACH_FILL
    c_today.alignment = Alignment(horizontal="center", vertical="center")

    # Group 2: Yazoo -FM: Col N to S (14 to 19)
    ws.merge_cells("N3:S3")
    c_y = ws.cell(row=3, column=14, value="Yazoo -FM")
    c_y.font = Font(name=FONT_NAME, size=11, bold=True)
    c_y.fill = HDR_PEACH_FILL
    c_y.alignment = Alignment(horizontal="center", vertical="center")

    # Group 3: Milker-FM: Col T to Z (20 to 26)
    ws.merge_cells("T3:Z3")
    c_m = ws.cell(row=3, column=20, value="Milker-FM")
    c_m.font = Font(name=FONT_NAME, size=11, bold=True)
    c_m.fill = HDR_PEACH_FILL
    c_m.alignment = Alignment(horizontal="center", vertical="center")

    # Group 4: Maaza: Col AA to AE (27 to 31)
    ws.merge_cells("AA3:AE3")
    c_mz = ws.cell(row=3, column=27, value="Maaza")
    c_mz.font = Font(name=FONT_NAME, size=11, bold=True)
    c_mz.fill = HDR_PEACH_FILL
    c_mz.alignment = Alignment(horizontal="center", vertical="center")

    # Group 5: Maza/Polly Juices Pet Bottle: Col AF to AM (32 to 39)
    ws.merge_cells("AF3:AM3")
    c_j = ws.cell(row=3, column=32, value="Maza/Polly Juices Pet Bottle")
    c_j.font = Font(name=FONT_NAME, size=11, bold=True)
    c_j.fill = HDR_PEACH_FILL
    c_j.alignment = Alignment(horizontal="center", vertical="center")

    subheaders = [
        # today TW
        (5, "Target  In Tray"),
        (6, "Target  In Liters"),
        (7, "112ml"),
        (8, "180ml"),
        (9, "225ml"),
        (10, "Ach In Tray"),
        (11, "Ach In Liters"),
        (12, "Ach % Tray"),
        (13, "Ach % Liters"),
        # Yazoo -FM
        (14, "Target in trays"),
        (15, "112ml"),
        (16, "180ml"),
        (17, "225ml"),
        (18, "Ach In Tray"),
        (19, "Ach %"),
        # Milker-FM
        (20, "Target  In Tray"),
        (21, "180ml"),
        (22, "225ml"),
        (23, "200ml"),
        (24, "250ml"),
        (25, "Ach In Tray"),
        (26, "Ach %"),
        # Maaza
        (27, "Target in trays"),
        (28, "Maaza 330ml"),
        (29, "Maaza 1000ml"),
        (30, "Ach In Tray"),
        (31, "Ach %"),
        # Maza/Polly Juices Pet Bottle
        (32, "Target  In Tray"),
        (33, "250ml"),
        (34, "500ml"),
        (35, "1000ml"),
        (36, "1500ml"),
        (37, "2000ml"),
        (38, "Ach In Tray"),
        (39, "Ach %"),
    ]

    for c_idx, sub_text in subheaders:
        cell = ws.cell(row=4, column=c_idx, value=sub_text)
        cell.font = Font(name=FONT_NAME, size=9, bold=True)
        cell.fill = HDR_PEACH_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx in (3, 4):
        ws.row_dimensions[r_idx].height = 22 if r_idx == 3 else 30
        for col_idx in range(2, 40):
            ws.cell(row=r_idx, column=col_idx).border = BOX_BORDER
            ws.cell(row=r_idx, column=col_idx).fill = HDR_PEACH_FILL

    # Distributors (all 9)
    curr_row = 5
    for item in targets:
        dist_name = item["distributor"]
        r_ach = dist_ach.get(dist_name, {})

        # today
        t_112 = r_ach.get(("today", "112ml"), 0.0)
        t_180 = r_ach.get(("today", "180ml"), 0.0)
        t_225 = r_ach.get(("today", "225ml"), 0.0)

        # Yazoo
        y_112 = r_ach.get(("Yazoo-FM", "112ml"), 0.0)
        y_180 = r_ach.get(("Yazoo-FM", "180ml"), 0.0)
        y_225 = r_ach.get(("Yazoo-FM", "225ml"), 0.0)

        # Milker
        m_180 = r_ach.get(("Milker-FM", "180ml"), 0.0)
        m_225 = r_ach.get(("Milker-FM", "225ml"), 0.0)
        m_200 = r_ach.get(("Milker-FM", "200ml"), 0.0)
        m_250 = r_ach.get(("Milker-FM", "250ml"), 0.0)

        # Maaza
        mz_330 = r_ach.get(("Juices Pet Bottle", "Maaza 330ml"), 0.0)
        mz_1000 = r_ach.get(("Juices Pet Bottle", "Maaza 1000 new"), 0.0)

        # Juices Pet Bottle
        j_250 = r_ach.get(("Juices Pet Bottle", "Maza 250ml"), 0.0) + r_ach.get(("Juices Pet Bottle", "Polly 250ml"), 0.0)
        j_500 = r_ach.get(("Juices Pet Bottle", "Maza 500ml"), 0.0) + r_ach.get(("Juices Pet Bottle", "Polly 500ml"), 0.0)
        j_1000 = r_ach.get(("Juices Pet Bottle", "Maza 1000ml"), 0.0) + r_ach.get(("Juices Pet Bottle", "Polly 1000ml"), 0.0)
        j_1500 = r_ach.get(("Juices Pet Bottle", "Maza 1500ml"), 0.0) + r_ach.get(("Juices Pet Bottle", "Polly 1500ml"), 0.0)
        j_2000 = r_ach.get(("Juices Pet Bottle", "Maza 2000ml"), 0.0) + r_ach.get(("Juices Pet Bottle", "Polly 2000ml"), 0.0)

        ws.cell(row=curr_row, column=2, value=item["sno"]).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=3, value=dist_name).alignment = Alignment(horizontal="left")
        ws.cell(row=curr_row, column=4, value=item.get("region", "KPK")).alignment = Alignment(horizontal="center")

        # today columns (5 to 13)
        ws.cell(row=curr_row, column=5, value=item.get("today_target_tray", 0)).number_format = "#,##0"
        ws.cell(row=curr_row, column=6, value=item.get("today_target_liters", 0.0)).number_format = "#,##0.00"
        ws.cell(row=curr_row, column=7, value=t_112).number_format = "#,##0"
        ws.cell(row=curr_row, column=8, value=t_180).number_format = "#,##0"
        ws.cell(row=curr_row, column=9, value=t_225).number_format = "#,##0"
        ws.cell(row=curr_row, column=10, value=f"=SUM(G{curr_row}:I{curr_row})").number_format = "#,##0"
        ws.cell(row=curr_row, column=11, value=f"=(G{curr_row}*24*0.112)+(H{curr_row}*24*0.18)+(I{curr_row}*27*0.225)").number_format = "#,##0.00"
        ws.cell(row=curr_row, column=12, value=f"=IFERROR(J{curr_row}/E{curr_row},0)").number_format = "0.0%"
        ws.cell(row=curr_row, column=13, value=f"=IFERROR(K{curr_row}/F{curr_row},0)").number_format = "0.0%"

        # Yazoo columns (14 to 19)
        ws.cell(row=curr_row, column=14, value=item.get("yazoo_target", 0)).number_format = "#,##0"
        ws.cell(row=curr_row, column=15, value=y_112).number_format = "#,##0"
        ws.cell(row=curr_row, column=16, value=y_180).number_format = "#,##0"
        ws.cell(row=curr_row, column=17, value=y_225).number_format = "#,##0"
        ws.cell(row=curr_row, column=18, value=f"=SUM(O{curr_row}:Q{curr_row})").number_format = "#,##0"
        ws.cell(row=curr_row, column=19, value=f"=IFERROR(R{curr_row}/N{curr_row},0)").number_format = "0.0%"

        # Milker columns (20 to 26)
        ws.cell(row=curr_row, column=20, value=item.get("milker_target", 0)).number_format = "#,##0"
        ws.cell(row=curr_row, column=21, value=m_180).number_format = "#,##0"
        ws.cell(row=curr_row, column=22, value=m_225).number_format = "#,##0"
        ws.cell(row=curr_row, column=23, value=m_200).number_format = "#,##0"
        ws.cell(row=curr_row, column=24, value=m_250).number_format = "#,##0"
        ws.cell(row=curr_row, column=25, value=f"=SUM(U{curr_row}:X{curr_row})").number_format = "#,##0"
        ws.cell(row=curr_row, column=26, value=f"=IFERROR(Y{curr_row}/T{curr_row},0)").number_format = "0.0%"

        # Maaza columns (27 to 31)
        ws.cell(row=curr_row, column=27, value=item.get("maaza_target", 0)).number_format = "#,##0"
        ws.cell(row=curr_row, column=28, value=mz_330).number_format = "#,##0"
        ws.cell(row=curr_row, column=29, value=mz_1000).number_format = "#,##0"
        ws.cell(row=curr_row, column=30, value=f"=SUM(AB{curr_row}:AC{curr_row})").number_format = "#,##0"
        ws.cell(row=curr_row, column=31, value=f"=IFERROR(AD{curr_row}/AA{curr_row},0)").number_format = "0.0%"

        # Maza/Polly Juices columns (32 to 39)
        ws.cell(row=curr_row, column=32, value=item.get("jpb_target", 0)).number_format = "#,##0"
        ws.cell(row=curr_row, column=33, value=j_250).number_format = "#,##0"
        ws.cell(row=curr_row, column=34, value=j_500).number_format = "#,##0"
        ws.cell(row=curr_row, column=35, value=j_1000).number_format = "#,##0"
        ws.cell(row=curr_row, column=36, value=j_1500).number_format = "#,##0"
        ws.cell(row=curr_row, column=37, value=j_2000).number_format = "#,##0"
        ws.cell(row=curr_row, column=38, value=f"=SUM(AG{curr_row}:AK{curr_row})").number_format = "#,##0"
        ws.cell(row=curr_row, column=39, value=f"=IFERROR(AL{curr_row}/AF{curr_row},0)").number_format = "0.0%"

        for col_idx in range(2, 40):
            c_node = ws.cell(row=curr_row, column=col_idx)
            c_node.font = Font(name=FONT_NAME, size=10)
            c_node.border = BOX_BORDER
            c_node.fill = DATA_ROW_FILL
            if col_idx in (2, 4):
                c_node.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 3:
                c_node.alignment = Alignment(horizontal="left", vertical="center")
            else:
                c_node.alignment = Alignment(horizontal="right", vertical="center")

        ws.row_dimensions[curr_row].height = 18
        curr_row += 1

    # Grand Total Row (Row 14)
    gt_row = curr_row
    ws.merge_cells(start_row=gt_row, start_column=2, end_row=gt_row, end_column=4)
    gt_lbl = ws.cell(row=gt_row, column=2, value="Grand Total KPK (Sajid Hussain ZSM)")
    gt_lbl.font = Font(name=FONT_NAME, size=11, bold=True)
    gt_lbl.alignment = Alignment(horizontal="center", vertical="center")

    # Sum columns for 5 to 13
    sum_cols = [5, 6, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 20, 21, 22, 23, 24, 25, 27, 28, 29, 30, 32, 33, 34, 35, 36, 37, 38]
    for col_idx in sum_cols:
        cl = get_column_letter(col_idx)
        ws.cell(row=gt_row, column=col_idx, value=f"=SUM({cl}5:{cl}{gt_row-1})").number_format = "#,##0" if col_idx not in (6, 11) else "#,##0.00"

    # Percentage formulas for totals
    ws.cell(row=gt_row, column=12, value=f"=IFERROR(J{gt_row}/E{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=13, value=f"=IFERROR(K{gt_row}/F{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=19, value=f"=IFERROR(R{gt_row}/N{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=26, value=f"=IFERROR(Y{gt_row}/T{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=31, value=f"=IFERROR(AD{gt_row}/AA{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=39, value=f"=IFERROR(AL{gt_row}/AF{gt_row},0)").number_format = "0.0%"

    for col_idx in range(2, 40):
        c_gt = ws.cell(row=gt_row, column=col_idx)
        c_gt.font = Font(name=FONT_NAME, size=11, bold=True)
        c_gt.fill = GRAND_TOTAL_FILL
        c_gt.border = BOX_BORDER
        if col_idx not in (2, 3, 4):
            c_gt.alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[gt_row].height = 22

    # Column widths
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 46
    ws.column_dimensions["D"].width = 12
    for c_idx in range(5, 40):
        ws.column_dimensions[get_column_letter(c_idx)].width = 11

    ws.freeze_panes = "E5"
    ws.sheet_view.showGridLines = True
    return ws


def build_category_sheet(wb):
    """Build Sheet 3: Master product catalog."""
    with open(os.path.join(ACH_DATA_DIR, "catgery.json"), encoding="utf-8") as f:
        catgery = json.load(f)
    ws = wb.create_sheet("Category")
    headers = ["SKU Name", "Category", "Size"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        c.fill = SUMMARY_HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX_BORDER
    ws.row_dimensions[1].height = 22

    for r_idx, row in enumerate(catgery, start=2):
        ws.cell(row=r_idx, column=1, value=row.get("sku", ""))
        ws.cell(row=r_idx, column=2, value=row.get("category", ""))
        ws.cell(row=r_idx, column=3, value=row.get("size", ""))
        for col_idx in (1, 2, 3):
            ws.cell(row=r_idx, column=col_idx).font = Font(name=FONT_NAME, size=10)
            ws.cell(row=r_idx, column=col_idx).border = BOX_BORDER

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 18
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = True
    return ws


def build_workbook(cumulative_rows, month_label, out_path):
    """Generate the single-sheet KPK Target vs Achievement Excel workbook."""
    sku_lookup, party_lookup, targets = load_reference_data()
    dist_ach, unmatched_skus = aggregate_achievements(cumulative_rows, sku_lookup, party_lookup)

    wb = Workbook()
    wb.remove(wb.active)

    build_summary_sheet(wb, cumulative_rows)
    build_kpk_sheet(wb, month_label, targets, dist_ach)
    build_category_sheet(wb)

    wb.save(out_path)
    return out_path, unmatched_skus
