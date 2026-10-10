"""
KPK Target vs Achievement Report Generator
--------------------------------------------
Builds the exact KPK Target vs Achievement workbook matching the office's
master Excel sheet (TSE Wise today, Milker & Juice Pet Bottle):

  1. Summary                   -- Cumulative raw dispatch log for KPK.
  2. KPK Target vs Achievement -- EXACT visual replica of the KPK master sheet:
                                  16 columns total:
                                  S.No | Distributor Name | TSE Range Buiness |
                                  today [Target In Tray, Ach In Tray, Ach % Tray] |
                                  Milker-FM [Target In Tray, Ach In Tray, Ach %] |
                                  JuicePet Bottle [Target In Tray, Ach In Tray, Ach %] |
                                  Maza Chaika [In tray] |
                                  Yazo [Targ In tray, Ach In tray, Ach %]
                                  plus Total TSE Vaccant and Grand Total Sajd Hussain ZSM.
  3. Category                  -- Master product catalog with SKU Name, Category, and Size.

Reference & Config files under kpk_data/:
  - kpk_data/targets.json      -- KPK distributors with targets for each brand.
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
HDR_PEACH_FILL = PatternFill("solid", fgColor="F8CBAD")
SUMMARY_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
PINK_FILL = PatternFill("solid", fgColor="FCE4D6")

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

    # Fuzzy matching for KPK distributors
    if "barkatullah" in low:
        return "Haji Barkatullah & Co. Peshawar (Mega Distributor) Dairy"
    if "badshah" in low or "mardan" in low:
        return "Badshah Traders-Mardan (Mega Distributor) Dairy"
    if "haris" in low:
        return "Haris Traders & Commission Agent-Karak (Mega Distributor) Dairy"
    if "mingora" in low or "swat" in low or "tariq" in low:
        return "Tariq Traders-Mingora Swat (Mega Distributor) Dairy"
    if "nowroz" in low or "hasnain" in low:
        return "Nowroz & Sons-Kohat ( Mega Distributor) Dairy"
    if "ahmad" in low and "kohat" in low:
        return "Ahmad Traders-Kohat (Mega Distributor) Dairy"
    if "ahmad" in low and "peshawar" in low:
        return "Ahmad Traders-Peshawar (Mega Distributor) Dairy"
    if "hafiz" in low or "d.i" in low or "di khan" in low:
        return "Hafiz Traders-D.I Khan (MEGA Distributor) Dairy"
    if "mohaimin" in low:
        return "Abdul Mohaimin Traders-Karak (Mega Distributor) Dairy"
    if "vaccent" in low or "vaccant" in low:
        return "Vaccent"

    return p


def extract_date_range_label(raw_text):
    """Extract 'MONTH OF 1 TO {day} {MON}-{YEAR}' from dispatch header."""
    m = re.search(r"Date:\s*(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})", raw_text)
    if m:
        try:
            d2 = datetime.strptime(m.group(2), "%Y-%m-%d")
            return f"MONTH OF 1 TO {d2.day} {d2.strftime('%b').upper()}-{d2.year}"
        except Exception:
            pass
    return "MONTH OF 1 TO " + datetime.today().strftime("%d %b-%Y").upper()


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
    """
    Aggregate total trays shipped per normalized distributor into the 5 KPK categories:
      - 'today'
      - 'Milker-FM'
      - 'JuicePet Bottle'
      - 'Maza Chaika'
      - 'Yazo'
    """
    dist_ach = {}
    unmatched_skus = set()
    for r in rows:
        norm_dist = normalize_party(r["party"], party_lookup)
        sku_clean = r["sku"].strip().upper()

        if "CHAIKA" in sku_clean or "CHAYKA" in sku_clean:
            brand = "Maza Chaika"
        else:
            mapping = sku_lookup.get(sku_clean)
            if mapping:
                cat, _ = mapping
                cat_lower = cat.lower().strip()
                if "today" in cat_lower:
                    brand = "today"
                elif "milker" in cat_lower:
                    brand = "Milker-FM"
                elif "yazoo" in cat_lower or "yazo" in cat_lower:
                    brand = "Yazo"
                elif "juice" in cat_lower or "maza" in cat_lower or "polly" in cat_lower:
                    brand = "JuicePet Bottle"
                else:
                    brand = cat
            else:
                s_low = sku_clean.lower()
                if "today" in s_low:
                    brand = "today"
                elif "milker" in s_low:
                    brand = "Milker-FM"
                elif "yazoo" in s_low or "yazo" in s_low:
                    brand = "Yazo"
                elif "juice" in s_low or "maza" in s_low or "polly" in s_low or "bottle" in s_low:
                    brand = "JuicePet Bottle"
                else:
                    unmatched_skus.add(r["sku"])
                    brand = "Other"

        if norm_dist not in dist_ach:
            dist_ach[norm_dist] = {}
        dist_ach[norm_dist][brand] = dist_ach[norm_dist].get(brand, 0.0) + float(r.get("qty", 0))

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
    Build Sheet 2: EXACT match of KPK Target vs Achievement sheet (16 columns).
    """
    ws = wb.create_sheet("KPK Target vs Achievement")

    # Row 1: Title
    ws.merge_cells("A1:P1")
    t = ws.cell(row=1, column=1, value="KPK TARGET VS ACHIEVEMENT  TSE  WISE today, Milker & Juice Pet Bottle")
    t.font = Font(name=FONT_NAME, size=14, bold=True, color="000000")
    t.fill = BLUE_TITLE_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Row 2: Subtitle
    ws.merge_cells("A2:P2")
    m = ws.cell(row=2, column=1, value=month_label)
    m.font = Font(name=FONT_NAME, size=12, bold=True, color="000000")
    m.fill = BLUE_TITLE_FILL
    m.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Row 3 & 4: Top headers
    ws.merge_cells("A3:A4")
    c_sno = ws.cell(row=3, column=1, value="S.No")
    c_sno.fill = HDR_PEACH_FILL
    c_sno.font = Font(name=FONT_NAME, size=10, bold=True)
    c_sno.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B3:B4")
    c_dist = ws.cell(row=3, column=2, value="Distributor Name")
    c_dist.fill = HDR_PEACH_FILL
    c_dist.font = Font(name=FONT_NAME, size=10, bold=True)
    c_dist.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C3:C4")
    c_tse = ws.cell(row=3, column=3, value="TSE Range Buiness")
    c_tse.fill = HDR_PEACH_FILL
    c_tse.font = Font(name=FONT_NAME, size=10, bold=True)
    c_tse.alignment = Alignment(horizontal="center", vertical="center")

    # Group 1: today: Col D to F (4 to 6)
    ws.merge_cells("D3:F3")
    c_today = ws.cell(row=3, column=4, value="today")
    c_today.fill = PatternFill("solid", fgColor="0070C0")
    c_today.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    c_today.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, sub in [(4, "Target In Tray"), (5, "Ach In Tray"), (6, "Ach % Tray")]:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = PatternFill("solid", fgColor="0070C0")
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 2: Milker-FM: Col G to I (7 to 9)
    ws.merge_cells("G3:I3")
    c_m = ws.cell(row=3, column=7, value="Milker-FM")
    c_m.fill = PatternFill("solid", fgColor="70AD47")
    c_m.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
    c_m.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, sub in [(7, "Target In Tray"), (8, "Ach In Tray"), (9, "Ach %")]:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = PatternFill("solid", fgColor="70AD47")
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 3: JuicePet Bottle: Col J to L (10 to 12)
    ws.merge_cells("J3:L3")
    c_j = ws.cell(row=3, column=10, value="JuicePet Bottle")
    c_j.fill = YELLOW_FILL
    c_j.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
    c_j.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, sub in [(10, "Target In Tray"), (11, "Ach In Tray"), (12, "Ach %")]:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = YELLOW_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 4: Maza Chaika: Col M (13)
    c_ch = ws.cell(row=3, column=13, value="Maza Chaika")
    c_ch.fill = HDR_PEACH_FILL
    c_ch.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
    c_ch.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    c_ch_sub = ws.cell(row=4, column=13, value="In tray")
    c_ch_sub.fill = HDR_PEACH_FILL
    c_ch_sub.font = Font(name=FONT_NAME, size=9, bold=True, color="000000")
    c_ch_sub.alignment = Alignment(horizontal="center", vertical="center")

    # Group 5: Yazo: Col N to P (14 to 16)
    ws.merge_cells("N3:P3")
    c_y = ws.cell(row=3, column=14, value="Yazo")
    c_y.fill = PatternFill("solid", fgColor="FF0000")
    c_y.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
    c_y.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, sub in [(14, "Targ In tray"), (15, "Ach In tray"), (16, "Ach %")]:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = PatternFill("solid", fgColor="FF0000")
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Format header borders
    for r in (3, 4):
        ws.row_dimensions[r].height = 24
        for col in range(1, 17):
            ws.cell(row=r, column=col).border = BOX_BORDER

    # Data Rows
    start_r = 5
    for idx, t in enumerate(targets):
        r = start_r + idx
        ws.row_dimensions[r].height = 20
        dist_name = t["distributor"]
        tse_name = t.get("tse", "")
        r_ach = dist_ach.get(dist_name, {})

        # Col A: S.No
        ws.cell(row=r, column=1, value=t.get("sno", idx + 1)).alignment = Alignment(horizontal="center", vertical="center")

        # Col B: Distributor Name
        ws.cell(row=r, column=2, value=dist_name).alignment = Alignment(horizontal="left", vertical="center")

        # Col C: TSE Range Buiness
        c_tse_val = ws.cell(row=r, column=3, value=tse_name)
        c_tse_val.alignment = Alignment(horizontal="center", vertical="center")
        if "vaccant" in tse_name.lower():
            c_tse_val.font = Font(name=FONT_NAME, size=10, bold=True, color="FF0000")
            c_tse_val.fill = PINK_FILL
        else:
            c_tse_val.font = Font(name=FONT_NAME, size=10)

        # Col D: today Target
        tgt_today = t.get("today_target", t.get("today_target_tray", 0))
        ws.cell(row=r, column=4, value=tgt_today).number_format = "#,##0"

        # Col E: today Ach
        ach_today = r_ach.get("today", 0.0)
        ws.cell(row=r, column=5, value=ach_today).number_format = "#,##0"

        # Col F: today Ach %
        ws.cell(row=r, column=6, value=f"=IFERROR(E{r}/D{r}, 0)").number_format = "0%"

        # Col G: Milker Target
        tgt_milker = t.get("milker_target", 0)
        ws.cell(row=r, column=7, value=tgt_milker).number_format = "#,##0"

        # Col H: Milker Ach
        ach_milker = r_ach.get("Milker-FM", 0.0)
        ws.cell(row=r, column=8, value=ach_milker).number_format = "#,##0"

        # Col I: Milker Ach %
        ws.cell(row=r, column=9, value=f"=IFERROR(H{r}/G{r}, 0)").number_format = "0%"

        # Col J: JuicePet Target
        tgt_jpb = t.get("jpb_target", 0)
        ws.cell(row=r, column=10, value=tgt_jpb).number_format = "#,##0"

        # Col K: JuicePet Ach
        ach_jpb = r_ach.get("JuicePet Bottle", 0.0)
        ws.cell(row=r, column=11, value=ach_jpb).number_format = "#,##0"

        # Col L: JuicePet Ach %
        ws.cell(row=r, column=12, value=f"=IFERROR(K{r}/J{r}, 0)").number_format = "0%"

        # Col M: Maza Chaika
        ach_chaika = r_ach.get("Maza Chaika", 0.0)
        ws.cell(row=r, column=13, value=ach_chaika).number_format = "#,##0"

        # Col N: Yazo Target
        tgt_yazo = t.get("yazoo_target", 0)
        ws.cell(row=r, column=14, value=tgt_yazo).number_format = "#,##0"

        # Col O: Yazo Ach
        ach_yazo = r_ach.get("Yazo", 0.0)
        ws.cell(row=r, column=15, value=ach_yazo).number_format = "#,##0"

        # Col P: Yazo Ach %
        ws.cell(row=r, column=16, value=f"=IFERROR(O{r}/N{r}, 0)").number_format = "0%"

        for c_idx in range(1, 17):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = BOX_BORDER
            if c_idx not in (2, 3):
                cell.font = Font(name=FONT_NAME, size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")

    end_r = start_r + len(targets) - 1

    # Row: Total TSE Vaccant
    r_vac = end_r + 1
    ws.row_dimensions[r_vac].height = 22
    ws.merge_cells(f"A{r_vac}:C{r_vac}")
    lbl_vac = ws.cell(row=r_vac, column=1, value="Total TSE Vaccant")
    lbl_vac.font = Font(name=FONT_NAME, size=11, bold=True)
    lbl_vac.alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=r_vac, column=4, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", D5:D{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=5, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", E5:E{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=6, value=f'=IFERROR(E{r_vac}/D{r_vac}, 0)').number_format = "0%"

    ws.cell(row=r_vac, column=7, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", G5:G{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=8, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", H5:H{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=9, value=f'=IFERROR(H{r_vac}/G{r_vac}, 0)').number_format = "0%"

    ws.cell(row=r_vac, column=10, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", J5:J{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=11, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", K5:K{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=12, value=f'=IFERROR(K{r_vac}/J{r_vac}, 0)').number_format = "0%"

    ws.cell(row=r_vac, column=13, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", M5:M{end_r})').number_format = "#,##0"

    ws.cell(row=r_vac, column=14, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", N5:N{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=15, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", O5:O{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=16, value=f'=IFERROR(O{r_vac}/N{r_vac}, 0)').number_format = "0%"

    for c_idx in range(1, 17):
        cell = ws.cell(row=r_vac, column=c_idx)
        cell.fill = YELLOW_FILL
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
        cell.border = BOX_BORDER
        if c_idx > 3:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Row: Grand Total Sajd Hussain ZSM
    r_gt = r_vac + 1
    ws.row_dimensions[r_gt].height = 24
    ws.merge_cells(f"A{r_gt}:C{r_gt}")
    lbl_gt = ws.cell(row=r_gt, column=1, value="Grand Total  Sajd Hussain ZSM")
    lbl_gt.font = Font(name=FONT_NAME, size=11, bold=True)
    lbl_gt.alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=r_gt, column=4, value=f'=SUM(D5:D{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=5, value=f'=SUM(E5:E{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=6, value=f'=IFERROR(E{r_gt}/D{r_gt}, 0)').number_format = "0%"

    ws.cell(row=r_gt, column=7, value=f'=SUM(G5:G{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=8, value=f'=SUM(H5:H{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=9, value=f'=IFERROR(H{r_gt}/G{r_gt}, 0)').number_format = "0%"

    ws.cell(row=r_gt, column=10, value=f'=SUM(J5:J{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=11, value=f'=SUM(K5:K{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=12, value=f'=IFERROR(K{r_gt}/J{r_gt}, 0)').number_format = "0%"

    ws.cell(row=r_gt, column=13, value=f'=SUM(M5:M{end_r})').number_format = "#,##0"

    ws.cell(row=r_gt, column=14, value=f'=SUM(N5:N{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=15, value=f'=SUM(O5:O{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=16, value=f'=IFERROR(O{r_gt}/N{r_gt}, 0)').number_format = "0%"

    for c_idx in range(1, 17):
        cell = ws.cell(row=r_gt, column=c_idx)
        cell.fill = YELLOW_FILL
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
        cell.border = BOX_BORDER
        if c_idx > 3:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Column Widths
    col_widths = {
        "A": 6,   # S.No
        "B": 44,  # Distributor Name
        "C": 18,  # TSE Range Buiness
        "D": 13,  # today Target
        "E": 12,  # today Ach
        "F": 12,  # today Ach %
        "G": 13,  # Milker Target
        "H": 13,  # Milker Ach
        "I": 11,  # Milker Ach %
        "J": 13,  # JuicePet Target
        "K": 13,  # JuicePet Ach
        "L": 11,  # JuicePet Ach %
        "M": 12,  # Maza Chaika
        "N": 12,  # Yazo Target
        "O": 12,  # Yazo Ach
        "P": 11,  # Yazo Ach %
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    ws.freeze_panes = "D5"
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
    """Generate the exact 16-column KPK Target vs Achievement Excel workbook."""
    sku_lookup, party_lookup, targets = load_reference_data()
    dist_ach, unmatched_skus = aggregate_achievements(cumulative_rows, sku_lookup, party_lookup)

    wb = Workbook()
    wb.remove(wb.active)

    build_summary_sheet(wb, cumulative_rows)
    build_kpk_sheet(wb, month_label, targets, dist_ach)
    build_category_sheet(wb)

    wb.save(out_path)
    return out_path, unmatched_skus
