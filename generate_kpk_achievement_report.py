"""
KPK Target vs Achievement Report Generator
--------------------------------------------
Builds the exact KPK Target vs Achievement workbook matching the office's
master Excel sheet:

  1. Summary                   -- Cumulative raw dispatch log for KPK.
  2. KPK Target vs Achievement -- Master sheet with all categories & size breakdowns:
                                  today TW/Maza Chaika, Yazoo -FM, Milker-FM (180ml & 225ml ONLY),
                                  Maaza, and Maza/Polly Juices Pet Bottle.
                                  Colors matching master design (Blue, Red, Olive Green, Orange, Yellow).
                                  Summary rows: Total TSE Vaccant & Grand Total KPK (Sajid Hussain ZSM).
  3. Category                  -- Master product catalog with SKU Name, Category, and Size.

Reference & Config files under kpk_data/:
  - kpk_data/targets.json      -- KPK distributors with targets for each brand/size.
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
PEACH_FILL = PatternFill("solid", fgColor="F8CBAD")
SUMMARY_HDR_FILL = PatternFill("solid", fgColor="1F4E79")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
PINK_FILL = PatternFill("solid", fgColor="FCE4D6")

# Group Header Fills
TODAY_BLUE_FILL = PatternFill("solid", fgColor="0070C0")
YAZOO_RED_FILL = PatternFill("solid", fgColor="C00000")
MILKER_GREEN_FILL = PatternFill("solid", fgColor="70AD47")
MAAZA_ORANGE_FILL = PatternFill("solid", fgColor="ED7D31")
JUICES_YELLOW_FILL = PatternFill("solid", fgColor="FFC000")

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
    """Extract 'Month of  01 TO {day}-{MON}-{YEAR}' from dispatch header."""
    m = re.search(r"Date:\s*(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})", raw_text)
    if m:
        try:
            d2 = datetime.strptime(m.group(2), "%Y-%m-%d")
            return f"Month of  01 TO {d2.day}-{d2.strftime('%b').upper()}-{d2.year}"
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
    """
    Aggregate total trays shipped per normalized distributor and (category, size).
    Guarantees KPK Milker only categorizes into '180ml' and '225ml'.
    """
    dist_ach = {}
    unmatched_skus = set()
    for r in rows:
        norm_dist = normalize_party(r["party"], party_lookup)
        sku_clean = r["sku"].strip().upper()

        mapping = sku_lookup.get(sku_clean)
        if not mapping:
            # Fallback deduction
            s_low = sku_clean.lower()
            if "today" in s_low or "chaika" in s_low:
                cat = "today"
                sz = "180ml" if "180" in s_low else ("225ml" if "225" in s_low else "112ml")
            elif "milker" in s_low:
                cat = "Milker-FM"
                sz = "180ml" if "180" in s_low else "225ml"
            elif "yazoo" in s_low or "yazo" in s_low:
                cat = "Yazoo-FM"
                sz = "112ml" if "112" in s_low else ("180ml" if "180" in s_low else "225ml")
            elif "maaza 330" in s_low:
                cat = "Juices Pet Bottle"
                sz = "Maaza 330ml"
            elif "maaza 1000" in s_low:
                cat = "Juices Pet Bottle"
                sz = "Maaza 1000 new"
            elif "juice" in s_low or "maza" in s_low or "polly" in s_low or "bottle" in s_low:
                cat = "Juices Pet Bottle"
                sz = "Maza 250ml"
            else:
                unmatched_skus.add(r["sku"])
                continue
        else:
            cat, sz = mapping
            cat = cat.strip()
            sz = sz.strip()

        # KPK Milker strictness: ONLY 180ml and 225ml!
        if "milker" in cat.lower():
            cat = "Milker-FM"
            # If SKU name contains 180, it's 180ml. If 225, it's 225ml.
            if "180" in sku_clean:
                sz = "180ml"
            elif "225" in sku_clean:
                sz = "225ml"
            elif sz in ("180ml", "185ml"):
                sz = "180ml"
            else:
                sz = "225ml"

        if norm_dist not in dist_ach:
            dist_ach[norm_dist] = {}
        dist_ach[norm_dist][(cat, sz)] = dist_ach[norm_dist].get((cat, sz), 0.0) + float(r.get("qty", 0))

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
    Build Sheet 2: KPK Target vs Achievement with all categories, exact colors,
    and NO 200ml/250ml in Milker-FM (Only 180ml and 225ml).
    Total Columns: 28 (A to AB).
    """
    ws = wb.create_sheet("KPK Target vs Achievement")

    # Row 1: Title
    ws.merge_cells("B1:AB1")
    t = ws.cell(row=1, column=2, value="KPK TARGET VS ACHIEVEMENT - Oct-26")
    t.font = Font(name=FONT_NAME, size=14, bold=True, color="000000")
    t.fill = BLUE_TITLE_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Row 2: Subtitle
    ws.merge_cells("B2:AB2")
    m = ws.cell(row=2, column=2, value=month_label)
    m.font = Font(name=FONT_NAME, size=12, bold=True, color="000000")
    m.fill = BLUE_TITLE_FILL
    m.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Row 3 & 4: Top headers
    # Col A: Hidden helper or empty, let's start from B as S.No or A as S.No
    # Let's use Col A as S.No, B as Distributor, C as TSE Range Buiness
    ws.merge_cells("A3:A4")
    c_sno = ws.cell(row=3, column=1, value="S.No")
    c_sno.fill = PEACH_FILL
    c_sno.font = Font(name=FONT_NAME, size=10, bold=True)
    c_sno.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B3:B4")
    c_dist = ws.cell(row=3, column=2, value="Distributor Name")
    c_dist.fill = PEACH_FILL
    c_dist.font = Font(name=FONT_NAME, size=10, bold=True)
    c_dist.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C3:C4")
    c_tse = ws.cell(row=3, column=3, value="TSE Range Buiness")
    c_tse.fill = PEACH_FILL
    c_tse.font = Font(name=FONT_NAME, size=10, bold=True)
    c_tse.alignment = Alignment(horizontal="center", vertical="center")

    # Group 1: today TW/Maza Chaika (Target & Ach) -> Cols D to L (4 to 12)
    ws.merge_cells("D3:L3")
    c_today = ws.cell(row=3, column=4, value="today TW/Maza Chaika  (Target & Ach)")
    c_today.fill = TODAY_BLUE_FILL
    c_today.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    c_today.alignment = Alignment(horizontal="center", vertical="center")

    today_sub = [
        (4, "Target  In Tray"), (5, "Target  In Liters"),
        (6, "112ml"), (7, "180ml"), (8, "225ml"),
        (9, "Ach In Tray"), (10, "Ach In Liters"),
        (11, "Ach % Tray"), (12, "Ach % Liters")
    ]
    for col_idx, sub in today_sub:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = TODAY_BLUE_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 2: Yazoo -FM -> Cols M to R (13 to 18)
    ws.merge_cells("M3:R3")
    c_y = ws.cell(row=3, column=13, value="Yazoo -FM")
    c_y.fill = YAZOO_RED_FILL
    c_y.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    c_y.alignment = Alignment(horizontal="center", vertical="center")

    yazoo_sub = [
        (13, "Target in trays"), (14, "112ml"), (15, "180ml"), (16, "225ml"),
        (17, "Ach In Tray"), (18, "Ach %")
    ]
    for col_idx, sub in yazoo_sub:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = YAZOO_RED_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 3: Milker-FM -> Cols S to W (19 to 23) (180ml and 225ml ONLY!)
    ws.merge_cells("S3:W3")
    c_m = ws.cell(row=3, column=19, value="Milker-FM")
    c_m.fill = MILKER_GREEN_FILL
    c_m.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    c_m.alignment = Alignment(horizontal="center", vertical="center")

    milker_sub = [
        (19, "Target  In Tray"), (20, "180ml"), (21, "225ml"),
        (22, "Ach In Tray"), (23, "Ach %")
    ]
    for col_idx, sub in milker_sub:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = MILKER_GREEN_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 4: Maaza -> Cols X to AB (24 to 28)
    ws.merge_cells("X3:AB3")
    c_mz = ws.cell(row=3, column=24, value="Maaza")
    c_mz.fill = MAAZA_ORANGE_FILL
    c_mz.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
    c_mz.alignment = Alignment(horizontal="center", vertical="center")

    maaza_sub = [
        (24, "Target in trays"), (25, "Maaza 330ml"), (26, "Maaza 1000ml"),
        (27, "Ach In Tray"), (28, "Ach %")
    ]
    for col_idx, sub in maaza_sub:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = MAAZA_ORANGE_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Format header borders and heights
    for r in (3, 4):
        ws.row_dimensions[r].height = 24
        for col in range(1, 29):
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

        # today Targets
        tgt_tray = t.get("today_target_tray", t.get("today_target", 0))
        tgt_lit = t.get("today_target_liters", round(tgt_tray * 4.32, 2))
        ws.cell(row=r, column=4, value=tgt_tray).number_format = "#,##0"
        ws.cell(row=r, column=5, value=tgt_lit).number_format = "#,##0.0"

        # today Ach
        a112 = r_ach.get(("today", "112ml"), 0.0)
        a180 = r_ach.get(("today", "180ml"), 0.0)
        a225 = r_ach.get(("today", "225ml"), 0.0)
        ws.cell(row=r, column=6, value=a112).number_format = "#,##0"
        ws.cell(row=r, column=7, value=a180).number_format = "#,##0"
        ws.cell(row=r, column=8, value=a225).number_format = "#,##0"
        ws.cell(row=r, column=9, value=f"=SUM(F{r}:H{r})").number_format = "#,##0"
        ws.cell(row=r, column=10, value=f"=ROUND(F{r}*2.688 + G{r}*4.32 + H{r}*6.075, 2)").number_format = "#,##0.0"
        ws.cell(row=r, column=11, value=f"=IFERROR(I{r}/D{r}, 0)").number_format = "0%"
        ws.cell(row=r, column=12, value=f"=IFERROR(J{r}/E{r}, 0)").number_format = "0%"

        # Yazoo -FM
        y_tgt = t.get("yazoo_target", 0)
        ws.cell(row=r, column=13, value=y_tgt).number_format = "#,##0"
        y112 = r_ach.get(("Yazoo-FM", "112ml"), 0.0)
        y180 = r_ach.get(("Yazoo-FM", "180ml"), 0.0)
        y225 = r_ach.get(("Yazoo-FM", "225ml"), 0.0)
        ws.cell(row=r, column=14, value=y112).number_format = "#,##0"
        ws.cell(row=r, column=15, value=y180).number_format = "#,##0"
        ws.cell(row=r, column=16, value=y225).number_format = "#,##0"
        ws.cell(row=r, column=17, value=f"=SUM(N{r}:P{r})").number_format = "#,##0"
        ws.cell(row=r, column=18, value=f"=IFERROR(Q{r}/M{r}, 0)").number_format = "0%"

        # Milker-FM (180ml and 225ml ONLY!)
        m_tgt = t.get("milker_target", 0)
        ws.cell(row=r, column=19, value=m_tgt).number_format = "#,##0"
        m180 = r_ach.get(("Milker-FM", "180ml"), 0.0)
        m225 = r_ach.get(("Milker-FM", "225ml"), 0.0)
        ws.cell(row=r, column=20, value=m180).number_format = "#,##0"
        ws.cell(row=r, column=21, value=m225).number_format = "#,##0"
        ws.cell(row=r, column=22, value=f"=SUM(T{r}:U{r})").number_format = "#,##0"
        ws.cell(row=r, column=23, value=f"=IFERROR(V{r}/S{r}, 0)").number_format = "0%"

        # Maaza
        mz_tgt = t.get("maaza_target", 0)
        ws.cell(row=r, column=24, value=mz_tgt).number_format = "#,##0"
        mz330 = r_ach.get(("Juices Pet Bottle", "Maaza 330ml"), 0.0)
        mz1000 = r_ach.get(("Juices Pet Bottle", "Maaza 1000 new"), 0.0)
        ws.cell(row=r, column=25, value=mz330).number_format = "#,##0"
        ws.cell(row=r, column=26, value=mz1000).number_format = "#,##0"
        ws.cell(row=r, column=27, value=f"=SUM(Y{r}:Z{r})").number_format = "#,##0"
        ws.cell(row=r, column=28, value=f"=IFERROR(AA{r}/X{r}, 0)").number_format = "0%"

        for c_idx in range(1, 29):
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

    # Formulas for Vaccant row
    ws.cell(row=r_vac, column=4, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", D5:D{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=5, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", E5:E{end_r})').number_format = "#,##0.0"
    ws.cell(row=r_vac, column=6, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", F5:F{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=7, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", G5:G{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=8, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", H5:H{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=9, value=f'=SUM(F{r_vac}:H{r_vac})').number_format = "#,##0"
    ws.cell(row=r_vac, column=10, value=f'=ROUND(F{r_vac}*2.688 + G{r_vac}*4.32 + H{r_vac}*6.075, 2)').number_format = "#,##0.0"
    ws.cell(row=r_vac, column=11, value=f'=IFERROR(I{r_vac}/D{r_vac}, 0)').number_format = "0%"
    ws.cell(row=r_vac, column=12, value=f'=IFERROR(J{r_vac}/E{r_vac}, 0)').number_format = "0%"

    ws.cell(row=r_vac, column=13, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", M5:M{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=14, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", N5:N{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=15, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", O5:O{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=16, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", P5:P{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=17, value=f'=SUM(N{r_vac}:P{r_vac})').number_format = "#,##0"
    ws.cell(row=r_vac, column=18, value=f'=IFERROR(Q{r_vac}/M{r_vac}, 0)').number_format = "0%"

    ws.cell(row=r_vac, column=19, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", S5:S{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=20, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", T5:T{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=21, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", U5:U{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=22, value=f'=SUM(T{r_vac}:U{r_vac})').number_format = "#,##0"
    ws.cell(row=r_vac, column=23, value=f'=IFERROR(V{r_vac}/S{r_vac}, 0)').number_format = "0%"

    ws.cell(row=r_vac, column=24, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", X5:X{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=25, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", Y5:Y{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=26, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", Z5:Z{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=27, value=f'=SUM(Y{r_vac}:Z{r_vac})').number_format = "#,##0"
    ws.cell(row=r_vac, column=28, value=f'=IFERROR(AA{r_vac}/X{r_vac}, 0)').number_format = "0%"

    for c_idx in range(1, 29):
        cell = ws.cell(row=r_vac, column=c_idx)
        cell.fill = YELLOW_FILL
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
        cell.border = BOX_BORDER
        if c_idx > 3:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Row: Grand Total KPK (Sajid Hussain ZSM)
    r_gt = r_vac + 1
    ws.row_dimensions[r_gt].height = 24
    ws.merge_cells(f"A{r_gt}:C{r_gt}")
    lbl_gt = ws.cell(row=r_gt, column=1, value="Grand Total KPK (Sajid Hussain ZSM)")
    lbl_gt.font = Font(name=FONT_NAME, size=11, bold=True)
    lbl_gt.alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=r_gt, column=4, value=f'=SUM(D5:D{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=5, value=f'=SUM(E5:E{end_r})').number_format = "#,##0.0"
    ws.cell(row=r_gt, column=6, value=f'=SUM(F5:F{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=7, value=f'=SUM(G5:G{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=8, value=f'=SUM(H5:H{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=9, value=f'=SUM(F{r_gt}:H{r_gt})').number_format = "#,##0"
    ws.cell(row=r_gt, column=10, value=f'=ROUND(F{r_gt}*2.688 + G{r_gt}*4.32 + H{r_gt}*6.075, 2)').number_format = "#,##0.0"
    ws.cell(row=r_gt, column=11, value=f'=IFERROR(I{r_gt}/D{r_gt}, 0)').number_format = "0%"
    ws.cell(row=r_gt, column=12, value=f'=IFERROR(J{r_gt}/E{r_gt}, 0)').number_format = "0%"

    ws.cell(row=r_gt, column=13, value=f'=SUM(M5:M{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=14, value=f'=SUM(N5:N{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=15, value=f'=SUM(O5:O{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=16, value=f'=SUM(P5:P{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=17, value=f'=SUM(N{r_gt}:P{r_gt})').number_format = "#,##0"
    ws.cell(row=r_gt, column=18, value=f'=IFERROR(Q{r_gt}/M{r_gt}, 0)').number_format = "0%"

    ws.cell(row=r_gt, column=19, value=f'=SUM(S5:S{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=20, value=f'=SUM(T5:T{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=21, value=f'=SUM(U5:U{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=22, value=f'=SUM(T{r_gt}:U{r_gt})').number_format = "#,##0"
    ws.cell(row=r_gt, column=23, value=f'=IFERROR(V{r_gt}/S{r_gt}, 0)').number_format = "0%"

    ws.cell(row=r_gt, column=24, value=f'=SUM(X5:X{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=25, value=f'=SUM(Y5:Y{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=26, value=f'=SUM(Z5:Z{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=27, value=f'=SUM(Y{r_gt}:Z{r_gt})').number_format = "#,##0"
    ws.cell(row=r_gt, column=28, value=f'=IFERROR(AA{r_gt}/X{r_gt}, 0)').number_format = "0%"

    for c_idx in range(1, 29):
        cell = ws.cell(row=r_gt, column=c_idx)
        cell.fill = YELLOW_FILL
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
        cell.border = BOX_BORDER
        if c_idx > 3:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Column Widths
    ws.column_dimensions["A"].width = 6   # S.No
    ws.column_dimensions["B"].width = 44  # Distributor Name
    ws.column_dimensions["C"].width = 18  # TSE Range Buiness
    for col_i in range(4, 29):
        ws.column_dimensions[get_column_letter(col_i)].width = 12

    ws.freeze_panes = "D5"
    ws.sheet_view.showGridLines = True
    return ws


def build_tse_summary_sheet(wb, month_label, targets, dist_ach):
    """
    Build Sheet 3: Compact 16-column TSE Summary sheet matching September reference image.
    """
    ws = wb.create_sheet("TSE Summary (Compact)")

    # Row 1: Title
    ws.merge_cells("A1:M1")
    t = ws.cell(row=1, column=1, value="KPK TARGET VS ACHIEVEMENT  TSE  WISE today, Milker & Yazoo")
    t.font = Font(name=FONT_NAME, size=14, bold=True, color="000000")
    t.fill = BLUE_TITLE_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Row 2: Subtitle
    ws.merge_cells("A2:M2")
    m = ws.cell(row=2, column=1, value=month_label)
    m.font = Font(name=FONT_NAME, size=12, bold=True, color="000000")
    m.fill = BLUE_TITLE_FILL
    m.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Row 3 & 4: Top headers
    ws.merge_cells("A3:A4")
    c_sno = ws.cell(row=3, column=1, value="S.No")
    c_sno.fill = PEACH_FILL
    c_sno.font = Font(name=FONT_NAME, size=10, bold=True)
    c_sno.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B3:B4")
    c_dist = ws.cell(row=3, column=2, value="Distributor Name")
    c_dist.fill = PEACH_FILL
    c_dist.font = Font(name=FONT_NAME, size=10, bold=True)
    c_dist.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C3:C4")
    c_tse = ws.cell(row=3, column=3, value="TSE Range Buiness")
    c_tse.fill = PEACH_FILL
    c_tse.font = Font(name=FONT_NAME, size=10, bold=True)
    c_tse.alignment = Alignment(horizontal="center", vertical="center")

    # Group 1: today: Col D to F (4 to 6)
    ws.merge_cells("D3:F3")
    c_today = ws.cell(row=3, column=4, value="today")
    c_today.fill = TODAY_BLUE_FILL
    c_today.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    c_today.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, sub in [(4, "Target In Tray"), (5, "Ach In Tray"), (6, "Ach % Tray")]:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = TODAY_BLUE_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 2: Milker-FM: Col G to I (7 to 9)
    ws.merge_cells("G3:I3")
    c_m = ws.cell(row=3, column=7, value="Milker-FM")
    c_m.fill = MILKER_GREEN_FILL
    c_m.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    c_m.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, sub in [(7, "Target In Tray"), (8, "Ach In Tray"), (9, "Ach %")]:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = MILKER_GREEN_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Group 3: Maza Chaika: Col J (10)
    c_ch = ws.cell(row=3, column=10, value="Maza Chaika")
    c_ch.fill = PEACH_FILL
    c_ch.font = Font(name=FONT_NAME, size=11, bold=True, color="000000")
    c_ch.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    c_ch_sub = ws.cell(row=4, column=10, value="In tray")
    c_ch_sub.fill = PEACH_FILL
    c_ch_sub.font = Font(name=FONT_NAME, size=9, bold=True, color="000000")
    c_ch_sub.alignment = Alignment(horizontal="center", vertical="center")

    # Group 4: Yazoo: Col K to M (11 to 13)
    ws.merge_cells("K3:M3")
    c_y = ws.cell(row=3, column=11, value="Yazoo")
    c_y.fill = YAZOO_RED_FILL
    c_y.font = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    c_y.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, sub in [(11, "Targ In tray"), (12, "Ach In tray"), (13, "Ach %")]:
        c = ws.cell(row=4, column=col_idx, value=sub)
        c.fill = YAZOO_RED_FILL
        c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r in (3, 4):
        ws.row_dimensions[r].height = 24
        for col in range(1, 14):
            ws.cell(row=r, column=col).border = BOX_BORDER

    start_r = 5
    for idx, t in enumerate(targets):
        r = start_r + idx
        ws.row_dimensions[r].height = 20
        dist_name = t["distributor"]
        tse_name = t.get("tse", "")

        ws.cell(row=r, column=1, value=t.get("sno", idx + 1)).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=2, value=dist_name).alignment = Alignment(horizontal="left", vertical="center")

        c_tse_val = ws.cell(row=r, column=3, value=tse_name)
        c_tse_val.alignment = Alignment(horizontal="center", vertical="center")
        if "vaccant" in tse_name.lower():
            c_tse_val.font = Font(name=FONT_NAME, size=10, bold=True, color="FF0000")
            c_tse_val.fill = PINK_FILL
        else:
            c_tse_val.font = Font(name=FONT_NAME, size=10)

        # today
        ws.cell(row=r, column=4, value=f"='KPK Target vs Achievement'!D{r}").number_format = "#,##0"
        ws.cell(row=r, column=5, value=f"='KPK Target vs Achievement'!I{r}").number_format = "#,##0"
        ws.cell(row=r, column=6, value=f"=IFERROR(E{r}/D{r}, 0)").number_format = "0%"

        # Milker-FM
        ws.cell(row=r, column=7, value=f"='KPK Target vs Achievement'!S{r}").number_format = "#,##0"
        ws.cell(row=r, column=8, value=f"='KPK Target vs Achievement'!V{r}").number_format = "#,##0"
        ws.cell(row=r, column=9, value=f"=IFERROR(H{r}/G{r}, 0)").number_format = "0%"

        # Maza Chaika
        r_ach = dist_ach.get(dist_name, {})
        chaika_ach = r_ach.get(("today", "chaika"), 0.0) + r_ach.get(("Maza Chaika", "180ml"), 0.0)
        ws.cell(row=r, column=10, value=chaika_ach).number_format = "#,##0"

        # Yazoo
        ws.cell(row=r, column=11, value=f"='KPK Target vs Achievement'!M{r}").number_format = "#,##0"
        ws.cell(row=r, column=12, value=f"='KPK Target vs Achievement'!Q{r}").number_format = "#,##0"
        ws.cell(row=r, column=13, value=f"=IFERROR(L{r}/K{r}, 0)").number_format = "0%"

        for c_idx in range(1, 14):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = BOX_BORDER
            if c_idx not in (2, 3):
                cell.font = Font(name=FONT_NAME, size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")

    end_r = start_r + len(targets) - 1

    # Total TSE Vaccant
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
    ws.cell(row=r_vac, column=12, value=f'=SUMIF($C$5:$C${end_r}, "*Vaccant*", L5:L{end_r})').number_format = "#,##0"
    ws.cell(row=r_vac, column=13, value=f'=IFERROR(L{r_vac}/K{r_vac}, 0)').number_format = "0%"

    for c_idx in range(1, 14):
        cell = ws.cell(row=r_vac, column=c_idx)
        cell.fill = YELLOW_FILL
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
        cell.border = BOX_BORDER
        if c_idx > 3:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Grand Total
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
    ws.cell(row=r_gt, column=12, value=f'=SUM(L5:L{end_r})').number_format = "#,##0"
    ws.cell(row=r_gt, column=13, value=f'=IFERROR(L{r_gt}/K{r_gt}, 0)').number_format = "0%"

    for c_idx in range(1, 14):
        cell = ws.cell(row=r_gt, column=c_idx)
        cell.fill = YELLOW_FILL
        cell.font = Font(name=FONT_NAME, size=11, bold=True)
        cell.border = BOX_BORDER
        if c_idx > 3:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = {
        "A": 6, "B": 44, "C": 18, "D": 13, "E": 12, "F": 12,
        "G": 13, "H": 13, "I": 11, "J": 12, "K": 12, "L": 12,
        "M": 11
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    ws.freeze_panes = "D5"
    ws.sheet_view.showGridLines = True
    return ws


def build_category_sheet(wb):
    """Build Sheet 4: Master product catalog."""
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
    """
    Generate the complete KPK Target vs Achievement Excel workbook:
      Sheet 1: Summary (raw dispatches)
      Sheet 2: KPK Target vs Achievement (Master multi-category sheet, NO 200/250ml in Milker, proper colors)
      Sheet 3: TSE Summary (Compact) (16-column TSE Wise sheet matching September reference)
      Sheet 4: Category (Product catalog)
    """
    sku_lookup, party_lookup, targets = load_reference_data()
    dist_ach, unmatched_skus = aggregate_achievements(cumulative_rows, sku_lookup, party_lookup)

    wb = Workbook()
    wb.remove(wb.active)

    build_summary_sheet(wb, cumulative_rows)
    build_kpk_sheet(wb, month_label, targets, dist_ach)
    build_tse_summary_sheet(wb, month_label, targets, dist_ach)
    build_category_sheet(wb)

    wb.save(out_path)
    return out_path, unmatched_skus
