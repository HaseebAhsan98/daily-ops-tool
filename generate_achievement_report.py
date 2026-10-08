"""
HYD Target vs Achievement Report Generator
--------------------------------------------
Builds the exact Hyderabad Target vs Achievement workbook matching the office's
master Excel sheets:

  1. Summary                      -- Cumulative raw dispatch log (Customer :,
                                     SKU NAME, Vehicle#, Bilty#, Sum of Total Tray,
                                     Date, Month, Freight, OB user).
  2. today TW                     -- TSE-wise Target vs Achievement for
                                     "today" (112ml, 180ml, 225ml) + Jam-e-Rahat (800ml),
                                     complete with Tray & Liters targets, achievements,
                                     and live Excel percentage formulas.
  3. today Milker & Juice Pet Bottle -- TSE-wise Target vs Achievement for
                                     Milker-FM (180, 225, 200, 250ml),
                                     Yazoo (112ml),
                                     Juice Pet Bottle (Maza & Polly 250, 500, 1000, 1500, 2000ml),
                                     and MAAZA (330ml, 1000ml new).
  4. Category                     -- Master product catalog with SKU Name, Category, and Size.

Reference & Config files under ach_data/:
  - ach_data/targets.json         -- Distributor targets for each brand/size.
  - ach_data/tse_map.json         -- Distributor to TSE mapping & roster order.
  - ach_data/catgery.json         -- SKU to category & size mapping.
  - ach_data/partname.json        -- Distributor name normalization lookup.
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
DATA_DIR = os.path.join(BASE_DIR, "ach_data")

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
    """Load SKU catalog, party normalization, and distributor targets."""
    with open(os.path.join(DATA_DIR, "catgery.json"), encoding="utf-8") as f:
        catgery = json.load(f)
    with open(os.path.join(DATA_DIR, "partname.json"), encoding="utf-8") as f:
        partname = json.load(f)
    with open(os.path.join(DATA_DIR, "targets.json"), encoding="utf-8") as f:
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
    """Save updated monthly targets to targets.json."""
    path = os.path.join(DATA_DIR, "targets.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(targets, f, indent=2)


def normalize_party(raw_party, party_lookup):
    p = raw_party.strip()
    return party_lookup.get(p.lower(), p)


def extract_date_range_label(raw_text):
    """Automatically extract 'Month of  01 TO D-MON -YYYY' from dispatch header."""
    m = re.search(r"Date:\s*(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})", raw_text)
    if m:
        try:
            d2 = datetime.strptime(m.group(2), "%Y-%m-%d")
            return f"Month of  01 TO {d2.day}-{d2.strftime('%b').upper()} -{d2.year}"
        except Exception:
            pass
    return "Month of  01 TO " + datetime.today().strftime("%d-%b-%Y").upper()


def flatten_dispatch_groups(groups):
    """Convert parsed DO groups into flat tabular rows for Summary and aggregation."""
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
    """Build Sheet 1: Raw dispatch cumulative rows."""
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


def build_today_sheet(wb, month_label, targets, dist_ach):
    """Build Sheet 2: HYD TARGET VS ACHIEVEMENT TSE WISE today TW."""
    ws = wb.create_sheet("today TW")

    # Row 1: Title
    ws.merge_cells("B1:P1")
    t = ws.cell(row=1, column=2, value="HYD TARGET VS ACHIEVEMENT  TSE  WISE today TW")
    t.font = Font(name=FONT_NAME, size=16, bold=True)
    t.fill = BLUE_TITLE_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Row 2: Subtitle
    ws.merge_cells("B2:P2")
    m = ws.cell(row=2, column=2, value=month_label)
    m.font = Font(name=FONT_NAME, size=12, bold=True)
    m.fill = BLUE_TITLE_FILL
    m.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Row 3 & 4: Top headers
    for c_idx, h_text in [(2, "S.No"), (3, "Distributor Name"), (4, "TSE")]:
        ws.merge_cells(start_row=3, start_column=c_idx, end_row=4, end_column=c_idx)
        cell = ws.cell(row=3, column=c_idx, value=h_text)
        cell.font = Font(name=FONT_NAME, size=10, bold=True)
        cell.fill = HDR_PEACH_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("E3:M3")
    c_today = ws.cell(row=3, column=5, value="today")
    c_today.font = Font(name=FONT_NAME, size=11, bold=True)
    c_today.fill = HDR_PEACH_FILL
    c_today.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("N3:P3")
    c_jam = ws.cell(row=3, column=14, value="Jam-e-Rahat")
    c_jam.font = Font(name=FONT_NAME, size=11, bold=True)
    c_jam.fill = HDR_PEACH_FILL
    c_jam.alignment = Alignment(horizontal="center", vertical="center")

    subheaders_today = [
        (5, "Target  In Tray"),
        (6, "Target  In Liters"),
        (7, "112ml"),
        (8, "180ml"),
        (9, "225ml "),
        (10, "Ach In Tray"),
        (11, "Ach In Liters"),
        (12, "Ach % Tray"),
        (13, "Ach % Liters"),
        (14, "Target  In Tray"),
        (15, "800 ml\nIn tray"),
        (16, "Ach % Liters"),
    ]
    for c_idx, sub_text in subheaders_today:
        cell = ws.cell(row=4, column=c_idx, value=sub_text)
        cell.font = Font(name=FONT_NAME, size=9, bold=True)
        cell.fill = HDR_PEACH_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx in (3, 4):
        ws.row_dimensions[r_idx].height = 22 if r_idx == 3 else 30
        for col_idx in range(2, 17):
            ws.cell(row=r_idx, column=col_idx).border = BOX_BORDER
            ws.cell(row=r_idx, column=col_idx).fill = HDR_PEACH_FILL

    curr_row = 5
    tse_subtotal_rows = []

    # Preserve order of TSE groupings
    tse_groups = []
    current_tse = None
    group_items = []
    for t_item in targets:
        tse_name = t_item["tse"]
        if tse_name != current_tse:
            if current_tse is not None:
                tse_groups.append((current_tse, group_items))
            current_tse = tse_name
            group_items = []
        group_items.append(t_item)
    if current_tse is not None:
        tse_groups.append((current_tse, group_items))

    for tse_name, items in tse_groups:
        start_row = curr_row
        for item in items:
            dist_name = item["distributor"]
            r_ach = dist_ach.get(dist_name, {})
            ach_112 = r_ach.get(("today", "112ml"), 0.0)
            ach_180 = r_ach.get(("today", "180ml"), 0.0)
            ach_225 = r_ach.get(("today", "225ml"), 0.0)
            ach_jam = r_ach.get(("JAM-E-RAHAT", "800ml"), 0.0)

            ws.cell(row=curr_row, column=2, value=item["sno"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_row, column=3, value=dist_name).alignment = Alignment(horizontal="left")
            ws.cell(row=curr_row, column=4, value=tse_name).alignment = Alignment(horizontal="center")

            tgt_tray = ws.cell(row=curr_row, column=5, value=item.get("today_target_tray", 0))
            tgt_tray.number_format = "#,##0"

            tgt_lit = ws.cell(row=curr_row, column=6, value=item.get("today_target_liters", 0.0))
            tgt_lit.number_format = "#,##0.00"

            ws.cell(row=curr_row, column=7, value=ach_112).number_format = "#,##0"
            ws.cell(row=curr_row, column=8, value=ach_180).number_format = "#,##0"
            ws.cell(row=curr_row, column=9, value=ach_225).number_format = "#,##0"

            # Formulas:
            # Ach In Tray: =SUM(G{r}:I{r})
            ws.cell(row=curr_row, column=10, value=f"=SUM(G{curr_row}:I{curr_row})").number_format = "#,##0"
            # Ach In Liters: =(G*24*0.112)+(H*24*0.18)+(I*27*0.225)
            ws.cell(row=curr_row, column=11, value=f"=(G{curr_row}*24*0.112)+(H{curr_row}*24*0.18)+(I{curr_row}*27*0.225)").number_format = "#,##0.00"
            # Ach % Tray: =IFERROR(J/E, 0)
            ws.cell(row=curr_row, column=12, value=f"=IFERROR(J{curr_row}/E{curr_row},0)").number_format = "0.0%"
            # Ach % Liters: =IFERROR(K/F, 0)
            ws.cell(row=curr_row, column=13, value=f"=IFERROR(K{curr_row}/F{curr_row},0)").number_format = "0.0%"

            # Jam-e-Rahat
            ws.cell(row=curr_row, column=14, value=item.get("jam_target", 0)).number_format = "#,##0"
            ws.cell(row=curr_row, column=15, value=ach_jam).number_format = "#,##0"
            ws.cell(row=curr_row, column=16, value=f"=IFERROR(O{curr_row}/N{curr_row},0)").number_format = "0.0%"

            for col_idx in range(2, 17):
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

        end_row = curr_row - 1
        # TSE Total Row
        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)
        lbl = ws.cell(row=curr_row, column=2, value=f"Total TSE '{tse_name}")
        lbl.font = Font(name=FONT_NAME, size=10, bold=True)
        lbl.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx in [5, 6, 7, 8, 9, 10, 11]:
            cl = get_column_letter(col_idx)
            ws.cell(row=curr_row, column=col_idx, value=f"=SUM({cl}{start_row}:{cl}{end_row})").number_format = "#,##0" if col_idx not in (6, 11) else "#,##0.00"

        ws.cell(row=curr_row, column=12, value=f"=IFERROR(J{curr_row}/E{curr_row},0)").number_format = "0.0%"
        ws.cell(row=curr_row, column=13, value=f"=IFERROR(K{curr_row}/F{curr_row},0)").number_format = "0.0%"

        for col_idx in [14, 15]:
            cl = get_column_letter(col_idx)
            ws.cell(row=curr_row, column=col_idx, value=f"=SUM({cl}{start_row}:{cl}{end_row})").number_format = "#,##0"
        ws.cell(row=curr_row, column=16, value=f"=IFERROR(O{curr_row}/N{curr_row},0)").number_format = "0.0%"

        for col_idx in range(2, 17):
            c_tot = ws.cell(row=curr_row, column=col_idx)
            c_tot.font = Font(name=FONT_NAME, size=10, bold=True)
            c_tot.fill = TOTAL_YELLOW_FILL
            c_tot.border = BOX_BORDER
            if col_idx not in (2, 3, 4):
                c_tot.alignment = Alignment(horizontal="right", vertical="center")

        ws.row_dimensions[curr_row].height = 20
        tse_subtotal_rows.append(curr_row)
        curr_row += 1

    # Grand Total Row
    gt_row = curr_row
    ws.merge_cells(start_row=gt_row, start_column=2, end_row=gt_row, end_column=4)
    gt_lbl = ws.cell(row=gt_row, column=2, value="Grand Total Waqas Bhutti Sb")
    gt_lbl.font = Font(name=FONT_NAME, size=11, bold=True)
    gt_lbl.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in [5, 6, 7, 8, 9, 10, 11, 14, 15]:
        cl = get_column_letter(col_idx)
        sum_terms = "+".join(f"{cl}{r}" for r in tse_subtotal_rows)
        ws.cell(row=gt_row, column=col_idx, value=f"={sum_terms}").number_format = "#,##0" if col_idx not in (6, 11) else "#,##0.00"

    ws.cell(row=gt_row, column=12, value=f"=IFERROR(J{gt_row}/E{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=13, value=f"=IFERROR(K{gt_row}/F{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=16, value=f"=IFERROR(O{gt_row}/N{gt_row},0)").number_format = "0.0%"

    for col_idx in range(2, 17):
        c_gt = ws.cell(row=gt_row, column=col_idx)
        c_gt.font = Font(name=FONT_NAME, size=11, bold=True)
        c_gt.fill = GRAND_TOTAL_FILL
        c_gt.border = BOX_BORDER
        if col_idx not in (2, 3, 4):
            c_gt.alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[gt_row].height = 22

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 44
    ws.column_dimensions["D"].width = 16
    for c_idx in range(5, 17):
        ws.column_dimensions[get_column_letter(c_idx)].width = 12
    ws.freeze_panes = "E5"
    ws.sheet_view.showGridLines = True
    return ws


def build_milker_juice_sheet(wb, month_label, targets, dist_ach):
    """Build Sheet 3: HYD TARGET VS ACHIEVEMENT TSE WISE today Milker & Juice Pet Bottle."""
    ws = wb.create_sheet("today Milker & Juice Pet Bottle")

    # Row 1: Title
    ws.merge_cells("B1:AF1")
    t = ws.cell(row=1, column=2, value="HYD TARGET VS ACHIEVEMENT  TSE  WISE today Milker & Juice Pet Bottle")
    t.font = Font(name=FONT_NAME, size=16, bold=True)
    t.fill = BLUE_TITLE_FILL
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Row 2: Subtitle
    ws.merge_cells("B2:AF2")
    m = ws.cell(row=2, column=2, value=month_label)
    m.font = Font(name=FONT_NAME, size=12, bold=True)
    m.fill = BLUE_TITLE_FILL
    m.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Row 3 & 4: Top headers
    for c_idx, h_text in [(2, "S.No"), (3, "Distributor Name"), (4, "TSE")]:
        ws.merge_cells(start_row=3, start_column=c_idx, end_row=4, end_column=c_idx)
        cell = ws.cell(row=3, column=c_idx, value=h_text)
        cell.font = Font(name=FONT_NAME, size=10, bold=True)
        cell.fill = HDR_PEACH_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Milker-FM: Col E to K (5 to 11)
    ws.merge_cells("E3:K3")
    c_m = ws.cell(row=3, column=5, value="Milker-FM")
    c_m.font = Font(name=FONT_NAME, size=11, bold=True)
    c_m.fill = HDR_PEACH_FILL
    c_m.alignment = Alignment(horizontal="center", vertical="center")

    # Yazoo: Col L to N (12 to 14)
    ws.merge_cells("L3:N3")
    c_y = ws.cell(row=3, column=12, value="Yazoo")
    c_y.font = Font(name=FONT_NAME, size=11, bold=True)
    c_y.fill = HDR_PEACH_FILL
    c_y.alignment = Alignment(horizontal="center", vertical="center")

    # JuicePet Bottle MAZA & Polly: Col O to AA (15 to 27)
    ws.merge_cells("O3:AA3")
    c_j = ws.cell(row=3, column=15, value="JuicePet Bottle  MAZA & Polly")
    c_j.font = Font(name=FONT_NAME, size=11, bold=True)
    c_j.fill = HDR_PEACH_FILL
    c_j.alignment = Alignment(horizontal="center", vertical="center")

    # MAAZA: Col AB to AF (28 to 32)
    ws.merge_cells("AB3:AF3")
    c_mz = ws.cell(row=3, column=28, value="MAAZA")
    c_mz.font = Font(name=FONT_NAME, size=11, bold=True)
    c_mz.fill = HDR_PEACH_FILL
    c_mz.alignment = Alignment(horizontal="center", vertical="center")

    subheaders = [
        # Milker-FM
        (5, "Target  In Tray"),
        (6, "180ml"),
        (7, "225ml"),
        (8, "200ml"),
        (9, "250ml"),
        (10, "Ach In Tray"),
        (11, "Ach %"),
        # Yazoo
        (12, "Target in trays"),
        (13, "Yazoo 112ml Ach In Tray "),
        (14, "Ach %"),
        # JuicePet Bottle MAZA & Polly
        (15, "Target  In Tray"),
        (16, "Maza 250ml "),
        (17, "Maza 500ml"),
        (18, "Maza 1000ml"),
        (19, "Maza 1500ml"),
        (20, "Maza 2000ml"),
        (21, "Polly 250ml"),
        (22, "Polly 500ml"),
        (23, "Polly 1000ml"),
        (24, "Polly 1500ml"),
        (25, "Polly 2000ml"),
        (26, "Ach In Tray"),
        (27, "Ach %"),
        # MAAZA
        (28, "Target in trays"),
        (29, "Maaza330ml"),
        (30, "Maaza 1000 new"),
        (31, "Ach In Tray"),
        (32, "Ach %"),
    ]

    for c_idx, sub_text in subheaders:
        cell = ws.cell(row=4, column=c_idx, value=sub_text)
        cell.font = Font(name=FONT_NAME, size=9, bold=True)
        cell.fill = HDR_PEACH_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx in (3, 4):
        ws.row_dimensions[r_idx].height = 22 if r_idx == 3 else 30
        for col_idx in range(2, 33):
            ws.cell(row=r_idx, column=col_idx).border = BOX_BORDER
            ws.cell(row=r_idx, column=col_idx).fill = HDR_PEACH_FILL

    curr_row = 5
    tse_subtotal_rows = []

    tse_groups = []
    current_tse = None
    group_items = []
    for t_item in targets:
        tse_name = t_item["tse"]
        if tse_name != current_tse:
            if current_tse is not None:
                tse_groups.append((current_tse, group_items))
            current_tse = tse_name
            group_items = []
        group_items.append(t_item)
    if current_tse is not None:
        tse_groups.append((current_tse, group_items))

    for tse_name, items in tse_groups:
        start_row = curr_row
        for item in items:
            dist_name = item["distributor"]
            r_ach = dist_ach.get(dist_name, {})

            # Milker
            m_180 = r_ach.get(("Milker-FM", "180ml"), 0.0)
            m_225 = r_ach.get(("Milker-FM", "225ml"), 0.0)
            m_200 = r_ach.get(("Milker-FM", "200ml"), 0.0)
            m_250 = r_ach.get(("Milker-FM", "250ml"), 0.0)

            # Yazoo
            y_112 = r_ach.get(("Yazoo-FM", "112ml"), 0.0)

            # Juice Pet Bottle
            j_m250 = r_ach.get(("Juices Pet Bottle", "Maza 250ml"), 0.0)
            j_m500 = r_ach.get(("Juices Pet Bottle", "Maza 500ml"), 0.0)
            j_m1000 = r_ach.get(("Juices Pet Bottle", "Maza 1000ml"), 0.0)
            j_m1500 = r_ach.get(("Juices Pet Bottle", "Maza 1500ml"), 0.0)
            j_m2000 = r_ach.get(("Juices Pet Bottle", "Maza 2000ml"), 0.0)
            j_p250 = r_ach.get(("Juices Pet Bottle", "Polly 250ml"), 0.0)
            j_p500 = r_ach.get(("Juices Pet Bottle", "Polly 500ml"), 0.0)
            j_p1000 = r_ach.get(("Juices Pet Bottle", "Polly 1000ml"), 0.0)
            j_p1500 = r_ach.get(("Juices Pet Bottle", "Polly 1500ml"), 0.0)
            j_p2000 = r_ach.get(("Juices Pet Bottle", "Polly 2000ml"), 0.0)

            # MAAZA
            mz_330 = r_ach.get(("Juices Pet Bottle", "Maaza 330ml"), 0.0)
            mz_1000new = r_ach.get(("Juices Pet Bottle", "Maaza 1000 new"), 0.0)

            ws.cell(row=curr_row, column=2, value=item["sno"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_row, column=3, value=dist_name).alignment = Alignment(horizontal="left")
            ws.cell(row=curr_row, column=4, value=tse_name).alignment = Alignment(horizontal="center")

            # Milker
            ws.cell(row=curr_row, column=5, value=item.get("milker_target", 0)).number_format = "#,##0"
            ws.cell(row=curr_row, column=6, value=m_180).number_format = "#,##0"
            ws.cell(row=curr_row, column=7, value=m_225).number_format = "#,##0"
            ws.cell(row=curr_row, column=8, value=m_200).number_format = "#,##0"
            ws.cell(row=curr_row, column=9, value=m_250).number_format = "#,##0"
            ws.cell(row=curr_row, column=10, value=f"=SUM(F{curr_row}:I{curr_row})").number_format = "#,##0"
            ws.cell(row=curr_row, column=11, value=f"=IFERROR(J{curr_row}/E{curr_row},0)").number_format = "0.0%"

            # Yazoo
            ws.cell(row=curr_row, column=12, value=item.get("yazoo_target", 0)).number_format = "#,##0"
            ws.cell(row=curr_row, column=13, value=y_112).number_format = "#,##0"
            ws.cell(row=curr_row, column=14, value=f"=IFERROR(M{curr_row}/L{curr_row},0)").number_format = "0.0%"

            # JuicePet Bottle
            ws.cell(row=curr_row, column=15, value=item.get("jpb_target", 0)).number_format = "#,##0"
            ws.cell(row=curr_row, column=16, value=j_m250).number_format = "#,##0"
            ws.cell(row=curr_row, column=17, value=j_m500).number_format = "#,##0"
            ws.cell(row=curr_row, column=18, value=j_m1000).number_format = "#,##0"
            ws.cell(row=curr_row, column=19, value=j_m1500).number_format = "#,##0"
            ws.cell(row=curr_row, column=20, value=j_m2000).number_format = "#,##0"
            ws.cell(row=curr_row, column=21, value=j_p250).number_format = "#,##0"
            ws.cell(row=curr_row, column=22, value=j_p500).number_format = "#,##0"
            ws.cell(row=curr_row, column=23, value=j_p1000).number_format = "#,##0"
            ws.cell(row=curr_row, column=24, value=j_p1500).number_format = "#,##0"
            ws.cell(row=curr_row, column=25, value=j_p2000).number_format = "#,##0"
            ws.cell(row=curr_row, column=26, value=f"=SUM(P{curr_row}:Y{curr_row})").number_format = "#,##0"
            ws.cell(row=curr_row, column=27, value=f"=IFERROR(Z{curr_row}/O{curr_row},0)").number_format = "0.0%"

            # MAAZA
            ws.cell(row=curr_row, column=28, value=item.get("maaza_target", 0)).number_format = "#,##0"
            ws.cell(row=curr_row, column=29, value=mz_330).number_format = "#,##0"
            ws.cell(row=curr_row, column=30, value=mz_1000new).number_format = "#,##0"
            ws.cell(row=curr_row, column=31, value=f"=SUM(AC{curr_row}:AD{curr_row})").number_format = "#,##0"
            ws.cell(row=curr_row, column=32, value=f"=IFERROR(AE{curr_row}/AB{curr_row},0)").number_format = "0.0%"

            for col_idx in range(2, 33):
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

        end_row = curr_row - 1
        # TSE Total Row
        ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)
        lbl = ws.cell(row=curr_row, column=2, value=f"Total TSE   '{tse_name}")
        lbl.font = Font(name=FONT_NAME, size=10, bold=True)
        lbl.alignment = Alignment(horizontal="center", vertical="center")

        sum_cols = [5, 6, 7, 8, 9, 10, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 28, 29, 30, 31]
        for col_idx in sum_cols:
            cl = get_column_letter(col_idx)
            ws.cell(row=curr_row, column=col_idx, value=f"=SUM({cl}{start_row}:{cl}{end_row})").number_format = "#,##0"

        ws.cell(row=curr_row, column=11, value=f"=IFERROR(J{curr_row}/E{curr_row},0)").number_format = "0.0%"
        ws.cell(row=curr_row, column=14, value=f"=IFERROR(M{curr_row}/L{curr_row},0)").number_format = "0.0%"
        ws.cell(row=curr_row, column=27, value=f"=IFERROR(Z{curr_row}/O{curr_row},0)").number_format = "0.0%"
        ws.cell(row=curr_row, column=32, value=f"=IFERROR(AE{curr_row}/AB{curr_row},0)").number_format = "0.0%"

        for col_idx in range(2, 33):
            c_tot = ws.cell(row=curr_row, column=col_idx)
            c_tot.font = Font(name=FONT_NAME, size=10, bold=True)
            c_tot.fill = TOTAL_YELLOW_FILL
            c_tot.border = BOX_BORDER
            if col_idx not in (2, 3, 4):
                c_tot.alignment = Alignment(horizontal="right", vertical="center")

        ws.row_dimensions[curr_row].height = 20
        tse_subtotal_rows.append(curr_row)
        curr_row += 1

    # Grand Total Row
    gt_row = curr_row
    ws.merge_cells(start_row=gt_row, start_column=2, end_row=gt_row, end_column=4)
    gt_lbl = ws.cell(row=gt_row, column=2, value="Grand Total WAQAS BHUTTI SB")
    gt_lbl.font = Font(name=FONT_NAME, size=11, bold=True)
    gt_lbl.alignment = Alignment(horizontal="center", vertical="center")

    sum_cols = [5, 6, 7, 8, 9, 10, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 28, 29, 30, 31]
    for col_idx in sum_cols:
        cl = get_column_letter(col_idx)
        sum_terms = "+".join(f"{cl}{r}" for r in tse_subtotal_rows)
        ws.cell(row=gt_row, column=col_idx, value=f"={sum_terms}").number_format = "#,##0"

    ws.cell(row=gt_row, column=11, value=f"=IFERROR(J{gt_row}/E{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=14, value=f"=IFERROR(M{gt_row}/L{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=27, value=f"=IFERROR(Z{gt_row}/O{gt_row},0)").number_format = "0.0%"
    ws.cell(row=gt_row, column=32, value=f"=IFERROR(AE{gt_row}/AB{gt_row},0)").number_format = "0.0%"

    for col_idx in range(2, 33):
        c_gt = ws.cell(row=gt_row, column=col_idx)
        c_gt.font = Font(name=FONT_NAME, size=11, bold=True)
        c_gt.fill = GRAND_TOTAL_FILL
        c_gt.border = BOX_BORDER
        if col_idx not in (2, 3, 4):
            c_gt.alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[gt_row].height = 22

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 44
    ws.column_dimensions["D"].width = 16
    for c_idx in range(5, 33):
        ws.column_dimensions[get_column_letter(c_idx)].width = 11
    ws.freeze_panes = "E5"
    ws.sheet_view.showGridLines = True
    return ws


def build_category_sheet(wb):
    """Build Sheet 4: Master product catalog."""
    with open(os.path.join(DATA_DIR, "catgery.json"), encoding="utf-8") as f:
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
    """Generate the full Target vs Achievement Excel workbook."""
    sku_lookup, party_lookup, targets = load_reference_data()
    dist_ach, unmatched_skus = aggregate_achievements(cumulative_rows, sku_lookup, party_lookup)

    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    build_summary_sheet(wb, cumulative_rows)
    build_today_sheet(wb, month_label, targets, dist_ach)
    build_milker_juice_sheet(wb, month_label, targets, dist_ach)
    build_category_sheet(wb)

    wb.save(out_path)
    return out_path, unmatched_skus


if __name__ == "__main__":
    import sys
    log_path = "achievement_log.json"
    raw_path = sys.argv[1] if len(sys.argv) > 1 else None
    if raw_path and os.path.exists(raw_path):
        with open(raw_path, encoding="utf-8") as f:
            raw_text = f.read()
        month_label = extract_date_range_label(raw_text)
        rows, added, total = set_month_log(log_path, raw_text)
        out_path = f"Achievement_Report_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
        out_path, unmatched = build_workbook(rows, month_label, out_path)
        print(f"Report written to {out_path}, unmatched SKUs: {unmatched}")