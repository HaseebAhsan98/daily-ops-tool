"""
Daily Ops Tool - simple local web app (branded)
------------------------------------------------------------
Run with:  streamlit run app.py
Then open the browser link it shows (usually http://localhost:8501)

Needs these 2 image files in the SAME folder as this script:
    logo.jpg     -> company logo, shown at the top
    banner.jpg   -> background image, shown faded behind the page
"""

import os
import json
import base64
import tempfile
import pandas as pd
import streamlit as st

from generate_attendance_report import load_raw_export, consolidate_attendance, write_report
from generate_dispatch_report import parse_raw_dispatch, build_report as build_dispatch_report
from generate_achievement_report import (
    set_month_log, add_day_to_log, build_workbook as build_achievement_workbook,
    load_reference_data, save_targets, extract_date_range_label, load_cumulative_log,
    aggregate_achievements
)
from generate_kpk_achievement_report import (
    set_month_log as set_kpk_month_log,
    add_day_to_log as add_kpk_day_to_log,
    build_workbook as build_kpk_achievement_workbook,
    load_reference_data as load_kpk_reference_data,
    save_targets as save_kpk_targets,
    extract_date_range_label as extract_kpk_date_range_label,
    load_cumulative_log as load_kpk_cumulative_log,
    aggregate_achievements as aggregate_kpk_achievements
)
from generate_monthly_attendance import load_employees, build_workbook as build_monthly_attendance, CL_ALLOWED
from datetime import datetime as _dt

st.set_page_config(page_title="Popular Group - Ops Tool", page_icon="logo.jpg", layout="wide")


def image_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
banner_path = os.path.join(BASE_DIR, "banner.jpg")
logo_path = os.path.join(BASE_DIR, "logo.jpg")

banner_b64 = image_to_base64(banner_path) if os.path.exists(banner_path) else None
logo_exists = os.path.exists(logo_path)

# ---- Branded background + styling ----
background_css = ""
if banner_b64:
    background_css = f"""
    .stApp {{
        background-image: linear-gradient(rgba(255,255,255,0.93), rgba(255,255,255,0.93)),
                           url("data:image/jpg;base64,{banner_b64}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}
    """

st.markdown(
    f"""
    <style>
    {background_css}
    .block-container {{
        padding-top: 2rem;
    }}
    div[data-testid="stFileUploader"] section {{
        border: 2px dashed #2a5db0;
        border-radius: 10px;
        background-color: rgba(255,255,255,0.7);
    }}
    div.stButton > button, div[data-testid="stDownloadButton"] > button {{
        background-color: #2a5db0;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.2rem;
        font-weight: 600;
    }}
    div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {{
        background-color: #1b4380;
        color: white;
    }}
    h1, h2 {{
        color: #1b4380;
    }}
    [data-testid="stMetricValue"] {{
        color: #2a5db0;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---- Header with logo ----
if logo_exists:
    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_b:
        st.image(logo_path, width=130)

st.markdown(
    "<h1 style='text-align:center;'>Daily Ops Tool</h1>",
    unsafe_allow_html=True,
)

tab_attendance, tab_dispatch, tab_achievement_hyd, tab_achievement_kpk, tab_monthly = st.tabs(
    ["📋 Attendance", "🚚 Dispatch Report", "📊 Target Achievement (HYD)", "📊 Target Achievement (KPK)", "🗓️ Monthly Attendance"]
)

# =====================================================================
# TAB 1: ATTENDANCE
# =====================================================================
with tab_attendance:
    st.markdown(
        "<p style='text-align:center; color:#444;'>Upload the raw Excel file here which exported from Software, processed report will be download from below.</p>",
        unsafe_allow_html=True,
    )
    st.write("")
    uploaded_file = st.file_uploader("Raw excel file yahan daalein (.xls)", type=["xls", "xlsx"], key="attendance_upload")

    if uploaded_file is not None:
        tmp_path = None
        out_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name

            raw = load_raw_export(tmp_path)
            summary = consolidate_attendance(raw)
            dist_counts = raw.groupby("Emp ID")["Distributor"].count()
            summary["Distributors"] = summary["Emp ID"].map(dist_counts)

            report_date = pd.Timestamp.today().strftime("%d-%b-%Y")
            out_path = tmp_path.replace(".xls", "_report.xlsx")
            write_report(summary, out_path, report_date)

            total = len(summary)
            absent_n = int((summary["Final_Status"] == "ABSENT").sum())
            present_n = total - absent_n

            col1, col2, col3 = st.columns(3)
            col1.metric("Total", total)
            col2.metric("Absent", absent_n)
            col3.metric("Present", present_n)

            st.dataframe(
                summary[["Emp ID", "Emp Name", "Designation", "Final_Status", "Check_In", "Distributor_Location"]],
                use_container_width=True,
                hide_index=True,
            )

            with open(out_path, "rb") as f:
                st.download_button(
                    "Download Report",
                    f,
                    file_name=f"Attendance_Report_{report_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="attendance_download",
                )

        except Exception as e:
            st.error(f"Kuch ghalat hua: {e}")
        finally:
            for p in (tmp_path, out_path):
                if p and os.path.exists(p):
                    os.remove(p)

# =====================================================================
# TAB 2: DISPATCH REPORT
# =====================================================================
with tab_dispatch:
    st.markdown(
        "<p style='text-align:center; color:#444;'>ERP se raw dispatch data copy karke neeche paste karein, formatted report neeche se download karein.</p>",
        unsafe_allow_html=True,
    )
    st.write("")

    report_date = st.text_input("Report Date (jaise 22-Sep-2026)", key="dispatch_date")
    raw_text = st.text_area(
        "Raw dispatch data yahan paste karein",
        height=300,
        placeholder="Date: 2026-09-22 to 2026-09-22\nSr.#\tLocation\tOB User:\t...",
        key="dispatch_raw",
    )

    if st.button("Generate Dispatch Report", key="dispatch_generate"):
        if not raw_text.strip():
            st.error("Pehle raw data paste karein.")
        elif not report_date.strip():
            st.error("Report date daalein.")
        else:
            out_path = None
            try:
                groups = parse_raw_dispatch(raw_text)
                if not groups:
                    st.error("Data parse nahi ho saka — format check karein (ERP se copy karte waqt headers/tabs sahi copy hone chahiye).")
                else:
                    out_path = os.path.join(tempfile.gettempdir(), f"Dispatch_Report_{report_date.replace(' ', '_')}.xlsx")
                    build_dispatch_report(groups, report_date, out_path)

                    total_qty = sum(qty for g in groups for _, _, qty in g["items"])
                    col1, col2 = st.columns(2)
                    col1.metric("DO Groups", len(groups))
                    col2.metric("Total Qty", f"{total_qty:,.0f}")

                    with open(out_path, "rb") as f:
                        st.download_button(
                            "Download Dispatch Report",
                            f,
                            file_name=f"Dispatch_Report_{report_date.replace(' ', '_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dispatch_download",
                        )
            except Exception as e:
                st.error(f"Kuch ghalat hua: {e}")
            finally:
                if out_path and os.path.exists(out_path):
                    os.remove(out_path)

# =====================================================================
# TAB 3: TARGET ACHIEVEMENT (HYD)
# =====================================================================
with tab_achievement_hyd:
    st.markdown(
        "<p style='text-align:center; color:#444;'>Hyderabad Target vs Achievement Report: ERP dispatch data paste karein, monthly targets manage karein, aur complete multi-sheet Excel report generate karein.</p>",
        unsafe_allow_html=True,
    )
    st.write("")

    LOG_PATH = os.path.join(BASE_DIR, "achievement_log.json")
    ACH_DATA_DIR = os.path.join(BASE_DIR, "ach_data")
    DEFAULT_OCT_PATH = os.path.join(ACH_DATA_DIR, "targets_october_default.json")

    subtab_ach_report, subtab_ach_targets = st.tabs(
        ["📊 Generate Report", "🎯 Set & Edit Monthly Targets"]
    )

    with subtab_ach_report:
        _, _, all_targets = load_reference_data()

        ach_raw_text = st.text_area(
            "Raw dispatch data yahan paste karein (ERP se copy kiya hua)",
            height=250,
            placeholder="Date: 2026-10-01 to 2026-10-04\nSr.#\tLocation\tOB User:\tOB Date:\tDO#\tDN#\tDN/Dispatch Date\tWorking Area:\tParty Name\tVehicle#\tBilty#\tFreight:\tRemarks\tFlavors\tTotal\nItem Name:\tCase Count\tQty:\n...",
            key="ach_raw",
        )

        detected_lbl = extract_date_range_label(ach_raw_text) if ach_raw_text.strip() else ""
        default_lbl = detected_lbl if detected_lbl else "Month of  01 TO " + _dt.today().strftime("%d-%b-%Y").upper()

        col_mode, col_lbl = st.columns([1, 1])
        with col_mode:
            data_mode = st.radio(
                "Data Mode:",
                [
                    "🔄 Full Month Data (Replaces current month log with pasted data)",
                    "➕ Append Daily Data (Adds new dispatches to existing cumulative log)",
                ],
                key="ach_mode"
            )
        with col_lbl:
            month_label = st.text_input("Month Label (Excel sheet header)", value=default_lbl, key="ach_month_label")

        col_add, col_reset = st.columns([3, 1])
        with col_add:
            add_clicked = st.button("🚀 Generate Achievement Report", key="ach_generate")
        with col_reset:
            reset_clicked = st.button("🗑️ Reset Month Data", key="ach_reset")

        if reset_clicked:
            if os.path.exists(LOG_PATH):
                os.remove(LOG_PATH)
            st.success("Cumulative log clear ho gaya — naya mahina shuru karein.")

        if add_clicked:
            if not ach_raw_text.strip():
                st.error("Pehle raw dispatch data paste karein.")
            elif not month_label.strip():
                st.error("Month label daalein.")
            else:
                out_path = None
                try:
                    if "Full Month Data" in data_mode:
                        rows, added, total_parsed = set_month_log(LOG_PATH, ach_raw_text)
                        st.info(f"Full month data loaded: {len(rows)} rows parsed successfully.")
                    else:
                        rows, added, total_parsed = add_day_to_log(LOG_PATH, ach_raw_text)
                        st.info(f"{added} nayi rows add huin (parsed {total_parsed}). Cumulative total: {len(rows)} rows.")

                    if not rows:
                        st.error("Data parse nahi ho saka — ERP text format check karein.")
                    else:
                        out_path = os.path.join(tempfile.gettempdir(), f"Achievement_Report_{month_label.replace(' ', '_').replace('/', '_')}.xlsx")
                        out_path, unmatched = build_achievement_workbook(rows, month_label, out_path)

                        if unmatched:
                            st.warning(f"Ye SKU names Catgery mapping mein nahi milin: {', '.join(sorted(unmatched))}")

                        total_shipped_trays = sum(float(r.get("qty", 0)) for r in rows)
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Total Items Shipped", len(rows))
                        c2.metric("Total Trays Shipped", f"{total_shipped_trays:,.0f}")
                        c3.metric("Report Date Range", month_label)

                        with open(out_path, "rb") as f:
                            st.download_button(
                                "📥 Download Achievement Report (.xlsx)",
                                f.read(),
                                file_name=f"Achievement_Report_{month_label.replace(' ', '_').replace('/', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="ach_download",
                            )

                        st.write("---")
                        st.subheader("📊 Report Preview")
                        sku_lookup, party_lookup, targets_list = load_reference_data()
                        dist_ach, _ = aggregate_achievements(rows, sku_lookup, party_lookup)

                        prev_tab1, prev_tab2, prev_tab3 = st.tabs(["🥛 today TW Preview", "🧃 Milker, Yazoo & Juice Preview", "📋 Summary Log Preview"])

                        with prev_tab1:
                            preview_today = []
                            for t in targets_list:
                                dist_name = t["distributor"]
                                r_ach = dist_ach.get(dist_name, {})
                                a112 = r_ach.get(("today", "112ml"), 0.0)
                                a180 = r_ach.get(("today", "180ml"), 0.0)
                                a225 = r_ach.get(("today", "225ml"), 0.0)
                                ach_tray = a112 + a180 + a225
                                ach_lit = round(a112 * 2.688 + a180 * 4.32 + a225 * 6.075, 2)
                                tgt_tray = t.get("today_target_tray", 0)
                                tgt_lit = t.get("today_target_liters", 0.0)
                                pct_tray = (ach_tray / tgt_tray * 100) if tgt_tray else 0.0
                                pct_lit = (ach_lit / tgt_lit * 100) if tgt_lit else 0.0
                                jam_ach = r_ach.get(("JAM-E-RAHAT", "800ml"), 0.0)
                                jam_tgt = t.get("jam_target", 0)
                                jam_pct = (jam_ach / jam_tgt * 100) if jam_tgt else 0.0

                                preview_today.append({
                                    "S.No": t["sno"],
                                    "Distributor": dist_name,
                                    "TSE": t["tse"],
                                    "Target Tray": tgt_tray,
                                    "Target Liters": tgt_lit,
                                    "112ml Ach": a112,
                                    "180ml Ach": a180,
                                    "225ml Ach": a225,
                                    "Ach Tray": ach_tray,
                                    "Ach Liters": ach_lit,
                                    "Ach % Tray": f"{pct_tray:.1f}%",
                                    "Ach % Liters": f"{pct_lit:.1f}%",
                                    "Jam Target": jam_tgt,
                                    "Jam Ach": jam_ach,
                                    "Jam %": f"{jam_pct:.1f}%",
                                })
                            st.dataframe(pd.DataFrame(preview_today), use_container_width=True, hide_index=True)

                        with prev_tab2:
                            preview_mj = []
                            for t in targets_list:
                                dist_name = t["distributor"]
                                r_ach = dist_ach.get(dist_name, {})
                                m180 = r_ach.get(("Milker-FM", "180ml"), 0.0)
                                m225 = r_ach.get(("Milker-FM", "225ml"), 0.0)
                                m200 = r_ach.get(("Milker-FM", "200ml"), 0.0)
                                m250 = r_ach.get(("Milker-FM", "250ml"), 0.0)
                                m_ach = m180 + m225 + m200 + m250
                                m_tgt = t.get("milker_target", 0)

                                y112 = r_ach.get(("Yazoo-FM", "112ml"), 0.0)
                                y_tgt = t.get("yazoo_target", 0)

                                j_ach = sum(r_ach.get(("Juices Pet Bottle", sz), 0.0) for sz in [
                                    "Maza 250ml", "Maza 500ml", "Maza 1000ml", "Maza 1500ml", "Maza 2000ml",
                                    "Polly 250ml", "Polly 500ml", "Polly 1000ml", "Polly 1500ml", "Polly 2000ml"
                                ])
                                j_tgt = t.get("jpb_target", 0)

                                mz_ach = r_ach.get(("Juices Pet Bottle", "Maaza 330ml"), 0.0) + r_ach.get(("Juices Pet Bottle", "Maaza 1000 new"), 0.0)
                                mz_tgt = t.get("maaza_target", 0)

                                preview_mj.append({
                                    "S.No": t["sno"],
                                    "Distributor": dist_name,
                                    "TSE": t["tse"],
                                    "Milker Target": m_tgt,
                                    "Milker Ach": m_ach,
                                    "Milker %": f"{(m_ach/m_tgt*100):.1f}%" if m_tgt else "0.0%",
                                    "Yazoo Target": y_tgt,
                                    "Yazoo Ach": y112,
                                    "Yazoo %": f"{(y112/y_tgt*100):.1f}%" if y_tgt else "0.0%",
                                    "Juice Target": j_tgt,
                                    "Juice Ach": j_ach,
                                    "Juice %": f"{(j_ach/j_tgt*100):.1f}%" if j_tgt else "0.0%",
                                    "MAAZA Target": mz_tgt,
                                    "MAAZA Ach": mz_ach,
                                    "MAAZA %": f"{(mz_ach/mz_tgt*100):.1f}%" if mz_tgt else "0.0%",
                                })
                            st.dataframe(pd.DataFrame(preview_mj), use_container_width=True, hide_index=True)

                        with prev_tab3:
                            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                except Exception as e:
                    st.error(f"Kuch ghalat hua: {e}")
                finally:
                    if out_path and os.path.exists(out_path):
                        os.remove(out_path)

    with subtab_ach_targets:
        st.markdown(
            """
            **Monthly Targets Setup:**
            Yahan se aap har TSE aur Distributor ke liye naye mahine ke targets set aur update kar sakte hain.
            TSE ke mutabiq filter karein ya table direct edit karke **💾 Save Target Changes** par click karein.
            """
        )

        _, _, current_targets = load_reference_data()

        tot_today_tray = sum(t.get("today_target_tray", 0) for t in current_targets)
        tot_today_lit = sum(t.get("today_target_liters", 0) for t in current_targets)
        tot_milker = sum(t.get("milker_target", 0) for t in current_targets)
        tot_yazoo = sum(t.get("yazoo_target", 0) for t in current_targets)
        tot_jpb = sum(t.get("jpb_target", 0) for t in current_targets)
        tot_maaza = sum(t.get("maaza_target", 0) for t in current_targets)
        tot_jam = sum(t.get("jam_target", 0) for t in current_targets)

        mc1, mc2, mc3, mc4, mc5, mc6 = st.columns(6)
        mc1.metric("today TW Trays", f"{tot_today_tray:,.0f}", f"{tot_today_lit:,.0f} L")
        mc2.metric("Milker Trays", f"{tot_milker:,.0f}")
        mc3.metric("Yazoo Trays", f"{tot_yazoo:,.0f}")
        mc4.metric("Juice Pet Trays", f"{tot_jpb:,.0f}")
        mc5.metric("MAAZA Trays", f"{tot_maaza:,.0f}")
        mc6.metric("Jam-e-Rahat", f"{tot_jam:,.0f}")

        st.write("")

        tse_options = ["All TSEs"] + sorted(list({t["tse"] for t in current_targets if t.get("tse")}))
        sel_tse = st.selectbox("Filter by TSE:", tse_options, key="target_tse_filter")

        target_df = pd.DataFrame(current_targets)
        display_df = target_df if sel_tse == "All TSEs" else target_df[target_df["tse"] == sel_tse]

        col_config = {
            "sno": st.column_config.NumberColumn("S.No", disabled=True, width="small"),
            "distributor": st.column_config.TextColumn("Distributor Name", disabled=True, width="large"),
            "tse": st.column_config.TextColumn("TSE", disabled=True, width="medium"),
            "today_112": st.column_config.NumberColumn("today 112ml", min_value=0, step=10),
            "today_180": st.column_config.NumberColumn("today 180ml", min_value=0, step=10),
            "today_225": st.column_config.NumberColumn("today 225ml", min_value=0, step=10),
            "jam_target": st.column_config.NumberColumn("Jam-e-Rahat", min_value=0, step=10),
            "milker_target": st.column_config.NumberColumn("Milker Target", min_value=0, step=50),
            "yazoo_target": st.column_config.NumberColumn("Yazoo Target", min_value=0, step=10),
            "jpb_target": st.column_config.NumberColumn("Juice Pet Target", min_value=0, step=50),
            "maaza_target": st.column_config.NumberColumn("MAAZA Target", min_value=0, step=25),
        }
        visible_cols = [
            "sno", "distributor", "tse", "today_112", "today_180", "today_225",
            "jam_target", "milker_target", "yazoo_target", "jpb_target", "maaza_target"
        ]

        edited_df = st.data_editor(
            display_df[visible_cols],
            column_config=col_config,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="targets_data_editor"
        )

        btn_c1, btn_c2 = st.columns([3, 1])
        if btn_c1.button("💾 Save Target Changes", key="save_targets_btn"):
            target_map = {row["distributor"]: row for row in current_targets}
            for _, edited_row in edited_df.iterrows():
                d_name = edited_row["distributor"]
                if d_name in target_map:
                    t_entry = target_map[d_name]
                    t_entry["today_112"] = float(edited_row["today_112"] or 0)
                    t_entry["today_180"] = float(edited_row["today_180"] or 0)
                    t_entry["today_225"] = float(edited_row["today_225"] or 0)
                    t_entry["today_target_tray"] = t_entry["today_112"] + t_entry["today_180"] + t_entry["today_225"]
                    t_entry["today_target_liters"] = round(
                        t_entry["today_112"] * 2.688 + t_entry["today_180"] * 4.32 + t_entry["today_225"] * 6.075, 2
                    )
                    t_entry["jam_target"] = float(edited_row["jam_target"] or 0)
                    t_entry["milker_target"] = float(edited_row["milker_target"] or 0)
                    t_entry["yazoo_target"] = float(edited_row["yazoo_target"] or 0)
                    t_entry["jpb_target"] = float(edited_row["jpb_target"] or 0)
                    t_entry["maaza_target"] = float(edited_row["maaza_target"] or 0)

            updated_list = list(target_map.values())
            save_targets(updated_list)
            st.success("Target values kamiyabi se save ho gaye!")
            st.rerun()

        if btn_c2.button("↺ Reset to October 2026 Defaults", key="reset_oct_defaults"):
            if os.path.exists(DEFAULT_OCT_PATH):
                with open(DEFAULT_OCT_PATH, encoding="utf-8") as f_def:
                    default_targets = json.load(f_def)
                save_targets(default_targets)
                st.success("October 2026 ke default targets restore ho gaye!")
                st.rerun()
            else:
                st.warning("Default targets backup file nahi mili.")


# =====================================================================
# TAB 4: TARGET ACHIEVEMENT (KPK)
# =====================================================================
with tab_achievement_kpk:
    st.markdown(
        "<p style='text-align:center; color:#444;'>KPK Target vs Achievement Report: ERP dispatch data paste karein, monthly targets manage karein, aur complete single-sheet Excel report generate karein.</p>",
        unsafe_allow_html=True,
    )
    st.write("")

    KPK_LOG_PATH = os.path.join(BASE_DIR, "kpk_achievement_log.json")
    KPK_DATA_DIR = os.path.join(BASE_DIR, "kpk_data")
    KPK_DEFAULT_OCT_PATH = os.path.join(KPK_DATA_DIR, "targets_october_default.json")

    subtab_kpk_report, subtab_kpk_targets = st.tabs(
        ["📊 Generate Report", "🎯 Set & Edit Monthly Targets"]
    )

    with subtab_kpk_report:
        _, _, kpk_all_targets = load_kpk_reference_data()

        kpk_raw_text = st.text_area(
            "Raw dispatch data yahan paste karein (ERP se copy kiya hua)",
            height=250,
            placeholder="Date: 2026-10-01 to 2026-10-04\nSr.#\tLocation\tOB User:\tOB Date:\tDO#\tDN#\tDN/Dispatch Date\tWorking Area:\tParty Name\tVehicle#\tBilty#\tFreight:\tRemarks\tFlavors\tTotal\nItem Name:\tCase Count\tQty:\n...",
            key="kpk_ach_raw",
        )

        detected_kpk_lbl = extract_kpk_date_range_label(kpk_raw_text) if kpk_raw_text.strip() else ""
        default_kpk_lbl = detected_kpk_lbl if detected_kpk_lbl else "Month of  01 TO " + _dt.today().strftime("%d-%b-%Y").upper()

        col_kpk_mode, col_kpk_lbl = st.columns([1, 1])
        with col_kpk_mode:
            kpk_data_mode = st.radio(
                "Data Mode:",
                [
                    "🔄 Full Month Data (Replaces current month log with pasted data)",
                    "➕ Append Daily Data (Adds new dispatches to existing cumulative log)",
                ],
                key="kpk_ach_mode"
            )
        with col_kpk_lbl:
            kpk_month_label = st.text_input("Month Label (Excel sheet header)", value=default_kpk_lbl, key="kpk_ach_month_label")

        col_kpk_add, col_kpk_reset = st.columns([3, 1])
        with col_kpk_add:
            kpk_add_clicked = st.button("🚀 Generate KPK Achievement Report", key="kpk_ach_generate")
        with col_kpk_reset:
            kpk_reset_clicked = st.button("🗑️ Reset KPK Month Data", key="kpk_ach_reset")

        if kpk_reset_clicked:
            if os.path.exists(KPK_LOG_PATH):
                os.remove(KPK_LOG_PATH)
            st.success("KPK cumulative log clear ho gaya — naya mahina shuru karein.")

        if kpk_add_clicked:
            if not kpk_raw_text.strip():
                st.error("Pehle raw dispatch data paste karein.")
            elif not kpk_month_label.strip():
                st.error("Month label daalein.")
            else:
                out_path = None
                try:
                    if "Full Month Data" in kpk_data_mode:
                        rows, added, total_parsed = set_kpk_month_log(KPK_LOG_PATH, kpk_raw_text)
                        st.info(f"Full month data loaded: {len(rows)} rows parsed successfully.")
                    else:
                        rows, added, total_parsed = add_kpk_day_to_log(KPK_LOG_PATH, kpk_raw_text)
                        st.info(f"{added} nayi rows add huin (parsed {total_parsed}). Cumulative total: {len(rows)} rows.")

                    if not rows:
                        st.error("Data parse nahi ho saka — ERP text format check karein.")
                    else:
                        out_path = os.path.join(tempfile.gettempdir(), f"KPK_Achievement_Report_{kpk_month_label.replace(' ', '_').replace('/', '_')}.xlsx")
                        out_path, unmatched = build_kpk_achievement_workbook(rows, kpk_month_label, out_path)

                        if unmatched:
                            st.warning(f"Ye SKU names Catgery mapping mein nahi milin: {', '.join(sorted(unmatched))}")

                        total_shipped_trays = sum(float(r.get("qty", 0)) for r in rows)
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Total Items Shipped", len(rows))
                        c2.metric("Total Trays Shipped", f"{total_shipped_trays:,.0f}")
                        c3.metric("Report Date Range", kpk_month_label)

                        with open(out_path, "rb") as f:
                            st.download_button(
                                "📥 Download KPK Achievement Report (.xlsx)",
                                f.read(),
                                file_name=f"KPK_Achievement_Report_{kpk_month_label.replace(' ', '_').replace('/', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="kpk_ach_download",
                            )

                        st.write("---")
                        st.subheader("📊 KPK Report Preview (TSE Wise today, Milker & Juice Pet Bottle)")
                        sku_lookup, party_lookup, targets_list = load_kpk_reference_data()
                        dist_ach, _ = aggregate_kpk_achievements(rows, sku_lookup, party_lookup)

                        preview_rows = []
                        for t in targets_list:
                            dist_name = t["distributor"]
                            tse_name = t.get("tse", "")
                            r_ach = dist_ach.get(dist_name, {})

                            # today
                            today_tgt = float(t.get("today_target", t.get("today_target_tray", 0)))
                            today_ach = sum(r_ach.get(("today", sz), 0.0) for sz in ["112ml", "180ml", "225ml"])
                            today_pct = (today_ach / today_tgt * 100) if today_tgt else 0.0

                            # Milker-FM (180ml & 225ml only!)
                            milker_tgt = float(t.get("milker_target", 0))
                            milker_ach = sum(r_ach.get(("Milker-FM", sz), 0.0) for sz in ["180ml", "225ml"])
                            milker_pct = (milker_ach / milker_tgt * 100) if milker_tgt else 0.0

                            # Maza Chaika
                            chaika_ach = float(r_ach.get(("today", "chaika"), 0.0) + r_ach.get(("Maza Chaika", "180ml"), 0.0))

                            # Yazoo
                            yazoo_tgt = float(t.get("yazoo_target", 0))
                            yazoo_ach = sum(r_ach.get(("Yazoo-FM", sz), 0.0) for sz in ["112ml", "180ml", "225ml"])
                            yazoo_pct = (yazoo_ach / yazoo_tgt * 100) if yazoo_tgt else 0.0

                            preview_rows.append({
                                "S.No": t.get("sno", ""),
                                "Distributor Name": dist_name,
                                "TSE Range Buiness": tse_name,
                                "today Target": f"{today_tgt:,.0f}",
                                "today Ach": f"{today_ach:,.0f}",
                                "today %": f"{today_pct:.0f}%",
                                "Milker Target": f"{milker_tgt:,.0f}",
                                "Milker Ach": f"{milker_ach:,.0f}",
                                "Milker %": f"{milker_pct:.0f}%",
                                "Maza Chaika": f"{chaika_ach:,.0f}",
                                "Yazoo Target": f"{yazoo_tgt:,.0f}",
                                "Yazoo Ach": f"{yazoo_ach:,.0f}",
                                "Yazoo %": f"{yazoo_pct:.0f}%",
                                "_tse": tse_name,
                                "_raw": {
                                    "tt": today_tgt, "ta": today_ach,
                                    "mt": milker_tgt, "ma": milker_ach,
                                    "ca": chaika_ach,
                                    "yt": yazoo_tgt, "ya": yazoo_ach
                                }
                            })

                        # Summary calculations
                        vac_tt = sum(r["_raw"]["tt"] for r in preview_rows if "vaccant" in r["_tse"].lower())
                        vac_ta = sum(r["_raw"]["ta"] for r in preview_rows if "vaccant" in r["_tse"].lower())
                        vac_tp = (vac_ta / vac_tt * 100) if vac_tt else 0.0

                        vac_mt = sum(r["_raw"]["mt"] for r in preview_rows if "vaccant" in r["_tse"].lower())
                        vac_ma = sum(r["_raw"]["ma"] for r in preview_rows if "vaccant" in r["_tse"].lower())
                        vac_mp = (vac_ma / vac_mt * 100) if vac_mt else 0.0

                        vac_ca = sum(r["_raw"]["ca"] for r in preview_rows if "vaccant" in r["_tse"].lower())

                        vac_yt = sum(r["_raw"]["yt"] for r in preview_rows if "vaccant" in r["_tse"].lower())
                        vac_ya = sum(r["_raw"]["ya"] for r in preview_rows if "vaccant" in r["_tse"].lower())
                        vac_yp = (vac_ya / vac_yt * 100) if vac_yt else 0.0

                        # Grand totals
                        gt_tt = sum(r["_raw"]["tt"] for r in preview_rows)
                        gt_ta = sum(r["_raw"]["ta"] for r in preview_rows)
                        gt_tp = (gt_ta / gt_tt * 100) if gt_tt else 0.0

                        gt_mt = sum(r["_raw"]["mt"] for r in preview_rows)
                        gt_ma = sum(r["_raw"]["ma"] for r in preview_rows)
                        gt_mp = (gt_ma / gt_mt * 100) if gt_mt else 0.0

                        gt_ca = sum(r["_raw"]["ca"] for r in preview_rows)

                        gt_yt = sum(r["_raw"]["yt"] for r in preview_rows)
                        gt_ya = sum(r["_raw"]["ya"] for r in preview_rows)
                        gt_yp = (gt_ya / gt_yt * 100) if gt_yt else 0.0

                        display_preview = [
                            {k: v for k, v in r.items() if not k.startswith("_")} for r in preview_rows
                        ]
                        display_preview.append({
                            "S.No": "",
                            "Distributor Name": "Total TSE Vaccant",
                            "TSE Range Buiness": "Vaccant",
                            "today Target": f"{vac_tt:,.0f}",
                            "today Ach": f"{vac_ta:,.0f}",
                            "today %": f"{vac_tp:.0f}%",
                            "Milker Target": f"{vac_mt:,.0f}",
                            "Milker Ach": f"{vac_ma:,.0f}",
                            "Milker %": f"{vac_mp:.0f}%",
                            "Maza Chaika": f"{vac_ca:,.0f}",
                            "Yazoo Target": f"{vac_yt:,.0f}",
                            "Yazoo Ach": f"{vac_ya:,.0f}",
                            "Yazoo %": f"{vac_yp:.0f}%",
                        })
                        display_preview.append({
                            "S.No": "",
                            "Distributor Name": "Grand Total  Sajd Hussain ZSM",
                            "TSE Range Buiness": "Total",
                            "today Target": f"{gt_tt:,.0f}",
                            "today Ach": f"{gt_ta:,.0f}",
                            "today %": f"{gt_tp:.0f}%",
                            "Milker Target": f"{gt_mt:,.0f}",
                            "Milker Ach": f"{gt_ma:,.0f}",
                            "Milker %": f"{gt_mp:.0f}%",
                            "Maza Chaika": f"{gt_ca:,.0f}",
                            "Yazoo Target": f"{gt_yt:,.0f}",
                            "Yazoo Ach": f"{gt_ya:,.0f}",
                            "Yazoo %": f"{gt_yp:.0f}%",
                        })

                        kpk_tab1, kpk_tab2 = st.tabs(["📊 Target vs Achievement Preview", "📋 Summary Log Preview"])
                        with kpk_tab1:
                            st.dataframe(pd.DataFrame(display_preview), use_container_width=True, hide_index=True)
                        with kpk_tab2:
                            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                except Exception as e:
                    st.error(f"Kuch ghalat hua: {e}")
                finally:
                    if out_path and os.path.exists(out_path):
                        os.remove(out_path)

    with subtab_kpk_targets:
        st.markdown(
            """
            **KPK Monthly Targets Setup:**
            Yahan se aap KPK ke tamam distributors ke liye Brand-wise monthly targets set aur update kar sakte hain.
            Table direct edit karein aur **💾 Save KPK Target Changes** par click karein.
            """
        )

        _, _, kpk_current_targets = load_kpk_reference_data()

        tot_kpk_today = sum(t.get("today_target", t.get("today_target_tray", 0)) for t in kpk_current_targets)
        tot_kpk_milker = sum(t.get("milker_target", 0) for t in kpk_current_targets)
        tot_kpk_yazoo = sum(t.get("yazoo_target", 0) for t in kpk_current_targets)

        km1, km2, km3 = st.columns(3)
        km1.metric("today Target Trays", f"{tot_kpk_today:,.0f}")
        km2.metric("Milker-FM Target Trays", f"{tot_kpk_milker:,.0f}")
        km3.metric("Yazoo Target Trays", f"{tot_kpk_yazoo:,.0f}")

        st.write("")

        kpk_target_df = pd.DataFrame(kpk_current_targets)

        kpk_col_config = {
            "sno": st.column_config.NumberColumn("S.No", disabled=True, width="small"),
            "distributor": st.column_config.TextColumn("Distributor Name", disabled=True, width="large"),
            "tse": st.column_config.TextColumn("TSE Range Buiness", width="medium"),
            "today_target": st.column_config.NumberColumn("today Target In Tray", min_value=0, step=25),
            "milker_target": st.column_config.NumberColumn("Milker-FM Target In Tray", min_value=0, step=500),
            "yazoo_target": st.column_config.NumberColumn("Yazoo Target In Tray", min_value=0, step=25),
        }
        kpk_visible_cols = [
            "sno", "distributor", "tse", "today_target", "milker_target", "yazoo_target"
        ]

        # Ensure all columns exist in df
        for col in kpk_visible_cols:
            if col not in kpk_target_df.columns:
                kpk_target_df[col] = 0

        kpk_edited_df = st.data_editor(
            kpk_target_df[kpk_visible_cols],
            column_config=kpk_col_config,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="kpk_targets_data_editor"
        )

        kpk_btn_c1, kpk_btn_c2 = st.columns([2, 1])
        if kpk_btn_c1.button("💾 Save KPK Target Changes", key="save_kpk_targets_btn"):
            kpk_map = {row["distributor"]: row for row in kpk_current_targets}
            for _, edited_row in kpk_edited_df.iterrows():
                d_name = edited_row["distributor"]
                if d_name in kpk_map:
                    t_entry = kpk_map[d_name]
                    t_entry["tse"] = str(edited_row["tse"] or "Vaccant")
                    t_entry["today_target"] = float(edited_row["today_target"] or 0)
                    t_entry["milker_target"] = float(edited_row["milker_target"] or 0)
                    t_entry["yazoo_target"] = float(edited_row["yazoo_target"] or 0)

            updated_kpk_list = list(kpk_map.values())
            save_kpk_targets(updated_kpk_list)
            st.success("KPK Target values kamiyabi se save ho gaye!")
            st.rerun()

        if kpk_btn_c2.button("↺ Reset to October Defaults", key="reset_kpk_oct_defaults"):
            if os.path.exists(KPK_DEFAULT_OCT_PATH):
                with open(KPK_DEFAULT_OCT_PATH, encoding="utf-8") as f_def:
                    default_kpk_targets = json.load(f_def)
                save_kpk_targets(default_kpk_targets)
                st.success("KPK October 2026 ke default targets restore ho gaye!")
                st.rerun()
            else:
                st.warning("KPK default targets backup file nahi mili.")


# =====================================================================
# TAB 5: MONTHLY ATTENDANCE
# =====================================================================
with tab_monthly:
    st.markdown(
        "<p style='text-align:center; color:#444;'>Har employee ka raw attendance data uske naam ke neeche paste karein, phir report generate karein.</p>",
        unsafe_allow_html=True,
    )
    try:
        region, employees = load_employees()
    except FileNotFoundError as e:
        st.error(str(e))
        st.stop()

    col_cl, col_h = st.columns(2)
    cl_allowed = col_cl.number_input("Casual Leaves allowed (pehle itne absent CL banenge)", min_value=0, max_value=10,
                                     value=CL_ALLOWED, step=1, key="ma_cl")
    extra_holidays_text = col_h.text_input("Extra holidays (optional, jaise 2026-09-06, 2026-09-07)", key="ma_extra_h")

    raw_inputs = {}
    for i, emp in enumerate(employees):
        with st.expander(f"{emp['name']}  -  {emp['designation']}  ({emp['base_town']})"):
            raw_inputs[i] = st.text_area(
                "Raw attendance data", height=160, key=f"ma_raw_{i}",
                placeholder="S.NO#\tSTATUS\tDATE\tDAY\t...",
            )

    if st.button("Generate Monthly Attendance Report", key="ma_generate"):
        filled = [i for i, t in raw_inputs.items() if t.strip()]
        if not filled:
            st.error("Kam az kam ek employee ka data paste karein.")
        else:
            out_path = None
            try:
                extra = []
                for token in extra_holidays_text.replace(";", ",").split(","):
                    token = token.strip()
                    if token:
                        extra.append(_dt.strptime(token, "%Y-%m-%d").date())

                out_path = os.path.join(tempfile.gettempdir(), "Monthly_Attendance.xlsx")
                summary, warns, (yr, mo) = build_monthly_attendance(
                    employees, region, raw_inputs, out_path,
                    extra_holidays=extra, cl_allowed=int(cl_allowed))

                st.success(f"Report tayyar: {_dt(yr, mo, 1).strftime('%B %Y')}")
                st.dataframe(
                    pd.DataFrame(summary).rename(columns={"name": "Name", "P": "P.Days", "A": "Absent",
                                                          "CL": "CL", "Holiday": "Holidays", "has_data": "Data mila"}),
                    use_container_width=True, hide_index=True)
                missing = [employees[i]["name"] for i in range(len(employees)) if i not in filled]
                if missing:
                    st.warning("In ka data paste nahi hua, in ki row khali rahegi: " + ", ".join(missing))
                for w in warns:
                    st.warning(w)

                with open(out_path, "rb") as f:
                    st.download_button(
                        "Download Monthly Attendance",
                        f.read(),
                        file_name=f"Monthly_Attendance_{yr}-{mo:02d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="ma_download",
                    )
            except Exception as e:
                st.error(f"Kuch ghalat hua: {e}")
            finally:
                if out_path and os.path.exists(out_path):
                    os.remove(out_path)