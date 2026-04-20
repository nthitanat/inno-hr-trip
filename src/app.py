"""
app.py — Streamlit web UI for the trip pipeline.

Two modes:
  1. รันทั้งหมดพร้อมกัน — upload once, run all 3 steps, download final workbook
  2. รันทีละขั้นตอน    — run each step individually, inspect and download between steps
"""

import io
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# Ensure imports from this directory work whether run locally or in Docker
sys.path.insert(0, str(Path(__file__).parent))

from pipeline import write_output
from step1_flags import build_trip_flags
from step2_behavior import build_trip_behavior
from step3_cleaned import build_trip_cleaned


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _read_uploaded(uploaded_file, sheet_name: str | None = None) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, encoding="utf-8-sig", header=0)
    kw = {"sheet_name": sheet_name} if sheet_name else {"sheet_name": 0}
    return pd.read_excel(uploaded_file, header=0, **kw)


def _df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


def _build_config() -> dict:
    return {
        "SAME_LOC_M":      float(same_loc_m),
        "BORDERLINE_M":    float(borderline_m),
        "SHORT_STOP_MIN":  float(short_stop_min),
        "NIGHT_HR_START":  int(night_hr_start),
        "NIGHT_HR_END":    int(night_hr_end),
        "NEXTDAY_MIN":     float(nextday_min),
        "COL_BRANCH":      int(col_branch),
        "COL_PERSON":      int(col_person),
        "COL_TITLE":       int(col_title),
        "COL_FNAME":       int(col_fname),
        "COL_LNAME":       int(col_lname),
        "COL_DATETIME":    int(col_datetime),
        "COL_LAT":         int(col_lat),
        "COL_LNG":         int(col_lng),
        "COL_GPS_ACC":     int(col_gps_acc),
        "COL_LOCNAME":     int(col_locname),
    }


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="ระบบวิเคราะห์การเดินทาง",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Global custom styles
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* ── Fonts: IBM Plex Sans + IBM Plex Sans Thai ────────────────────────── */
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600;700&family=IBM+Plex+Sans+Thai:wght@300;400;600;700&display=swap');

    /* ── Design tokens ────────────────────────────────────────────────────── */
    :root {
        --navy:        #123150;
        --navy-dark:   #0C2240;
        --navy-mid:    #1A4070;
        --gold:        #C9A84C;
        --gold-bright: #FFC000;
        --gray-light:  #E8E8E8;
        --gray-mid:    #9CA3AF;
        --white:       #FFFFFF;
    }

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans Thai', 'IBM Plex Sans', sans-serif !important;
    }

    /* ── Page title ───────────────────────────────────────────────────────── */
    h1 {
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: var(--navy) !important;
        letter-spacing: -0.01em;
        border-bottom: 3px solid var(--gold);
        padding-bottom: 0.4rem;
        margin-bottom: 0.2rem !important;
    }

    /* ── Section headings ─────────────────────────────────────────────────── */
    h2, h3 {
        font-weight: 600 !important;
        color: var(--navy) !important;
    }

    /* ── Sidebar — dominant navy surface ─────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background-color: var(--navy) !important;
        border-right: 3px solid var(--gold) !important;
    }
    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] div {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] h2 {
        font-size: 0.72rem !important;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        color: var(--gold) !important;
        border-bottom: 1px solid rgba(201,168,76,0.35);
        padding-bottom: 0.3rem;
        margin-top: 1.2rem !important;
    }
    section[data-testid="stSidebar"] h3 {
        font-size: 0.68rem !important;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: var(--gold-bright) !important;
        margin-top: 1rem !important;
    }
    section[data-testid="stSidebar"] input {
        background: rgba(255,255,255,0.12) !important;
        color: #000000 !important;
        border: 1px solid rgba(201,168,76,0.4) !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"] input:focus {
        border-color: var(--gold-bright) !important;
        box-shadow: 0 0 0 2px rgba(255,192,0,0.2) !important;
    }

    /* ── Primary buttons — gold CTA ───────────────────────────────────────── */
    div.stButton > button[kind="primary"],
    div.stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--gold) 0%, var(--gold-bright) 100%) !important;
        color: var(--navy) !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.5rem 1.4rem !important;
        font-weight: 700 !important;
        font-size: 0.9rem !important;
        letter-spacing: 0.02em;
        box-shadow: 0 2px 8px rgba(201,168,76,0.45) !important;
        transition: box-shadow 0.2s ease, transform 0.1s ease !important;
    }
    div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button[kind="primary"]:hover {
        box-shadow: 0 4px 16px rgba(255,192,0,0.5) !important;
        transform: translateY(-1px) !important;
    }

    /* ── Secondary buttons ────────────────────────────────────────────────── */
    div.stButton > button:not([kind="primary"]) {
        border-radius: 8px !important;
        border: 1.5px solid var(--navy) !important;
        color: var(--navy) !important;
        background: transparent !important;
        font-weight: 600 !important;
        transition: background 0.15s ease, color 0.15s ease !important;
    }
    div.stButton > button:not([kind="primary"]):hover {
        background: var(--navy) !important;
        color: #FFFFFF !important;
    }

    /* ── Metric cards — white with gold top accent ────────────────────────── */
    div[data-testid="stMetric"] {
        background: var(--white) !important;
        border: 1px solid var(--gray-light) !important;
        border-top: 4px solid var(--gold) !important;
        border-radius: 10px !important;
        padding: 1rem 1.2rem !important;
        box-shadow: 0 2px 8px rgba(18,49,80,0.08) !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.74rem !important;
        color: var(--gray-mid) !important;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        font-weight: 600 !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.65rem !important;
        font-weight: 700 !important;
        color: var(--navy) !important;
    }

    /* ── Expanders — navy left border ─────────────────────────────────────── */
    div[data-testid="stExpander"] {
        border: 1px solid var(--gray-light) !important;
        border-left: 4px solid var(--navy) !important;
        border-radius: 8px !important;
        margin-bottom: 0.75rem !important;
        box-shadow: 0 1px 4px rgba(18,49,80,0.07) !important;
        background: var(--white) !important;
    }
    div[data-testid="stExpander"] summary {
        font-weight: 600 !important;
        color: var(--navy) !important;
    }
    div[data-testid="stExpander"] summary:hover {
        color: var(--gold) !important;
    }

    /* ── File uploader — navy dashed border ───────────────────────────────── */
    div[data-testid="stFileUploader"] {
        border: 2px dashed var(--navy) !important;
        border-radius: 10px !important;
        background: rgba(18,49,80,0.03) !important;
        padding: 0.5rem !important;
    }

    /* ── Alerts ───────────────────────────────────────────────────────────── */
    div[data-testid="stAlert"] {
        border-radius: 8px !important;
    }

    /* ── Dataframe ────────────────────────────────────────────────────────── */
    div[data-testid="stDataFrame"] {
        border: 1px solid var(--gray-light) !important;
        border-radius: 8px !important;
        overflow: hidden !important;
    }

    /* ── Tabs — gold active indicator ────────────────────────────────────── */
    div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        border-bottom: 2px solid var(--gray-light) !important;
        gap: 0.25rem;
    }
    button[data-baseweb="tab"] {
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        color: var(--navy) !important;
        border-radius: 6px 6px 0 0 !important;
        padding: 0.5rem 1.2rem !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: var(--gold) !important;
        border-bottom: 3px solid var(--gold) !important;
        background: rgba(201,168,76,0.06) !important;
    }

    /* ── Progress bar — navy→gold gradient ───────────────────────────────── */
    div[data-testid="stProgressBar"] > div > div {
        background: linear-gradient(90deg, var(--navy) 0%, var(--gold-bright) 100%) !important;
    }

    /* ── Number inputs — gold focus ring ─────────────────────────────────── */
    div[data-testid="stNumberInput"] input:focus,
    div[data-testid="stTextInput"] input:focus {
        border-color: var(--gold) !important;
        box-shadow: 0 0 0 2px rgba(201,168,76,0.25) !important;
    }

    /* ── Caption text ─────────────────────────────────────────────────────── */
    .stCaption, div[data-testid="stCaptionContainer"] {
        color: var(--gray-mid) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("ระบบวิเคราะห์การเดินทาง")
st.caption("อัปโหลดไฟล์ข้อมูล Check-in ดิบ ปรับค่าตั้งต้นตามต้องการ แล้วกดประมวลผล")

# ---------------------------------------------------------------------------
# Sidebar — configuration (shared by both modes)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("ตั้งค่าระบบ")

    sheet_name = st.text_input(
        "ชื่อชีตในไฟล์ต้นทาง (Step 1)",
        value="Sheet1",
        help="RAWDATA_SHEET — ชื่อ Sheet ที่เก็บข้อมูล Check-in ดิบ",
    )

    st.subheader("เกณฑ์การตรวจจับการเดินทาง")
    same_loc_m     = st.number_input(
        "ระยะห่างที่ถือว่าอยู่จุดเดิม (เมตร)",
        value=30.0, min_value=0.0,
        help="SAME_LOC_M — ระยะ GPS ที่ถือว่า Check-in อยู่ตำแหน่งเดิม",
    )
    borderline_m   = st.number_input(
        "ระยะเขตกำกวม GPS (เมตร)",
        value=50.0, min_value=0.0,
        help="BORDERLINE_M — ระยะระหว่าง SAME_LOC_M ถึงค่านี้ถือว่าเป็น GPS drift",
    )
    short_stop_min = st.number_input(
        "เวลาหยุดสั้นสุดที่ไม่นับเป็นจุดแวะ (นาที)",
        value=5.0, min_value=0.0,
        help="SHORT_STOP_MIN — หยุดต่ำกว่านี้ถือเป็น SAME_LOC_SHORT",
    )
    nextday_min    = st.number_input(
        "ช่วงเวลาข้ามวันทำงาน (นาที)",
        value=720.0, min_value=0.0,
        help="NEXTDAY_MIN — ช่องว่างเวลา (นาที) ที่ถือว่าเป็นวันทำงานใหม่",
    )
    night_hr_start = st.number_input(
        "เริ่มช่วงกลางคืน (ชั่วโมง 0–23)",
        value=0, min_value=0, max_value=23, step=1,
        help="NIGHT_HR_START — ชั่วโมงเริ่มต้นช่วง Auto-ping กลางคืน",
    )
    night_hr_end   = st.number_input(
        "สิ้นสุดช่วงกลางคืน (ชั่วโมง 0–23)",
        value=5, min_value=0, max_value=23, step=1,
        help="NIGHT_HR_END — ชั่วโมงสิ้นสุดช่วง Auto-ping กลางคืน",
    )

    st.subheader("ตำแหน่งคอลัมน์ในไฟล์ต้นทาง (นับจาก 1)")
    col_branch   = st.number_input(
        "สาขา / หน่วยงาน",      value=2,  min_value=1, step=1, help="COL_BRANCH — คอลัมน์ B")
    col_person   = st.number_input(
        "รหัสพนักงาน",           value=7,  min_value=1, step=1, help="COL_PERSON — คอลัมน์ G")
    col_title    = st.number_input(
        "คำนำหน้าชื่อ",           value=8,  min_value=1, step=1, help="COL_TITLE — คอลัมน์ H")
    col_fname    = st.number_input(
        "ชื่อ",                   value=9,  min_value=1, step=1, help="COL_FNAME — คอลัมน์ I")
    col_lname    = st.number_input(
        "นามสกุล",                value=10, min_value=1, step=1, help="COL_LNAME — คอลัมน์ J")
    col_datetime = st.number_input(
        "วันที่-เวลา Check-in",   value=13, min_value=1, step=1, help="COL_DATETIME — คอลัมน์ M")
    col_lat      = st.number_input(
        "ละติจูด (Latitude)",     value=18, min_value=1, step=1, help="COL_LAT — คอลัมน์ R")
    col_lng      = st.number_input(
        "ลองจิจูด (Longitude)",   value=19, min_value=1, step=1, help="COL_LNG — คอลัมน์ S")
    col_gps_acc  = st.number_input(
        "ความแม่นยำ GPS (เมตร)",  value=20, min_value=1, step=1, help="COL_GPS_ACC — คอลัมน์ T")
    col_locname  = st.number_input(
        "ชื่อสถานที่",             value=21, min_value=1, step=1, help="COL_LOCNAME — คอลัมน์ U")


# ---------------------------------------------------------------------------
# Tabs — choose mode
# ---------------------------------------------------------------------------
tab_full, tab_step = st.tabs(["รันทั้งหมดพร้อมกัน", "รันทีละขั้นตอน"])


# ===========================================================================
# TAB 1 — Full pipeline
# ===========================================================================
with tab_full:
    st.subheader("รันทั้งหมดพร้อมกัน")
    st.write("อัปโหลดไฟล์ข้อมูลดิบ → ประมวลผลทั้ง 3 ขั้นตอน → ดาวน์โหลดผลลัพธ์")

    uploaded = st.file_uploader(
        "วางหรืออัปโหลดไฟล์ข้อมูลที่นี่",
        type=["xlsx", "xlsm", "xls", "csv"],
        help="รองรับไฟล์: .xlsx, .xlsm, .xls, .csv",
        key="full_uploader",
    )

    run_btn = st.button("▶  ประมวลผลทั้งหมด", type="primary", disabled=uploaded is None, key="full_run")

    if uploaded and run_btn:
        config = _build_config()
        try:
            with st.spinner("กำลังอ่านไฟล์ข้อมูล…"):
                rawdata_df = _read_uploaded(uploaded, sheet_name=sheet_name)

            st.info(f"โหลดข้อมูลสำเร็จ: **{len(rawdata_df):,}** แถว × **{len(rawdata_df.columns)}** คอลัมน์")

            progress = st.progress(0, text="กำลังเริ่มต้น…")

            progress.progress(20, text="ขั้นตอนที่ 1 — สร้าง Flag การเดินทาง…")
            flagged_df, trips_df = build_trip_flags(rawdata_df, config)

            progress.progress(55, text="ขั้นตอนที่ 2 — วิเคราะห์พฤติกรรมการเดินทาง…")
            behavior_df = build_trip_behavior(trips_df, config)

            progress.progress(80, text="ขั้นตอนที่ 3 — คัดกรองข้อมูลสะอาด…")
            cleaned_df, summary = build_trip_cleaned(behavior_df, config)

            progress.progress(95, text="กำลังสร้างไฟล์ผลลัพธ์…")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp_path = tmp.name
            write_output(tmp_path, rawdata_df, flagged_df, trips_df, behavior_df, cleaned_df, summary)
            with open(tmp_path, "rb") as f:
                output_bytes = f.read()
            os.unlink(tmp_path)

            progress.progress(100, text="เสร็จสมบูรณ์!")
            st.success("ประมวลผลเสร็จสมบูรณ์!")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("การเดินทางที่ใช้ได้",    summary["total_kept"])
            c2.metric("ระยะทางรวม",             f"{summary['total_distance_km']:.2f} km")
            c3.metric("ระยะทางเฉลี่ย/เที่ยว",   f"{summary['avg_distance_m']:.0f} m")
            c4.metric("การเดินทางที่ถูกตัดออก", summary["total_removed"])

            with st.expander("รายละเอียดการตัดออก"):
                st.write(f"- ลืม Check-out: **{summary['removed_forget']}** รายการ")
                st.write(f"- วันทำงานใหม่:  **{summary['removed_newday']}** รายการ")

            if summary.get("per_person"):
                with st.expander("สรุประยะทางรายบุคคล"):
                    st.dataframe(pd.DataFrame(summary["per_person"]), use_container_width=True)

            stem = Path(uploaded.name).stem
            st.download_button(
                label="⬇  ดาวน์โหลดผลลัพธ์ (.xlsx)",
                data=output_bytes,
                file_name=f"{stem}_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        except Exception as exc:
            st.error(f"เกิดข้อผิดพลาด: {exc}")
            st.exception(exc)


# ===========================================================================
# TAB 2 — Step-by-step
# ===========================================================================
with tab_step:
    st.subheader("รันทีละขั้นตอน")
    st.write(
        "รันและตรวจสอบผลแต่ละขั้นตอนแยกกัน "
        "ดาวน์โหลดไฟล์ผลลัพธ์ระหว่างขั้นตอน แล้วนำมาอัปโหลดต่อในขั้นตอนถัดไป"
    )

    # ── Step 1 ──────────────────────────────────────────────────────────────
    with st.expander("ขั้นตอนที่ 1 — สร้าง Flag การเดินทาง (BuildTripFlags)", expanded=True):
        st.caption(
            "**Input:** ไฟล์ข้อมูลดิบ (.xlsx / .csv)  \n"
            "**Output:** `step1_trips.xlsx` → นำไปอัปโหลดใน Step 2"
        )
        up1 = st.file_uploader(
            "อัปโหลดไฟล์ข้อมูลดิบ",
            type=["xlsx", "xlsm", "xls", "csv"],
            key="step1_uploader",
        )
        run1 = st.button("▶  รัน Step 1", type="primary", disabled=up1 is None, key="step1_run")

        if up1 and run1:
            try:
                with st.spinner("กำลังอ่านไฟล์…"):
                    rawdata_df = _read_uploaded(up1, sheet_name=sheet_name)
                st.info(f"โหลดสำเร็จ: **{len(rawdata_df):,}** แถว × **{len(rawdata_df.columns)}** คอลัมน์")

                with st.spinner("กำลังสร้าง Flag การเดินทาง…"):
                    flagged_df, trips_df = build_trip_flags(rawdata_df, _build_config())

                st.success(
                    f"Step 1 เสร็จสมบูรณ์ — "
                    f"rawdata_flagged: **{len(flagged_df):,}** แถว | trips: **{len(trips_df):,}** แถว"
                )

                if st.checkbox("ดูตัวอย่าง trips (20 แถวแรก)", key="step1_preview"):
                    st.dataframe(trips_df.head(20), use_container_width=True)

                st.download_button(
                    label="⬇  ดาวน์โหลด step1_trips.xlsx (ส่งต่อ Step 2)",
                    data=_df_to_excel_bytes(trips_df, sheet_name="trips"),
                    file_name="step1_trips.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as exc:
                st.error(f"เกิดข้อผิดพลาด: {exc}")
                st.exception(exc)

    # ── Step 2 ──────────────────────────────────────────────────────────────
    with st.expander("ขั้นตอนที่ 2 — วิเคราะห์พฤติกรรมการเดินทาง (BuildTripBehavior)"):
        st.caption(
            "**Input:** `step1_trips.xlsx` จาก Step 1  \n"
            "**Output:** `step2_behavior.xlsx` → นำไปอัปโหลดใน Step 3"
        )
        up2 = st.file_uploader(
            "อัปโหลดไฟล์ผลลัพธ์จาก Step 1 (step1_trips.xlsx)",
            type=["xlsx", "xlsm", "xls"],
            key="step2_uploader",
        )
        run2 = st.button("▶  รัน Step 2", type="primary", disabled=up2 is None, key="step2_run")

        if up2 and run2:
            try:
                with st.spinner("กำลังอ่านไฟล์…"):
                    trips_df = _read_uploaded(up2)
                st.info(f"โหลดสำเร็จ: **{len(trips_df):,}** แถว")

                with st.spinner("กำลังวิเคราะห์พฤติกรรมการเดินทาง…"):
                    behavior_df = build_trip_behavior(trips_df, _build_config())

                st.success(f"Step 2 เสร็จสมบูรณ์ — trip_behavior: **{len(behavior_df):,}** แถว")

                if st.checkbox("ดูตัวอย่าง trip_behavior (20 แถวแรก)", key="step2_preview"):
                    st.dataframe(behavior_df.head(20), use_container_width=True)

                st.download_button(
                    label="⬇  ดาวน์โหลด step2_behavior.xlsx (ส่งต่อ Step 3)",
                    data=_df_to_excel_bytes(behavior_df, sheet_name="trip_behavior"),
                    file_name="step2_behavior.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as exc:
                st.error(f"เกิดข้อผิดพลาด: {exc}")
                st.exception(exc)

    # ── Step 3 ──────────────────────────────────────────────────────────────
    with st.expander("ขั้นตอนที่ 3 — คัดกรองข้อมูลสะอาด (BuildTripCleaned)"):
        st.caption(
            "**Input:** `step2_behavior.xlsx` จาก Step 2  \n"
            "**Output:** `step3_cleaned.xlsx` พร้อมสรุปสถิติ"
        )
        up3 = st.file_uploader(
            "อัปโหลดไฟล์ผลลัพธ์จาก Step 2 (step2_behavior.xlsx)",
            type=["xlsx", "xlsm", "xls"],
            key="step3_uploader",
        )
        run3 = st.button("▶  รัน Step 3", type="primary", disabled=up3 is None, key="step3_run")

        if up3 and run3:
            try:
                with st.spinner("กำลังอ่านไฟล์…"):
                    behavior_df = _read_uploaded(up3)
                st.info(f"โหลดสำเร็จ: **{len(behavior_df):,}** แถว")

                with st.spinner("กำลังคัดกรองข้อมูลสะอาด…"):
                    cleaned_df, summary = build_trip_cleaned(behavior_df, _build_config())

                st.success("Step 3 เสร็จสมบูรณ์!")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("การเดินทางที่ใช้ได้",    summary["total_kept"])
                c2.metric("ระยะทางรวม",             f"{summary['total_distance_km']:.2f} km")
                c3.metric("ระยะทางเฉลี่ย/เที่ยว",   f"{summary['avg_distance_m']:.0f} m")
                c4.metric("การเดินทางที่ถูกตัดออก", summary["total_removed"])

                if st.checkbox("รายละเอียดการตัดออก", key="step3_removal"):
                    st.write(f"- ลืม Check-out: **{summary['removed_forget']}** รายการ")
                    st.write(f"- วันทำงานใหม่:  **{summary['removed_newday']}** รายการ")

                if summary.get("per_person"):
                    if st.checkbox("สรุประยะทางรายบุคคล", key="step3_perperson"):
                        st.dataframe(pd.DataFrame(summary["per_person"]), use_container_width=True)

                if st.checkbox("ดูตัวอย่าง trip_cleaned (20 แถวแรก)", key="step3_preview"):
                    st.dataframe(cleaned_df.head(20), use_container_width=True)

                st.download_button(
                    label="⬇  ดาวน์โหลด step3_cleaned.xlsx",
                    data=_df_to_excel_bytes(cleaned_df, sheet_name="trip_cleaned"),
                    file_name="step3_cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

            except Exception as exc:
                st.error(f"เกิดข้อผิดพลาด: {exc}")
                st.exception(exc)

