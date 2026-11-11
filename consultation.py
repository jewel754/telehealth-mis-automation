"""
consultation_mis_final.py

Enhanced MIS script matching SQL logic from MIS query.docx.
Outputs a single Excel workbook with many sheets.

Usage:
    python consultation_mis_final.py
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date   # ← add ‘date’ here
import pytz
import os
import sys

# ---------------------------
# Editable parameters
# ---------------------------
APPOINTMENT_FILE = "appointment.csv"   # path to CSV file
TIMEZONE = "Asia/Kolkata"
FORCE_TODAY = "2026-11-10"  # Pinned to latest dataset date for consistent Weekly/MTD outputs

# FY boundaries (editable)
FY_START = pd.Timestamp("2025-04-01").date()
FY_END = pd.Timestamp("2026-11-30").date()

EXPORT_FOLDER = "./exports/"
EXPORT_EXCEL = True
EXPORT_EXCEL_FILENAME = "consultation_metrics_final.xlsx"
EXPORT_EXCEL_PATH = os.path.join(EXPORT_FOLDER, EXPORT_EXCEL_FILENAME)

# ---------------------------
# Helpers
# ---------------------------
def get_today(tz_name=TIMEZONE, force_date=None):
    if force_date:
        return pd.to_datetime(force_date).date()
    tz = pytz.timezone(tz_name)
    return datetime.now(tz).date()

TODAY = get_today(TIMEZONE, FORCE_TODAY)

# Weekly: Monday to Sunday of the previous week
def get_last_week_range(today):
    this_monday = today - timedelta(days=today.weekday())
    last_monday = this_monday - timedelta(weeks=1)
    last_sunday = this_monday - timedelta(days=1)
    return last_monday, last_sunday

# Monthly: 1st of month to yesterday
def get_month_to_date_range(today):
    first_of_month = today.replace(day=1)
    end_date = today - timedelta(days=1)
    return first_of_month, end_date

WEEK_START, WEEK_END = get_last_week_range(TODAY)
MONTH_START, MONTH_END = get_month_to_date_range(TODAY)

print(f"Today (tz={TIMEZONE}): {TODAY}")
print(f"Weekly range (last Monday to last Sunday): {WEEK_START} -> {WEEK_END}")
print(f"Monthly range (1st to day-before-today): {MONTH_START} -> {MONTH_END}")

# ---------------------------
# Load data
# ---------------------------
df = pd.read_csv(
    APPOINTMENT_FILE,
    parse_dates=["booked_date"],
    dayfirst=True,  # Indian date format (DD-MM-YYYY)
    low_memory=False
)

# Basic cleanup
df["patient_name"] = df["patient_name"].astype(str).replace({"nan": None})
df["booking_status"] = df["booking_status"].astype(str).str.strip().str.lower()
df["appointment_type"] = df["appointment_type"].astype(str).str.strip().str.lower()
df["consulting_doctor"] = df["consulting_doctor"].astype(str).replace({"nan": None})
if "symptoms" not in df.columns:
    df["symptoms"] = None
if "time" not in df.columns:
    df["time"] = None

# Parse time safely
from datetime import datetime as _dt
def parse_time(val):
    if pd.isna(val) or val is None:
        return pd.NaT
    s = str(val).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M:%S %p"):
        try:
            return _dt.strptime(s, fmt).time()
        except Exception:
            continue
    try:
        return pd.to_datetime(s).time()
    except Exception:
        return pd.NaT

df["time_parsed"] = df["time"].apply(parse_time)

# Create datetime for precise comparisons
df["booked_datetime"] = df["booked_date"] + pd.to_timedelta(
    df["time_parsed"].apply(lambda t: (t.hour*3600 + t.minute*60 + t.second) if pd.notna(t) else 0),
    unit='s'
)

# Helper to ensure date-only columns for comparisons
df["booked_date"] = df["booked_date"].dt.date

# ---------------------------
# ---------------------------
# Helper to export (Tables + Auto-Embedded Excel Charts)
# ---------------------------
def export_all_to_excel(dfs_dict, path=EXPORT_EXCEL_PATH):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, LineChart, Reference

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df_out in dfs_dict.items():
            safe_name = sheet_name[:31]
            out_df = df_out if isinstance(df_out, pd.DataFrame) else pd.DataFrame(df_out)
            out_df.to_excel(writer, sheet_name=safe_name, index=False)

            ws = writer.sheets[safe_name]
            ws.views.sheetView[0].showGridLines = True

            # Styling definitions
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            zebra_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
            thin_border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9")
            )

            # Format Headers & Rows
            for col_idx in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col_idx)
                header_cell = ws.cell(row=1, column=col_idx)
                header_cell.fill = header_fill
                header_cell.font = header_font
                header_cell.alignment = Alignment(horizontal="center", vertical="center")

                max_len = max(
                    [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(1, min(ws.max_row + 1, 200))] + [12]
                )
                ws.column_dimensions[col_letter].width = min(max_len + 4, 35)

                for row_idx in range(2, ws.max_row + 1):
                    c = ws.cell(row=row_idx, column=col_idx)
                    c.border = thin_border
                    if row_idx % 2 == 0:
                        c.fill = zebra_fill

            ws.row_dimensions[1].height = 24
            ws.freeze_panes = "A2"

            # ---------------------------------------------------------
            # Auto-Embed Native Excel Charts on Key Summary Tabs
            # ---------------------------------------------------------
            max_r = ws.max_row
            if max_r > 2:
                # 1. Doctor Conversion Chart (Top 10 Doctors)
                if safe_name == "conversion_grouped_by_doctor":
                    chart = BarChart()
                    chart.type = "col"
                    chart.style = 10
                    chart.title = "Top Doctors: Freeaudio vs. Paid Conversions"
                    chart.y_axis.title = "Patients"
                    chart.x_axis.title = "Doctor"
                    end_r = min(max_r, 11)
                    data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=end_r)
                    cats = Reference(ws, min_col=1, min_row=2, max_row=end_r)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
                    chart.width = 18
                    chart.height = 11
                    ws.add_chart(chart, "G2")

                # 2. Doctor Revenue Chart (MTD)
                elif safe_name == "revenue_doctor_month_m2d":
                    chart = BarChart()
                    chart.type = "col"
                    chart.style = 11
                    chart.title = "Doctor-Wise Paid Revenue (MTD)"
                    chart.y_axis.title = "Revenue (INR)"
                    end_r = min(max_r, 15)
                    data = Reference(ws, min_col=6, max_col=6, min_row=1, max_row=end_r)
                    cats = Reference(ws, min_col=1, min_row=2, max_row=end_r)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
                    chart.width = 18
                    chart.height = 11
                    ws.add_chart(chart, "H2")

                # 3. Symptoms Breakdown Chart (Horizontal Bar)
                elif safe_name == "symptoms_month":
                    chart = BarChart()
                    chart.type = "bar"
                    chart.style = 12
                    chart.title = "Monthly Consultations by Clinical Category"
                    end_r = min(max_r, 12)
                    data = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=end_r)
                    cats = Reference(ws, min_col=1, min_row=2, max_row=end_r)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
                    chart.legend = None
                    chart.width = 18
                    chart.height = 11
                    ws.add_chart(chart, "D2")

                # 4. 24-Hour Consultation Volume Line Chart
                elif safe_name == "hourly_fy_till_yesterday":
                    chart = LineChart()
                    chart.style = 13
                    chart.title = "24-Hour Consultation Traffic Distribution (FY)"
                    chart.y_axis.title = "Completed Consultations"
                    chart.x_axis.title = "Hour of Day (0-23)"
                    data = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=max_r)
                    cats = Reference(ws, min_col=1, min_row=2, max_row=max_r)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
                    chart.legend = None
                    chart.width = 18
                    chart.height = 10
                    ws.add_chart(chart, "D2")

    print(f"Saved formatted Excel with charts: {path}")
    return path

# ---------------------------
# Base prepared frames
# ---------------------------
completed = df[df["booking_status"] == "completed"].copy()

# ---------------------------
# Base prepared frames
# ---------------------------
completed = df[df["booking_status"] == "completed"].copy()
paid_types = ["audio", "video", "chat"]

# ---------------------------
# NEW PATIENTS: FY and WEEKLY
# (first-ever completed consult per patient)
# ---------------------------
first_overall = (
    completed.dropna(subset=["patient_name"])
    .sort_values(["patient_name", "booked_datetime"])
    .groupby("patient_name", as_index=False)
    .first()[["patient_name", "booked_date", "time_parsed", "appointment_type", "booked_datetime"]]
    .rename(columns={"time_parsed": "time"})
)

def new_patients_between(start_date, end_date):
    start = pd.to_datetime(start_date).date() if not isinstance(start_date, date) else start_date
    end = pd.to_datetime(end_date).date() if not isinstance(end_date, date) else end_date
    mask = (first_overall["booked_date"] >= start) & (first_overall["booked_date"] <= end)
    return first_overall.loc[mask].sort_values(["booked_date", "booked_datetime"]).reset_index(drop=True)

new_patients_fy = new_patients_between(FY_START, FY_END)
new_patients_week = new_patients_between(WEEK_START, WEEK_END)

# ---------------------------
# EXISTING PATIENTS (weekly) - as before
# ---------------------------
def existing_patients_for_week(start_date, end_date):
    existing = set(completed[completed["booked_date"] < pd.to_datetime(start_date).date()].patient_name.dropna())
    week_mask = (completed["booked_date"] >= pd.to_datetime(start_date).date()) & (completed["booked_date"] <= pd.to_datetime(end_date).date())
    week_consults = completed.loc[week_mask & completed["patient_name"].isin(existing)]
    out = week_consults.groupby("appointment_type").size().reset_index(name="consultation_count")
    return out.sort_values("appointment_type")

existing_weekly = existing_patients_for_week(WEEK_START, WEEK_END)

# ---------------------------
# SHIFTED FREE → PAID: exact SQL logic replication
# Part 1: Detailed rows per patient (unique first paid after their first freeaudio)
# Requirements from SQL:
#  - First completed freeaudio per patient (distinct on patient_name, earliest datetime)
#  - For each patient, pick the first paid appointment (audio/video/chat) satisfying:
#      a.booked_date > freeaudio_date OR (a.booked_date == freeaudio_date AND a.time > freeaudio_time)
#  - paid_date must be within FY 2025-04-01 to 2026-03-31 (inclusive)
# Output columns: patient_name consulting_doctor freeaudio_date freeaudio_time freeaudio_datetime paid_doctor paid_format paid_date shift_days
# ---------------------------
# first completed freeaudio per patient
freeaudio_completed = completed[completed["appointment_type"] == "freeaudio"].dropna(subset=["patient_name"]).copy()
freeaudio_completed = freeaudio_completed.sort_values(["patient_name", "booked_datetime"])
first_freeaudio = freeaudio_completed.drop_duplicates(subset=["patient_name"], keep="first").copy()
first_freeaudio = first_freeaudio[["patient_name", "consulting_doctor", "booked_date", "time_parsed", "booked_datetime"]]
first_freeaudio = first_freeaudio.rename(columns={
    "consulting_doctor": "freeaudio_doctor",
    "booked_date": "freeaudio_date",
    "time_parsed": "freeaudio_time",
    "booked_datetime": "freeaudio_datetime"
})

# paid appointments (completed) - we'll keep all paid rows and then filter
paid_completed = completed[completed["appointment_type"].isin(paid_types)].copy()
paid_completed = paid_completed[["patient_name", "consulting_doctor", "appointment_type", "booked_date", "time_parsed", "booked_datetime"]]
paid_completed = paid_completed.rename(columns={
    "consulting_doctor": "paid_doctor",
    "appointment_type": "paid_format",
    "booked_date": "paid_date",
    "time_parsed": "paid_time",
    "booked_datetime": "paid_datetime"
})

# Merge and apply SQL-like condition:
conv_merge = pd.merge(first_freeaudio, paid_completed, on="patient_name", how="left")

# Keep only paid rows that happened after freeaudio according to SQL condition:
# (paid_date > freeaudio_date) OR (paid_date == freeaudio_date AND paid_time > freeaudio_time)
def paid_after_freeaudio(row):
    if pd.isna(row["paid_date"]):
        return False
    if row["paid_date"] > row["freeaudio_date"]:
        return True
    if row["paid_date"] == row["freeaudio_date"]:
        # compare times; treat NaT carefully
        if pd.isna(row["paid_time"]) or pd.isna(row["freeaudio_time"]):
            return False
        try:
            return (row["paid_time"] > row["freeaudio_time"])
        except Exception:
            return False
    return False

conv_merge["after_condition"] = conv_merge.apply(paid_after_freeaudio, axis=1)
# Also restrict paid_date to be within FY (>= FY_START and <= FY_END)
conv_merge["paid_in_fy"] = conv_merge["paid_date"].apply(lambda d: (not pd.isna(d)) and (FY_START <= d <= FY_END))
conv_valid = conv_merge.loc[conv_merge["after_condition"] & conv_merge["paid_in_fy"]].copy()

# For each patient pick the earliest paid_datetime (ROW_NUMBER() ... ORDER BY paid_date, paid_time)
conv_valid = conv_valid.sort_values(["patient_name", "paid_datetime"])
conv_first_paid = conv_valid.drop_duplicates(subset=["patient_name"], keep="first").copy()

# shift days
conv_first_paid["shift_days"] = (pd.to_datetime(conv_first_paid["paid_date"]) - pd.to_datetime(conv_first_paid["freeaudio_date"])).dt.days

# Final detailed output with required columns
shifted_detailed = conv_first_paid[[
    "freeaudio_doctor", "patient_name", "freeaudio_date", "paid_doctor", "paid_date", "paid_format", "shift_days",
    "freeaudio_time", "paid_time"
]].copy()

# reorder / rename to match sample precisely
shifted_detailed = shifted_detailed.rename(columns={
    "freeaudio_doctor": "freeaudio_doctor",
    "patient_name": "patient_name",
    "freeaudio_date": "freeaudio_date",
    "freeaudio_time": "freeaudio_time",
    "paid_doctor": "paid_doctor",
    "paid_date": "first_paid_date",
    "paid_format": "paid_format",
    "shift_days": "shift_days",
    "paid_time": "first_paid_time"
})[
    ["freeaudio_doctor", "patient_name", "freeaudio_date", "first_paid_date", "paid_doctor", "paid_format", "shift_days"]
].reset_index(drop=True)

# ---------------------------
# SHIFTED FREE → PAID: Part 2 - doctor summary (counts, % share, avg shift days)
# SQL logic: count freeaudio patients per doctor, count converted (first_paid_date not null), avg(shift_days)
# Exclude rows with 0 conversions (HAVING COUNT(first_paid_date) > 0)
# ---------------------------
# Build conversion_details like SQL: join unique_first_freeaudio with first paid per patient (could be null)
# We'll reuse conv_merge but dedupe first_paid per patient irrespective of FY for counting freeaudio_patient_count.
unique_first_freeaudio = first_freeaudio.copy()

# join to get the first paid AFTER freeaudio irrespective of paid_in_fy for full counts (SQL earlier used paid_date >= '2025-04-01' in some variants;
# but for summary we should match the SQL section in doc: restrict paid_date within FY for counting conversions in FY summary)
conv_for_summary = conv_merge.copy()
conv_for_summary = conv_for_summary.loc[conv_for_summary["after_condition"] & conv_for_summary["paid_in_fy"]].copy()
conv_for_summary = conv_for_summary.sort_values(["patient_name", "paid_datetime"])
conv_for_summary = conv_for_summary.drop_duplicates(subset=["patient_name"], keep="first")

# now aggregation per freeaudio_doctor
conversion_details = pd.merge(
    unique_first_freeaudio.rename(columns={"freeaudio_doctor": "freeaudio_doctor"}),
    conv_for_summary[["patient_name", "paid_date", "paid_format", "paid_doctor", "paid_datetime"]],
    on="patient_name",
    how="left"
)

conversion_details["shift_days"] = (pd.to_datetime(conversion_details["paid_date"]) - pd.to_datetime(conversion_details["freeaudio_date"])).dt.days

conversion_grouped = (
    conversion_details.groupby("freeaudio_doctor", dropna=False)
    .agg(
        freeaudio_patient_count=("patient_name", lambda s: s.nunique()),
        converted_to_paid_count=("paid_date", lambda s: s.notna().sum()),
        avg_shift_days=("shift_days", lambda s: round(s.dropna().mean(), 2) if s.notna().any() else np.nan)
    )
    .reset_index()
)

# % Share
conversion_grouped["%_share"] = (conversion_grouped["converted_to_paid_count"] / conversion_grouped["freeaudio_patient_count"] * 100).round(2)

# keep only rows with converted_to_paid_count > 0 (as SQL HAVING did)
conversion_grouped = conversion_grouped.loc[conversion_grouped["converted_to_paid_count"] > 0].sort_values(by="freeaudio_patient_count", ascending=False).reset_index(drop=True)
conversion_grouped = conversion_grouped.rename(columns={
    "freeaudio_doctor": "Doctors",
    "freeaudio_patient_count": "Freeaudio Patient",
    "converted_to_paid_count": "Paid Patient",
    "%_share": "% Share",
    "avg_shift_days": "Average shift days"
})[["Doctors", "Freeaudio Patient", "Paid Patient", "% Share", "Average shift days"]]

# ---------------------------
# Modified Logic: SHIFTED FREE → PAID conversion groups (prev_fy / this_fy counts)
# SQL:
# prev_fy: freeaudio_date < FY_START and first_paid_date >= FY_START (i.e. paid in this FY)
# this_fy: freeaudio_date >= FY_START and first_paid_date within FY (>=FY_START)
# We'll count unique patients falling in each bucket (based on first paid after freeaudio)
# ---------------------------
# Use conv_first_paid (which already enforced after_condition & paid_in_fy)
conv_first_paid_local = conv_first_paid.copy()
conv_first_paid_local["conversion_group"] = np.where(
    pd.to_datetime(conv_first_paid_local["freeaudio_date"]).dt.date < FY_START,
    "prev_fy",
    "this_fy"
)
conv_counts = conv_first_paid_local.groupby("conversion_group").size().reindex(["prev_fy", "this_fy"], fill_value=0).reset_index()
conv_counts.columns = ["conversion_group", "count"]

# ---------------------------
# MORE THAN 1 consultation
# - FY-to-yesterday (FY_START -> MONTH_END)
# - Weekly separately
# SQL logic: count only 1 freeaudio per patient; other paid counted as they are
# ---------------------------
completed_fy_period = completed[
    (completed["booked_date"] >= FY_START) & (completed["booked_date"] <= MONTH_END)
].copy()

# create flags
# count 1 for freeaudio per patient
fa_once = completed_fy_period[completed_fy_period["appointment_type"] == "freeaudio"].sort_values(["patient_name", "booked_datetime"]).drop_duplicates("patient_name")[["patient_name"]].assign(freeaudio_count=1)
others = completed_fy_period[completed_fy_period["appointment_type"] != "freeaudio"].groupby("patient_name", as_index=False).size().rename(columns={"size": "non_freeaudio_count"})
mt1_fy = pd.merge(fa_once, others, on="patient_name", how="outer").fillna(0)
mt1_fy["total_consultations"] = mt1_fy["freeaudio_count"].fillna(0) + mt1_fy["non_freeaudio_count"].fillna(0)
mt1_fy = mt1_fy[mt1_fy["total_consultations"] > 1].sort_values("total_consultations", ascending=False).reset_index(drop=True)

# Weekly
completed_week_period = completed[
    (completed["booked_date"] >= WEEK_START) & (completed["booked_date"] <= WEEK_END)
].copy()

fa_once_w = completed_week_period[completed_week_period["appointment_type"] == "freeaudio"].sort_values(["patient_name", "booked_datetime"]).drop_duplicates("patient_name")[["patient_name"]].assign(freeaudio_count=1)
others_w = completed_week_period[completed_week_period["appointment_type"] != "freeaudio"].groupby("patient_name", as_index=False).size().rename(columns={"size": "non_freeaudio_count"})
mt1_week = pd.merge(fa_once_w, others_w, on="patient_name", how="outer").fillna(0)
mt1_week["total_consultations"] = mt1_week["freeaudio_count"].fillna(0) + mt1_week["non_freeaudio_count"].fillna(0)
mt1_week = mt1_week[mt1_week["total_consultations"] > 1].sort_values("total_consultations", ascending=False).reset_index(drop=True)

# ---------------------------
# DOCTOR NO-SHOW (counts and time-interval)
# - Weekly and monthly (M2D)
# - two sheets: doctor_no_show_counts_week, doctor_no_show_counts_month
# - and time-interval hourly distributions for both periods
# ---------------------------
no_show_week = df[
    (df["booking_status"] == "doctor-no-show")
    & (df["booked_date"] >= WEEK_START)
    & (df["booked_date"] <= WEEK_END)
].copy()
no_show_week_counts = no_show_week.groupby("consulting_doctor").size().reset_index(name="no_show_count").sort_values("no_show_count", ascending=False)

# hourly distribution (0-23) for doctor no-shows
def hourly_no_show(df_in, start_date, end_date):
    sub = df_in[
        (df_in["booking_status"] == "doctor-no-show")
        & (df_in["booked_date"] >= start_date)
        & (df_in["booked_date"] <= end_date)
    ].copy()
    sub["hour"] = sub["time_parsed"].apply(lambda t: t.hour if pd.notna(t) else np.nan)
    counts = sub.groupby("hour").size().reindex(range(0,24), fill_value=0).reset_index()
    counts.columns = ["hour", "doctor_no_show_count"]
    return counts

no_show_hour_week = hourly_no_show(df, WEEK_START, WEEK_END)
# --- CHANGE HERE ---
# Changed from MONTH_START, MONTH_END to FY_START, MONTH_END
no_show_hour_fy_m2d = hourly_no_show(df, FY_START, MONTH_END)
# --- END CHANGE ---


# doctor no-show month counts (Kept as-is, per request)
no_show_month = df[
    (df["booking_status"] == "doctor-no-show")
    & (df["booked_date"] >= MONTH_START)
    & (df["booked_date"] <= MONTH_END)
].copy()
no_show_month_counts = no_show_month.groupby("consulting_doctor").size().reset_index(name="no_show_count").sort_values("no_show_count", ascending=False)

# ---------------------------
# Keep earlier metrics too (hourly distributions, doctor weekly/monthly consultations, revenue weekly/monthly)
# ---------------------------
def hourly_distribution(df_in, start_date, end_date, status="completed"):
    sub = df_in[
        (df_in["booking_status"] == status)
        & (df_in["booked_date"] >= start_date)
        & (df_in["booked_date"] <= end_date)
    ].copy()
    sub["hour"] = sub["time_parsed"].apply(lambda t: t.hour if pd.notna(t) else np.nan)
    counts = sub.groupby("hour").size().reindex(range(0,24), fill_value=0).reset_index()
    counts.columns = ["hour", "consultations"]
    return counts

hourly_week = hourly_distribution(df, WEEK_START, WEEK_END)
hourly_fy_till_yesterday = hourly_distribution(df, FY_START, MONTH_END)

def doctors_counts(start_date, end_date):
    sub = df[
        (df["booking_status"] == "completed")
        & (df["booked_date"] >= start_date)
        & (df["booked_date"] <= end_date)
    ]
    pivot = (
        sub.groupby(["consulting_doctor", "appointment_type"]).size().unstack(fill_value=0).reset_index()
    )
    for c in ["freeaudio", "audio", "video", "chat"]:
        if c not in pivot.columns:
            pivot[c] = 0
    pivot["total"] = pivot[["freeaudio", "audio", "video", "chat"]].sum(axis=1)
    return pivot

doctor_weekly = doctors_counts(WEEK_START, WEEK_END)
doctor_monthly = doctors_counts(MONTH_START, MONTH_END)

def revenue_by_doctor(start_date, end_date):
    sub = df[
        (df["booking_status"] == "completed")
        & (df["booked_date"] >= start_date)
        & (df["booked_date"] <= end_date)
        & df["patient_name"].notna()
    ].copy()
    sub["amount"] = pd.to_numeric(sub["amount"], errors="coerce").fillna(0)
    pivot = sub.groupby(["consulting_doctor", "appointment_type"])["amount"].sum().unstack(fill_value=0).reset_index()
    for c in ["freeaudio", "audio", "video", "chat"]:
        if c not in pivot.columns:
            pivot[c] = 0
    pivot["total_paid_amount"] = pivot[["audio", "video", "chat"]].sum(axis=1)
    return pivot[["consulting_doctor", "freeaudio", "audio", "video", "chat", "total_paid_amount"]]

rev_doctor_week = revenue_by_doctor(WEEK_START, WEEK_END)
rev_doctor_month = revenue_by_doctor(MONTH_START, MONTH_END)

# ---------------------------
# Symptoms (keeps earlier logic)
# ---------------------------
def symptoms_for_period(start_date, end_date, dedupe=True):
    sub = df[
        (df["booking_status"] == "completed")
        & (df["booked_date"] >= start_date)
        & (df["booked_date"] <= end_date)
    ]
    if dedupe:
        fa = sub[sub["appointment_type"] == "freeaudio"].sort_values(["patient_name", "booked_datetime"]).drop_duplicates("patient_name")[["symptoms"]]
        others = sub[sub["appointment_type"] != "freeaudio"][["symptoms"]]
        all_symptoms = pd.concat([fa, others])
    else:
        all_symptoms = sub[["symptoms"]]
    out = all_symptoms.value_counts().reset_index()
    out.columns = ["symptoms", "consultation_count"]
    return out

symptoms_week = symptoms_for_period(WEEK_START, WEEK_END, dedupe=True)
symptoms_month = symptoms_for_period(MONTH_START, MONTH_END, dedupe=False)

# ---------------------------
# Consolidate outputs and export
# ---------------------------
all_outputs = {
    # New patients
    "new_patients_fy": new_patients_fy,
    "new_patients_week": new_patients_week,

    # Existing
    "existing_weekly": existing_weekly,

    # Shifted detailed + summary
    "shifted_free_to_paid_detailed": shifted_detailed,
    "conversion_grouped_by_doctor": conversion_grouped,
    "conversion_prev_this_counts": conv_counts,

    # More-than-1
    "more_than_one_fy_m2d": mt1_fy,
    "more_than_one_week": mt1_week,

    # Hourly distributions
    "hourly_week": hourly_week,
    "hourly_fy_till_yesterday": hourly_fy_till_yesterday,

    # Doctors consultations
    "doctor_weekly": doctor_weekly,
    "doctor_monthly_m2d": doctor_monthly,

    # Revenue
    "revenue_doctor_week": rev_doctor_week,
    "revenue_doctor_month_m2d": rev_doctor_month,

    # Doctor no-show counts & hourly
    "doctor_no_show_week_counts": no_show_week_counts,
    "doctor_no_show_month_counts": no_show_month_counts, # <-- Kept as-is
    "doctor_no_show_hour_week": no_show_hour_week,
    # --- CHANGE HERE ---
    "doctor_no_show_hour_fy_m2d": no_show_hour_fy_m2d, # <-- Updated from _month to _fy_m2d
    # --- END CHANGE ---

    # Conversion summary (original)
    "conversion_summary": conversion_grouped,

    # Symptoms
    "symptoms_week": symptoms_week,
    "symptoms_month": symptoms_month
}

if EXPORT_EXCEL:
    saved = export_all_to_excel(all_outputs, EXPORT_EXCEL_PATH)

print("\nAll metrics computed successfully ✅")
if EXPORT_EXCEL:
    print(f"Combined Excel saved at: {EXPORT_EXCEL_PATH}")
else:
    print("Set EXPORT_EXCEL = True at top to generate combined Excel file.")