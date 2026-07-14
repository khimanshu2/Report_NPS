"""
Call Center + Patient Data Automation Report
==============================================
Upload the two data files (Call Data = df2, Patient/Dressing Data = df)
and get every summary table automatically, plus one combined Excel
download with each table as a separate sheet.
"""

import io
import re
import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(page_title="Call & Patient Report Automation", layout="wide")

# ------------------------------------------------------------------
# ---------------------------- CONSTANTS ----------------------------
# ------------------------------------------------------------------

_DASH_RE = re.compile(r'[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]')


def normalize_dashes(text):
    """Fold any dash-like unicode character down to a plain ASCII hyphen."""
    if not isinstance(text, str):
        return text
    return _DASH_RE.sub('-', text)


NOT_CONNECTED_LIST = [
    'Not Connected', 'Busy / asked to call later', 'DNP 1', 'DNP 2', 'DNP 3',
    'Continue switched off', 'Invalid Number', 'Number not active'
]
POSITIVE_LIST = ['Connected – Feedback Positive', 'Everything is good, no issue at all']
NEGATIVE_LIST = ['Connected – Feedback Negative', 'Overall negative feedback or dissatisfaction']
LANGUAGE_LIST = ['Connected - Language Barrier', 'Telgu', 'Malyalam', 'Kannada', 'Tamil']
EXCLUDED_LIST = POSITIVE_LIST + NEGATIVE_LIST

NOT_CONNECTED_LOWER = {normalize_dashes(x.strip().lower()) for x in NOT_CONNECTED_LIST}
POSITIVE_LOWER = {normalize_dashes(x.strip().lower()) for x in POSITIVE_LIST}
NEGATIVE_LOWER = {normalize_dashes(x.strip().lower()) for x in NEGATIVE_LIST}
LANGUAGE_LOWER = {normalize_dashes(x.strip().lower()) for x in LANGUAGE_LIST}
EXCLUDED_LOWER = POSITIVE_LOWER | NEGATIVE_LOWER

NUMERIC_FEEDBACK_LOWER = {
    normalize_dashes(POSITIVE_LIST[0].strip().lower()),
    normalize_dashes(NEGATIVE_LIST[0].strip().lower()),
}

POSITIVE_NUM_LOWER = normalize_dashes(POSITIVE_LIST[0].strip().lower())
POSITIVE_VERBAL_LOWER = normalize_dashes(POSITIVE_LIST[1].strip().lower())
NEGATIVE_NUM_LOWER = normalize_dashes(NEGATIVE_LIST[0].strip().lower())
NEGATIVE_VERBAL_LOWER = normalize_dashes(NEGATIVE_LIST[1].strip().lower())

MONTH_NAME_MAP = {
    1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June",
    7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December",
}


def normalize_month(series):
    """
    Turns whatever month text a sheet uses ("Apr-26", "April 2026", "01-04-2026",
    "Apr 2026", plain "April", etc.) into ONE canonical label ("Apr-26") + a numeric
    sort key (YYYYMM), so df and df2 always line up on the same Month values for
    filtering — even if one file writes months differently than the other.

    *** FIX: EXCEL EPOCH ARTIFACT ("Dec-1899" etc.) ***
    A Month cell formatted as a Date in Excel but actually blank (or holding a
    stray 0) gets read back as a real Timestamp near Excel's epoch (year 1899
    or 1900). No real call-center record is from 1899, so any date that parses
    to a year before 1990 is this artifact, not a genuine month. Those rows
    are dropped to NaN (not shown as text at all, not even "Unknown") so they
    silently disappear from every filter dropdown, chart, and table instead of
    leaking through as "Dec-1899".
    """
    s = series.astype(str).str.strip()
    parsed = pd.Series(pd.NaT, index=s.index).astype("datetime64[ns]")

    mask = parsed.isna()
    if mask.any():
        parsed.loc[mask] = pd.to_datetime(s[mask], format="%b-%y", errors="coerce")
    mask = parsed.isna()
    if mask.any():
        parsed.loc[mask] = pd.to_datetime(s[mask], format="%b-%Y", errors="coerce")

    mask = parsed.isna()
    if mask.any():
        parsed.loc[mask] = pd.to_datetime(s[mask], format="%B %Y", errors="coerce")
    mask = parsed.isna()
    if mask.any():
        parsed.loc[mask] = pd.to_datetime(s[mask], format="%b %Y", errors="coerce")

    mask = parsed.isna() & s.str.contains(r"\d", regex=True)
    if mask.any():
        try:
            parsed.loc[mask] = pd.to_datetime(s[mask], errors="coerce", dayfirst=True)
        except Exception:
            pass

    # Capture the Excel-epoch rows BEFORE we touch `parsed`, based on the actual
    # parsed year rather than pattern-matching the raw string — this catches
    # the artifact no matter how it was originally formatted (date object,
    # "1899-12-30", "12/30/1899", "Dec-1899", etc).
    epoch_mask = parsed.notna() & (parsed.dt.year < 1990)

    label = parsed.dt.strftime("%b-%y")
    sort_key = (parsed.dt.year * 100 + parsed.dt.month)

    # Genuinely unparseable text (e.g. a bare "May" with no year) falls back
    # to a title-cased version of the raw text.
    unparsed = label.isna() & (~epoch_mask)
    label = label.mask(unparsed, s.str.title())
    sort_key = sort_key.mask(unparsed, 999999).astype(int)

    # Epoch artifacts are dropped entirely (real NaN), never shown as text.
    label = label.mask(epoch_mask, pd.NA)
    sort_key = sort_key.mask(epoch_mask, pd.NA)

    return label, sort_key


# ------------------------------------------------------------------
# --------------------------- HELPERS -------------------------------
# ------------------------------------------------------------------

def clean_text_col(series):
    """Trim spaces + Title case, so 'Ravi ', 'RAVI', 'ravi' all match."""
    return series.astype(str).str.strip().str.title()


def load_file(uploaded_file):
    if uploaded_file is None:
        return None
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def find_col(df, candidates):
    """Find the first matching column name (case-insensitive) from a list of candidates."""
    lower_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower().strip() in lower_map:
            return lower_map[cand.lower().strip()]
    return None


def prepare_df2(df2_raw):
    """Clean & standardize the call data (df2)."""
    df2 = df2_raw.copy()

    col_contact = find_col(df2, ["Contact", "Contact Number"])
    col_subdisp = find_col(df2, ["Sub Disposition"])
    col_calllist = find_col(df2, ["Call List"])
    col_applicator = find_col(df2, ["Applicator", "Applicators"])
    col_city = find_col(df2, ["City"])
    col_nps = find_col(df2, ["NPS SCORE", "NPS Score"])
    col_hospital = find_col(df2, ["Address 1", "Hospital"])
    col_doctor = find_col(df2, ["State", "Doctor"])
    col_month = find_col(df2, ["Month"])

    rename_map = {}
    if col_contact: rename_map[col_contact] = "Contact"
    if col_subdisp: rename_map[col_subdisp] = "Sub Disposition"
    if col_calllist: rename_map[col_calllist] = "Call List"
    if col_applicator: rename_map[col_applicator] = "Applicator"
    if col_city: rename_map[col_city] = "City"
    if col_nps: rename_map[col_nps] = "NPS SCORE"
    if col_hospital: rename_map[col_hospital] = "Hospital Name"
    if col_doctor: rename_map[col_doctor] = "Doctor"
    if col_month: rename_map[col_month] = "Month"

    df2 = df2.rename(columns=rename_map)

    if "Sub Disposition" in df2.columns:
        df2["Sub Disposition"] = (
            df2["Sub Disposition"].astype(str).str.strip()
            .str.replace("â€“", "–", regex=False)
            .str.replace("â€™", "'", regex=False)
        )
        df2["Sub Disposition"] = df2["Sub Disposition"].apply(normalize_dashes)
        blank_tokens = {"nan", "na", "n/a", "none", "null", ""}
        is_blank = df2["Sub Disposition"].str.lower().str.strip().isin(blank_tokens)
        df2.loc[is_blank, "Sub Disposition"] = pd.NA

    if "Applicator" in df2.columns:
        df2["Applicator"] = clean_text_col(df2["Applicator"])

    if "City" in df2.columns:
        df2["City"] = clean_text_col(df2["City"])

    if "Hospital Name" in df2.columns:
        df2["Hospital Name"] = clean_text_col(df2["Hospital Name"])

    if "Doctor" in df2.columns:
        df2["Doctor"] = clean_text_col(df2["Doctor"])

    if "NPS SCORE" in df2.columns:
        df2["NPS SCORE"] = pd.to_numeric(df2["NPS SCORE"], errors="coerce")

    if "Month" in df2.columns:
        month_label, month_sort = normalize_month(df2["Month"])
        df2["Month"] = month_label
        df2["Month Sort"] = month_sort

    return df2


def prepare_df(df_raw):
    """Clean & standardize the patient/dressing data (df)."""
    if df_raw is None:
        return None
    df = df_raw.copy()

    col_applicator = find_col(df, ["Applicators", "Applicator"])
    col_patient = find_col(df, ["Patient Name"])
    col_kit = find_col(df, ["Total Kit"])
    col_hospital = find_col(df, ["Hospital"])
    col_doctor = find_col(df, ["Doctor"])
    col_city = find_col(df, ["City"])
    col_date = find_col(df, ["DATE", "Date"])
    col_month = find_col(df, ["Month"])

    rename_map = {}
    if col_applicator: rename_map[col_applicator] = "Applicators"
    if col_patient: rename_map[col_patient] = "Patient Name"
    if col_kit: rename_map[col_kit] = "Total Kit"
    if col_hospital: rename_map[col_hospital] = "Hospital Name"
    if col_doctor: rename_map[col_doctor] = "Doctor"
    if col_city: rename_map[col_city] = "City"
    if col_date: rename_map[col_date] = "DATE"
    if col_month: rename_map[col_month] = "Month"

    df = df.rename(columns=rename_map)

    if "Applicators" in df.columns:
        df["Applicators"] = clean_text_col(df["Applicators"])
    if "Patient Name" in df.columns:
        df["Patient Name"] = clean_text_col(df["Patient Name"])
    if "Hospital Name" in df.columns:
        df["Hospital Name"] = clean_text_col(df["Hospital Name"])
    if "Doctor" in df.columns:
        df["Doctor"] = clean_text_col(df["Doctor"])
    if "City" in df.columns:
        df["City"] = clean_text_col(df["City"])
    if "Total Kit" in df.columns:
        df["Total Kit"] = pd.to_numeric(df["Total Kit"], errors="coerce").fillna(0)

    if "DATE" in df.columns:
        df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce", dayfirst=True)
        df.loc[df["DATE"].dt.year < 1990, "DATE"] = pd.NaT
        fy_label, fy_sort = zip(*df["DATE"].apply(get_fy_quarter))
        df["FY Quarter"] = fy_label
        df["FY Quarter Sort"] = fy_sort

    if "Month" in df.columns:
        month_label, month_sort = normalize_month(df["Month"])
        df["Month"] = month_label
        df["Month Sort"] = month_sort
    elif "DATE" in df.columns:
        df["Month"] = df["DATE"].dt.strftime("%b-%y")
        df["Month Sort"] = df["DATE"].dt.year * 100 + df["DATE"].dt.month

    return df


def get_fy_quarter(date):
    if pd.isna(date):
        return None, None
    month, year = date.month, date.year
    fy_start = year if month >= 4 else year - 1
    fy_label = f"{fy_start}-{str(fy_start + 1)[-2:]}"
    if month in (4, 5, 6):
        q = 1
    elif month in (7, 8, 9):
        q = 2
    elif month in (10, 11, 12):
        q = 3
    else:
        q = 4
    return f"Q{q} ({fy_label})", fy_start * 10 + q


def get_clean_contact_df2(df2):
    if "Contact" not in df2.columns:
        return df2.copy()
    mask = df2["Contact"].notna() & (df2["Contact"].astype(str).str.strip().str.upper() != "NA")
    out = df2[mask].copy()
    if "Sub Disposition" in out.columns:
        out = out[out["Sub Disposition"].notna()]
    return out


def add_connectivity_status(df2_clean):
    df2_clean = df2_clean.copy()
    df2_clean["Connectivity Status"] = df2_clean["Sub Disposition"].apply(
        lambda x: "Not Connected" if str(x).strip().lower() in NOT_CONNECTED_LOWER else "Connected"
    )
    return df2_clean


def apply_rating_shift(df2_clean):
    df2_clean = df2_clean.copy()
    if "Call List" not in df2_clean.columns:
        return df2_clean

    df2_clean["Effective Call List"] = df2_clean["Call List"]

    rated_mask = pd.Series(False, index=df2_clean.index)

    if "NPS SCORE" in df2_clean.columns:
        rated_mask = rated_mask | (df2_clean["NPS SCORE"] > 0)

    if "Sub Disposition" in df2_clean.columns:
        sub_lower = df2_clean["Sub Disposition"].astype(str).str.strip().str.lower()
        rated_mask = rated_mask | sub_lower.isin(NUMERIC_FEEDBACK_LOWER)

    df2_clean.loc[rated_mask, "Effective Call List"] = "Removal"
    return df2_clean


# ------------------------------------------------------------------
# ------------------------- REPORT BUILDERS --------------------------
# ------------------------------------------------------------------

def build_city_summary(df2_raw, df2_clean, df):
    missing_contact = df2_raw[df2_raw["Contact"].isna()].groupby("City").size().reset_index(name="Missing Contact")

    nc_data = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)]
    nc_total = nc_data.groupby("City").size().reset_index(name="Total Not Connected")

    conn_data = df2_clean[~df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)]
    conn_total = conn_data.groupby("City").size().reset_index(name="Total Connected")

    pos_data = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(POSITIVE_LOWER)]
    pos_total = pos_data.groupby("City").size().reset_index(name="Positive Feedback")

    neg_data = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(NEGATIVE_LOWER)]
    neg_total = neg_data.groupby("City").size().reset_index(name="Negative Feedback")

    gen_data = df2_clean[
        (~df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)) &
        (~df2_clean["Sub Disposition"].str.lower().isin(EXCLUDED_LOWER))
    ]
    gen_total = gen_data.groupby("City").size().reset_index(name="General Query")

    nps_given = df2_clean[df2_clean["NPS SCORE"] > 0].groupby("City").size().reset_index(name="NPS Given")
    avg_nps = df2_clean[df2_clean["NPS SCORE"] > 0].groupby("City")["NPS SCORE"].mean().round(2).reset_index(name="Avg NPS Score")

    table = conn_total.merge(nc_total, on="City", how="outer") \
        .merge(missing_contact, on="City", how="left") \
        .merge(pos_total, on="City", how="left") \
        .merge(neg_total, on="City", how="left") \
        .merge(gen_total, on="City", how="left") \
        .merge(nps_given, on="City", how="left") \
        .merge(avg_nps, on="City", how="left")

    if "Effective Call List" in df2_clean.columns:
        running_total = df2_clean[df2_clean["Effective Call List"] == "Running"].groupby("City").size().reset_index(name="Running")
        removal_total = df2_clean[df2_clean["Effective Call List"] == "Removal"].groupby("City").size().reset_index(name="Removal")
        table = table.merge(running_total, on="City", how="left")
        table = table.merge(removal_total, on="City", how="left")

    fill_cols = ["Total Connected", "Total Not Connected", "Missing Contact",
                 "Positive Feedback", "Negative Feedback", "General Query", "NPS Given",
                 "Running", "Removal"]
    for c in fill_cols:
        if c in table.columns:
            table[c] = table[c].fillna(0).astype(int)
    table["Avg NPS Score"] = table["Avg NPS Score"].fillna(0)

    if df is not None and "City" in df.columns and "Patient Name" in df.columns:
        unique_patients = df.groupby("City")["Patient Name"].nunique().reset_index(name="Unique Patients")
        table = table.merge(unique_patients, on="City", how="outer")
        table["Unique Patients"] = table["Unique Patients"].fillna(0).astype(int)

    table = table.sort_values(by="Total Connected", ascending=False).reset_index(drop=True)

    grand_total = {
        "City": "Grand Total",
        "Total Connected": conn_data.shape[0],
        "Total Not Connected": nc_data.shape[0],
        "Missing Contact": df2_raw["Contact"].isna().sum(),
        "Positive Feedback": pos_data.shape[0],
        "Negative Feedback": neg_data.shape[0],
        "General Query": gen_data.shape[0],
        "NPS Given": (df2_clean["NPS SCORE"] > 0).sum(),
        "Avg NPS Score": round(df2_clean.loc[df2_clean["NPS SCORE"] > 0, "NPS SCORE"].mean(), 2)
        if (df2_clean["NPS SCORE"] > 0).any() else 0,
    }
    if "Effective Call List" in df2_clean.columns:
        grand_total["Running"] = (df2_clean["Effective Call List"] == "Running").sum()
        grand_total["Removal"] = (df2_clean["Effective Call List"] == "Removal").sum()
    if df is not None and "Patient Name" in df.columns:
        grand_total["Unique Patients"] = df["Patient Name"].nunique()

    table = pd.concat([table, pd.DataFrame([grand_total])], ignore_index=True)
    return table


def build_table2_city_nps(df2_raw, df2_clean, df):
    base = build_city_summary(df2_raw, df2_clean, df)
    cols = ["City", "Unique Patients", "Missing Contact", "Total Not Connected",
            "Total Connected", "Positive Feedback", "Negative Feedback", "General Query"]
    cols = [c for c in cols if c in base.columns]
    table = base[cols].copy()

    is_grand_total = table["City"] == "Grand Total"
    body = table[~is_grand_total].sort_values(by="Missing Contact", ascending=False).reset_index(drop=True)
    tail = table[is_grand_total]
    table = pd.concat([body, tail], ignore_index=True)
    return table


def build_applicator_summary(df2_raw, df2_clean, df):
    total_calls = df2_clean.groupby("Applicator").size().reset_index(name="Total Calls")

    missing_number = df2_raw[
        (df2_raw["Contact"].isna()) | (df2_raw["Contact"].astype(str).str.strip().str.upper() == "NA")
    ].groupby("Applicator").size().reset_index(name="Missing Number")

    nc = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)].groupby("Applicator").size().reset_index(name="Not Connected")
    conn = df2_clean[~df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)].groupby("Applicator").size().reset_index(name="Connected Calls")
    pos = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(POSITIVE_LOWER)].groupby("Applicator").size().reset_index(name="Positive Feedback")
    neg = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(NEGATIVE_LOWER)].groupby("Applicator").size().reset_index(name="Negative Feedback")
    gen = df2_clean[
        (~df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)) &
        (~df2_clean["Sub Disposition"].str.lower().isin(EXCLUDED_LOWER))
    ].groupby("Applicator").size().reset_index(name="General Feedback")
    nps_given = df2_clean[df2_clean["NPS SCORE"] > 0].groupby("Applicator").size().reset_index(name="NPS Given")
    avg_nps = df2_clean[df2_clean["NPS SCORE"] > 0].groupby("Applicator")["NPS SCORE"].mean().round(2).reset_index(name="Avg NPS Score")

    table = total_calls \
        .merge(missing_number, on="Applicator", how="left") \
        .merge(nc, on="Applicator", how="left") \
        .merge(conn, on="Applicator", how="left") \
        .merge(pos, on="Applicator", how="left") \
        .merge(neg, on="Applicator", how="left") \
        .merge(gen, on="Applicator", how="left") \
        .merge(nps_given, on="Applicator", how="left") \
        .merge(avg_nps, on="Applicator", how="left")

    if "Effective Call List" in df2_clean.columns:
        running_total = df2_clean[df2_clean["Effective Call List"] == "Running"].groupby("Applicator").size().reset_index(name="Running")
        removal_total = df2_clean[df2_clean["Effective Call List"] == "Removal"].groupby("Applicator").size().reset_index(name="Removal")
        table = table.merge(running_total, on="Applicator", how="left")
        table = table.merge(removal_total, on="Applicator", how="left")

    count_cols = ["Missing Number", "Not Connected", "Connected Calls", "Positive Feedback",
                  "Negative Feedback", "General Feedback", "NPS Given", "Running", "Removal"]
    for c in count_cols:
        if c in table.columns:
            table[c] = table[c].fillna(0).astype(int)
    table["Avg NPS Score"] = table["Avg NPS Score"].fillna(0)

    table = table.rename(columns={"Applicator": "Applicators"})

    if df is not None and "Applicators" in df.columns:
        unique_patients = df.groupby("Applicators")["Patient Name"].nunique().reset_index(name="Unique Patients")
        total_dressings = df.groupby("Applicators")["Total Kit"].sum().reset_index(name="Total Dressings")
        patient_summary = unique_patients.merge(total_dressings, on="Applicators", how="outer")
        table = table.merge(patient_summary, on="Applicators", how="outer")
        for c in ["Unique Patients", "Total Dressings"]:
            table[c] = table[c].fillna(0).astype(int)

    table = table.sort_values(by="Total Calls", ascending=False).reset_index(drop=True)

    grand_total = {
        "Applicators": "Grand Total",
        "Total Calls": df2_clean.shape[0],
        "Missing Number": ((df2_raw["Contact"].isna()) | (df2_raw["Contact"].astype(str).str.strip().str.upper() == "NA")).sum(),
        "Not Connected": df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER).sum(),
        "Connected Calls": (~df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)).sum(),
        "Positive Feedback": df2_clean["Sub Disposition"].str.lower().isin(POSITIVE_LOWER).sum(),
        "Negative Feedback": df2_clean["Sub Disposition"].str.lower().isin(NEGATIVE_LOWER).sum(),
        "General Feedback": ((~df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)) & (~df2_clean["Sub Disposition"].str.lower().isin(EXCLUDED_LOWER))).sum(),
        "NPS Given": (df2_clean["NPS SCORE"] > 0).sum(),
        "Avg NPS Score": round(df2_clean.loc[df2_clean["NPS SCORE"] > 0, "NPS SCORE"].mean(), 2) if (df2_clean["NPS SCORE"] > 0).any() else 0,
    }
    if "Effective Call List" in df2_clean.columns:
        grand_total["Running"] = (df2_clean["Effective Call List"] == "Running").sum()
        grand_total["Removal"] = (df2_clean["Effective Call List"] == "Removal").sum()
    if df is not None:
        grand_total["Unique Patients"] = df["Patient Name"].nunique()
        grand_total["Total Dressings"] = df["Total Kit"].sum()

    table = pd.concat([table, pd.DataFrame([grand_total])], ignore_index=True)
    return table


def build_table3_applicator_defaulter(df2_raw, df2_clean, df):
    base = build_applicator_summary(df2_raw, df2_clean, df)
    cols = ["Applicators", "Unique Patients", "Total Dressings", "Missing Number",
            "Not Connected", "Connected Calls", "Positive Feedback", "Negative Feedback",
            "General Feedback", "NPS Given", "Avg NPS Score"]
    cols = [c for c in cols if c in base.columns]
    table = base[cols].copy()

    is_grand_total = table["Applicators"] == "Grand Total"
    body = table[~is_grand_total].sort_values(by="Missing Number", ascending=False).reset_index(drop=True)
    tail = table[is_grand_total]
    table = pd.concat([body, tail], ignore_index=True)
    return table


def build_hospital_doctor_summary(df2_raw, df):
    if "Hospital Name" not in df2_raw.columns or "Doctor" not in df2_raw.columns:
        return None

    total_contact = df2_raw.groupby(["Hospital Name", "Doctor"], dropna=False).size().reset_index(name="Total Contact")
    missing_contact = df2_raw[
        (df2_raw["Contact"].isna()) | (df2_raw["Contact"].astype(str).str.strip().str.upper() == "NA")
    ].groupby(["Hospital Name", "Doctor"], dropna=False).size().reset_index(name="Missing Contact")
    unique_contact = df2_raw.groupby(["Hospital Name", "Doctor"], dropna=False)["Contact"].nunique().reset_index(name="Total Unique Contact")

    table = total_contact.merge(missing_contact, on=["Hospital Name", "Doctor"], how="left")
    table = table.merge(unique_contact, on=["Hospital Name", "Doctor"], how="left")

    for c in ["Missing Contact", "Total Unique Contact"]:
        table[c] = table[c].fillna(0).astype(int)
    table["Hospital Name"] = table["Hospital Name"].fillna("Not Mentioned")
    table["Doctor"] = table["Doctor"].fillna("Not Mentioned")

    table = table.sort_values(by="Total Contact", ascending=False).reset_index(drop=True)

    grand_total = {
        "Hospital Name": "Grand Total",
        "Doctor": "",
        "Total Contact": df2_raw.shape[0],
        "Missing Contact": ((df2_raw["Contact"].isna()) | (df2_raw["Contact"].astype(str).str.strip().str.upper() == "NA")).sum(),
        "Total Unique Contact": df2_raw["Contact"].nunique(),
    }

    table = pd.concat([table, pd.DataFrame([grand_total])], ignore_index=True)
    return table


def build_quarterly_kit_summary(df):
    if df is None or "FY Quarter" not in df.columns or "Total Kit" not in df.columns:
        return None
    valid = df[df["FY Quarter"].notna()]
    if valid.empty:
        return None

    summary = (
        valid.groupby(["FY Quarter", "FY Quarter Sort"])["Total Kit"]
        .sum()
        .reset_index()
        .sort_values("FY Quarter Sort")
        .drop(columns="FY Quarter Sort")
        .reset_index(drop=True)
    )
    summary["Total Kit"] = summary["Total Kit"].astype(int)

    grand_total = pd.DataFrame({"FY Quarter": ["Grand Total"], "Total Kit": [int(valid["Total Kit"].sum())]})
    summary = pd.concat([summary, grand_total], ignore_index=True)
    return summary


def build_quarterly_kit_by_applicator(df):
    if df is None or "FY Quarter" not in df.columns or "Applicators" not in df.columns:
        return None
    valid = df[df["FY Quarter"].notna()]
    if valid.empty:
        return None

    quarter_order = (
        valid[["FY Quarter", "FY Quarter Sort"]]
        .drop_duplicates()
        .sort_values("FY Quarter Sort")["FY Quarter"]
        .tolist()
    )

    pivot = valid.groupby(["Applicators", "FY Quarter"])["Total Kit"].sum().unstack(fill_value=0)
    pivot = pivot.reindex(columns=quarter_order, fill_value=0)
    pivot = pivot.astype(int)
    pivot["Grand Total"] = pivot.sum(axis=1)
    pivot.loc["Grand Total"] = pivot.sum()
    pivot = pivot.reset_index()
    return pivot


def build_nps_category_table(df2_clean):
    nps_data = df2_clean[df2_clean["NPS SCORE"] > 0].copy()
    if nps_data.empty:
        return None

    def categorize(score):
        if score >= 9:
            return "Promoters (9-10)"
        elif score >= 7:
            return "Passives (7-8)"
        else:
            return "Detractors (0-6)"

    nps_data["NPS Category"] = nps_data["NPS SCORE"].apply(categorize)
    order = ["Promoters (9-10)", "Passives (7-8)", "Detractors (0-6)"]

    if "Effective Call List" in nps_data.columns:
        pivot = nps_data.groupby(["Effective Call List", "NPS Category"]).size().unstack(fill_value=0)
        pivot = pivot.reindex(columns=order, fill_value=0)
        pivot.loc["Overall"] = pivot.sum()
        pivot = pivot.reset_index().rename(columns={"Effective Call List": "Metric"})
    else:
        counts = nps_data["NPS Category"].value_counts().reindex(order, fill_value=0)
        pivot = pd.DataFrame({"Metric": ["Overall"], **{k: [v] for k, v in counts.items()}})

    return pivot


def build_connectivity_breakdown(df2_clean):
    if "Effective Call List" not in df2_clean.columns:
        return None

    conn = df2_clean[df2_clean["Connectivity Status"] == "Connected"]
    pos = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(POSITIVE_LOWER)]
    neg = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(NEGATIVE_LOWER)]
    gen = df2_clean[
        (df2_clean["Connectivity Status"] == "Connected") &
        (~df2_clean["Sub Disposition"].str.lower().isin(EXCLUDED_LOWER))
    ]

    def calllist_counts(data):
        counts = data.groupby("Effective Call List").size()
        running = counts.get("Running", 0)
        removal = counts.get("Removal", 0)
        return running, removal, running + removal

    rows = []
    for label, data in [("Total Connectivity", conn), ("Positive Feedback", pos),
                         ("Negative Feedback", neg), ("General Feedback", gen)]:
        r, rem, ov = calllist_counts(data)
        rows.append({"Metric": label, "Running": r, "Removal": rem, "Overall": ov})

    return pd.DataFrame(rows)


def build_not_connected_breakdown(df2_clean):
    nc_data = df2_clean[df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)]
    if nc_data.empty or "Effective Call List" not in nc_data.columns:
        return None
    pivot = nc_data.groupby(["Effective Call List", "Sub Disposition"]).size().unstack(fill_value=0)
    pivot.loc["Overall"] = pivot.sum()
    final = pivot.T.reset_index()
    final.columns.name = None
    final = final.rename(columns={"Sub Disposition": "METRIC"})
    return final


def build_general_breakdown(df2_clean):
    gen_data = df2_clean[
        (~df2_clean["Sub Disposition"].str.lower().isin(NOT_CONNECTED_LOWER)) &
        (~df2_clean["Sub Disposition"].str.lower().isin(EXCLUDED_LOWER))
    ]
    if gen_data.empty or "Effective Call List" not in gen_data.columns:
        return None
    pivot = gen_data.groupby(["Effective Call List", "Sub Disposition"]).size().unstack(fill_value=0)
    pivot.loc["Overall"] = pivot.sum()
    final = pivot.T.reset_index()
    final.columns.name = None
    final = final.rename(columns={"Sub Disposition": "METRIC"})
    return final


def compute_top_kpis(df2_clean):
    if df2_clean is None or df2_clean.empty:
        return {
            "Total Connectivity": 0, "Total Unique Connectivity": 0, "Not Connected": 0,
            "Positive Feedback": 0, "Negative Feedback": 0, "Avg NPS Score": 0,
        }

    connected_mask = df2_clean["Connectivity Status"] == "Connected"
    not_connected_mask = ~connected_mask
    pos_mask = df2_clean["Sub Disposition"].str.lower().isin(POSITIVE_LOWER)
    neg_mask = df2_clean["Sub Disposition"].str.lower().isin(NEGATIVE_LOWER)

    unique_connectivity = df2_clean.loc[connected_mask, "Contact"].nunique() if "Contact" in df2_clean.columns else connected_mask.sum()
    avg_nps = df2_clean.loc[df2_clean["NPS SCORE"] > 0, "NPS SCORE"].mean() if "NPS SCORE" in df2_clean.columns else None

    return {
        "Total Connectivity": int(connected_mask.sum()),
        "Total Unique Connectivity": int(unique_connectivity),
        "Not Connected": int(not_connected_mask.sum()),
        "Positive Feedback": int(pos_mask.sum()),
        "Negative Feedback": int(neg_mask.sum()),
        "Avg NPS Score": round(avg_nps, 2) if pd.notna(avg_nps) else 0,
    }


def build_performance_summary(df2_clean, df):
    if df2_clean is None or df2_clean.empty:
        return None

    has_ecl = "Effective Call List" in df2_clean.columns

    def bucket(mask):
        data = df2_clean[mask]
        if has_ecl:
            counts = data.groupby("Effective Call List").size()
            running = int(counts.get("Running", 0))
            removal = int(counts.get("Removal", 0))
        else:
            running, removal = 0, 0
        overall = int(mask.sum())
        return running, removal, overall

    def pct(a, b):
        return f"{(a / b * 100):.2f}%" if b else "0.00%"

    rows = []

    def add_row(metric, running, removal, overall):
        rows.append({"METRIC": metric, "RUNNING": running, "REMOVAL": removal, "OVERALL": overall})

    def add_section(name):
        rows.append({"METRIC": name, "RUNNING": "", "REMOVAL": "", "OVERALL": ""})

    add_section("OVERALL PERFORMANCE SUMMARY")

    if df is not None and "Patient Name" in df.columns:
        add_row("Total Unique Patients", "—", "—", int(df["Patient Name"].nunique()))
    if df is not None and "Total Kit" in df.columns:
        add_row("Total Kits Used", "—", "—", int(df["Total Kit"].sum()))

    dialed_r, dialed_rem, dialed_ov = bucket(pd.Series(True, index=df2_clean.index))
    add_row("Total Numbers Dialed", dialed_r, dialed_rem, dialed_ov)

    add_section("CONNECTIVITY")
    connected_mask = df2_clean["Connectivity Status"] == "Connected"
    conn_r, conn_rem, conn_ov = bucket(connected_mask)
    add_row("Total Connectivity", conn_r, conn_rem, conn_ov)
    add_row("Connectivity % (Connectivity / Dialed)",
            pct(conn_r, dialed_r), pct(conn_rem, dialed_rem), pct(conn_ov, dialed_ov))

    add_section("NPS SCORE ANALYSIS (of patients who gave a numeric score)")
    if has_ecl and "NPS SCORE" in df2_clean.columns:
        nps_by_list = df2_clean.loc[df2_clean["NPS SCORE"] > 0].groupby("Effective Call List")["NPS SCORE"].mean()
        avg_running = nps_by_list.get("Running", None)
        avg_removal = nps_by_list.get("Removal", None)
    else:
        avg_running = avg_removal = None
    avg_overall = df2_clean.loc[df2_clean["NPS SCORE"] > 0, "NPS SCORE"].mean() if "NPS SCORE" in df2_clean.columns else None
    add_row(
        "Avg NPS Score",
        round(avg_running, 2) if pd.notna(avg_running) else "—",
        round(avg_removal, 2) if pd.notna(avg_removal) else "—",
        round(avg_overall, 2) if pd.notna(avg_overall) else "—",
    )

    add_row("Overall Feedback given", conn_r, conn_rem, conn_ov)

    sub = df2_clean["Sub Disposition"].str.lower()
    has_number = df2_clean["NPS SCORE"] > 0 if "NPS SCORE" in df2_clean.columns else pd.Series(False, index=df2_clean.index)

    pos_base_mask = sub.isin(POSITIVE_LOWER)
    neg_base_mask = sub.isin(NEGATIVE_LOWER)

    pos_num_mask = pos_base_mask & has_number
    pos_verbal_mask = pos_base_mask & (~has_number)
    neg_num_mask = neg_base_mask & has_number
    neg_verbal_mask = neg_base_mask & (~has_number)
    gen_mask = (~sub.isin(NOT_CONNECTED_LOWER)) & (~sub.isin(EXCLUDED_LOWER))

    pr, prem, pov = bucket(pos_num_mask)
    add_row("Positive Feedback (NPS score in Number)", pr, prem, pov)
    pr2, prem2, pov2 = bucket(pos_verbal_mask)
    add_row("Positive Feedback (NPS score in verbal)", pr2, prem2, pov2)
    nr, nrem, nov = bucket(neg_num_mask)
    add_row("Negative Feedback (NPS score in Number)", nr, nrem, nov)
    nr2, nrem2, nov2 = bucket(neg_verbal_mask)
    add_row("Negative Feedback (NPS score not given)", nr2, nrem2, nov2)
    gr, grem, gov = bucket(gen_mask)
    add_row("General Query", gr, grem, gov)

    add_row("Total", pr + pr2 + nr + nr2 + gr, prem + prem2 + nrem + nrem2 + grem, pov + pov2 + nov + nov2 + gov)

    add_section("PROMOTERS / PASSIVES / DETRACTORS")
    if "NPS SCORE" in df2_clean.columns:
        promoters_mask = df2_clean["NPS SCORE"] >= 9
        passives_mask = (df2_clean["NPS SCORE"] >= 7) & (df2_clean["NPS SCORE"] < 9)
        detractors_mask = (df2_clean["NPS SCORE"] > 0) & (df2_clean["NPS SCORE"] < 7)
    else:
        promoters_mask = passives_mask = detractors_mask = pd.Series(False, index=df2_clean.index)

    r3, rem3, ov3 = bucket(promoters_mask)
    add_row("Promoters (Score 9-10)", r3, rem3, ov3)
    r4, rem4, ov4 = bucket(passives_mask)
    add_row("Passives (Score 7-8)", r4, rem4, ov4)
    r5, rem5, ov5 = bucket(detractors_mask)
    add_row("Detractors (Score 0-6)", r5, rem5, ov5)

    add_section("DIAL EFFICIENCY")
    if has_ecl and "Contact" in df2_clean.columns:
        contacts_by_list = df2_clean.groupby("Effective Call List")["Contact"].nunique()
        calls_by_list = df2_clean.groupby("Effective Call List").size()

        def dials_per_patient(key):
            c = contacts_by_list.get(key, 0)
            n = calls_by_list.get(key, 0)
            return round(n / c, 2) if c else "—"

        dpp_running = dials_per_patient("Running")
        dpp_removal = dials_per_patient("Removal")
    else:
        dpp_running = dpp_removal = "—"

    overall_contacts = df2_clean["Contact"].nunique() if "Contact" in df2_clean.columns else 0
    dpp_overall = round(len(df2_clean) / overall_contacts, 2) if overall_contacts else "—"
    add_row("Avg Dials per Patient", dpp_running, dpp_removal, dpp_overall)

    return pd.DataFrame(rows)


def build_monthly_trend(df2_clean):
    if df2_clean is None or df2_clean.empty or "Month" not in df2_clean.columns:
        return None

    d = df2_clean.copy()
    if "Month Sort" not in d.columns:
        return None

    d = d[d["Month"].notna()]
    if d.empty:
        return None

    conn_mask = d["Connectivity Status"] == "Connected"
    nc_mask = ~conn_mask
    pos_mask = d["Sub Disposition"].str.lower().isin(POSITIVE_LOWER)
    neg_mask = d["Sub Disposition"].str.lower().isin(NEGATIVE_LOWER)

    base = d[["Month", "Month Sort"]].drop_duplicates()

    def counts_for(mask, name):
        return d[mask].groupby(["Month", "Month Sort"]).size().reset_index(name=name)

    summary = base
    for mask, name in [
        (conn_mask, "Total Connectivity"),
        (nc_mask, "Total Not Connected"),
        (pos_mask, "Positive Feedback"),
        (neg_mask, "Negative Feedback"),
    ]:
        summary = summary.merge(counts_for(mask, name), on=["Month", "Month Sort"], how="left")

    for c in ["Total Connectivity", "Total Not Connected", "Positive Feedback", "Negative Feedback"]:
        summary[c] = summary[c].fillna(0).astype(int)

    summary = summary.sort_values("Month Sort").drop(columns="Month Sort").reset_index(drop=True)
    return summary


def build_monthly_trend_chart(monthly_trend):
    metrics = ["Total Connectivity", "Total Not Connected", "Positive Feedback", "Negative Feedback"]
    long_df = monthly_trend.melt(id_vars="Month", value_vars=metrics, var_name="Metric", value_name="Count")
    month_order = monthly_trend["Month"].tolist()

    base = alt.Chart(long_df).encode(
        x=alt.X("Month:N", sort=month_order, title="Month"),
        y=alt.Y("Count:Q", title="Count"),
        color=alt.Color("Metric:N", title="Metric"),
    )

    line = base.mark_line(point=True)
    labels = base.mark_text(dy=-10, fontSize=11).encode(text="Count:Q")

    chart = (line + labels).properties(height=400)
    return chart


# ------------------------------------------------------------------
# ------------------- MONTH-ON-MONTH COMPARISON -----------------------
# ------------------------------------------------------------------

def pct_diff(old_val, new_val):
    def is_blank(v):
        return v is None or (isinstance(v, str) and v.strip() in ("—", "-", ""))

    if is_blank(old_val) or is_blank(new_val):
        return "—"

    def parse_percent(v):
        if isinstance(v, str) and v.strip().endswith("%"):
            try:
                return float(v.strip().rstrip("%"))
            except ValueError:
                return None
        return None

    old_pct = parse_percent(old_val)
    new_pct = parse_percent(new_val)
    if old_pct is not None and new_pct is not None:
        return f"{(new_pct - old_pct):.2f}%"

    try:
        old_num = float(old_val)
        new_num = float(new_val)
    except (TypeError, ValueError):
        return ""

    if old_num == 0:
        return "0.00%" if new_num == 0 else "New"

    return f"{((new_num - old_num) / abs(old_num) * 100):.2f}%"


def build_performance_comparison(table_a, table_b):
    if table_a is None or table_b is None:
        return None
    rows = []
    for (_, ra), (_, rb) in zip(table_a.iterrows(), table_b.iterrows()):
        metric = ra["METRIC"]
        if ra["RUNNING"] == "" and ra["REMOVAL"] == "" and ra["OVERALL"] == "":
            rows.append({"METRIC": metric, "RUNNING": "", "REMOVAL": "", "OVERALL": ""})
            continue
        rows.append({
            "METRIC": metric,
            "RUNNING": pct_diff(ra["RUNNING"], rb["RUNNING"]),
            "REMOVAL": pct_diff(ra["REMOVAL"], rb["REMOVAL"]),
            "OVERALL": pct_diff(ra["OVERALL"], rb["OVERALL"]),
        })
    return pd.DataFrame(rows)


def build_generic_comparison(table_a, table_b, key_cols, metric_cols, rename_map=None):
    if table_a is None or table_b is None:
        return None

    key0 = key_cols[0]
    a = table_a[table_a[key0] != "Grand Total"].copy()
    b = table_b[table_b[key0] != "Grand Total"].copy()

    merged = a.merge(b, on=key_cols, how="outer", suffixes=("_A", "_B"))
    out = merged[key_cols].copy()

    for m in metric_cols:
        col_a, col_b = f"{m}_A", f"{m}_B"
        if col_a not in merged.columns or col_b not in merged.columns:
            continue
        label = (rename_map or {}).get(m, m)
        out[f"{label} % Diff"] = [pct_diff(x, y) for x, y in zip(merged[col_a], merged[col_b])]

    return out


# ------------------------------------------------------------------
# ------------------------------ UI ----------------------------------
# ------------------------------------------------------------------

st.title("Call & Patient Data — Automated Report")
st.caption("Upload your files once, get every report instantly. Use the filters below to drill into any Applicator, City, Month, etc. — just like an Excel filter.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Call Data (df2)")
    file_df2 = st.file_uploader("Upload call/telecalling data (CSV or Excel)", type=["csv", "xlsx", "xls"], key="df2")
with col2:
    st.subheader("2. Patient / Dressing Data (df) — optional")
    file_df = st.file_uploader("Upload patient/dressing data (CSV or Excel)", type=["csv", "xlsx", "xls"], key="df")

if st.button("Load Data", type="primary", disabled=(file_df2 is None)):
    with st.spinner("Reading and cleaning your data..."):
        df2_raw = load_file(file_df2)
        df_raw = load_file(file_df)

        df2_raw = prepare_df2(df2_raw)
        df_prepared = prepare_df(df_raw)

        st.session_state["df2_prepared"] = df2_raw
        st.session_state["df_prepared"] = df_prepared
        st.success("Data loaded! Use the filters below, then scroll down for your tables.")

if "df2_prepared" in st.session_state:
    df2_full = st.session_state["df2_prepared"]
    df_full = st.session_state["df_prepared"]

    st.markdown("### Filters")
    st.caption("Leave a filter empty to include everyone / everything for that field.")

    f1, f2, f3, f4, f5 = st.columns(5)

    applicator_options = sorted(df2_full["Applicator"].dropna().unique()) if "Applicator" in df2_full.columns else []
    city_options = sorted(df2_full["City"].dropna().unique()) if "City" in df2_full.columns else []
    if "Month" in df2_full.columns and "Month Sort" in df2_full.columns:
        month_options = (
            df2_full[["Month", "Month Sort"]].dropna().drop_duplicates()
            .sort_values("Month Sort")["Month"].tolist()
        )
    elif "Month" in df2_full.columns:
        month_options = sorted(df2_full["Month"].dropna().astype(str).unique())
    else:
        month_options = []
    calllist_options = sorted(df2_full["Call List"].dropna().unique()) if "Call List" in df2_full.columns else []
    if df_full is not None and "FY Quarter" in df_full.columns:
        quarter_options = (
            df_full[["FY Quarter", "FY Quarter Sort"]].dropna().drop_duplicates()
            .sort_values("FY Quarter Sort")["FY Quarter"].tolist()
        )
    else:
        quarter_options = []

    with f1:
        sel_applicator = st.multiselect("Applicator", applicator_options, key="filter_applicator")
    with f2:
        sel_city = st.multiselect("City", city_options, key="filter_city")
    with f3:
        sel_month = st.multiselect("Month", month_options, key="filter_month")
    with f4:
        sel_calllist = st.multiselect("Call List", calllist_options, key="filter_calllist")
    with f5:
        sel_quarter = st.multiselect("Quarter (FY)", quarter_options, key="filter_quarter")

    df2_filtered = df2_full.copy()
    if sel_applicator:
        df2_filtered = df2_filtered[df2_filtered["Applicator"].isin(sel_applicator)]
    if sel_city:
        df2_filtered = df2_filtered[df2_filtered["City"].isin(sel_city)]
    if sel_month:
        df2_filtered = df2_filtered[df2_filtered["Month"].astype(str).isin(sel_month)]
    if sel_calllist:
        df2_filtered = df2_filtered[df2_filtered["Call List"].isin(sel_calllist)]

    df_filtered = df_full
    if df_full is not None:
        df_filtered = df_full.copy()
        if sel_applicator and "Applicators" in df_filtered.columns:
            df_filtered = df_filtered[df_filtered["Applicators"].isin(sel_applicator)]
        if sel_city and "City" in df_filtered.columns:
            df_filtered = df_filtered[df_filtered["City"].isin(sel_city)]
        if sel_month and "Month" in df_filtered.columns:
            df_filtered = df_filtered[df_filtered["Month"].astype(str).isin(sel_month)]
        if sel_quarter and "FY Quarter" in df_filtered.columns:
            df_filtered = df_filtered[df_filtered["FY Quarter"].isin(sel_quarter)]

    st.caption(f"Showing **{len(df2_filtered)}** of **{len(df2_full)}** call records after filters.")

    df2_clean = get_clean_contact_df2(df2_filtered)
    df2_clean = add_connectivity_status(df2_clean)
    df2_clean = apply_rating_shift(df2_clean)

    kpis = compute_top_kpis(df2_clean)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total Connectivity", kpis["Total Connectivity"])
    k2.metric("Unique Connectivity", kpis["Total Unique Connectivity"])
    k3.metric("Not Connected", kpis["Not Connected"])
    k4.metric("Positive Feedback", kpis["Positive Feedback"])
    k5.metric("Negative Feedback", kpis["Negative Feedback"])
    k6.metric("Avg NPS Score", kpis["Avg NPS Score"])

    k7, k8, k9, k10 = st.columns(4)
    k7.metric("Total Unique Patients", int(df_filtered["Patient Name"].nunique()) if df_filtered is not None and "Patient Name" in df_filtered.columns else "—")
    k8.metric("Total Cities", int(df_filtered["City"].nunique()) if df_filtered is not None and "City" in df_filtered.columns else "—")
    k9.metric("Total Hospitals", int(df_filtered["Hospital Name"].nunique()) if df_filtered is not None and "Hospital Name" in df_filtered.columns else "—")
    k10.metric("Total Applicators", int(df_filtered["Applicators"].nunique()) if df_filtered is not None and "Applicators" in df_filtered.columns else "—")

    st.divider()

    st.markdown("## India Level — Table 1: Overall Performance Summary")
    india_summary = build_performance_summary(df2_clean, df_filtered)
    if india_summary is not None:
        st.dataframe(india_summary, use_container_width=True, hide_index=True)

    india_nc = build_not_connected_breakdown(df2_clean)
    if india_nc is not None:
        st.markdown("**Not Connected — breakdown**")
        st.dataframe(india_nc, use_container_width=True, hide_index=True)

    india_gen = build_general_breakdown(df2_clean)
    if india_gen is not None:
        st.markdown("**General Query — breakdown**")
        st.dataframe(india_gen, use_container_width=True, hide_index=True)

    st.divider()

    st.markdown("## Month-Wise Trend — India Level")
    monthly_trend = build_monthly_trend(df2_clean)
    if monthly_trend is not None and not monthly_trend.empty:
        st.dataframe(monthly_trend, use_container_width=True, hide_index=True)
        st.altair_chart(build_monthly_trend_chart(monthly_trend), use_container_width=True)
    else:
        st.info("No 'Month' column found in the call data, so the month-wise trend can't be built.")

    st.divider()

    st.markdown("## Table 2 — City-Wise NPS Performance Breakdown")
    table2_city = build_table2_city_nps(df2_filtered, df2_clean, df_filtered)
    st.dataframe(table2_city, use_container_width=True, hide_index=True)

    st.divider()

    st.markdown("## Table 3 — Applicator Performance & Defaulter Ranking")
    table3_applicator = build_table3_applicator_defaulter(df2_filtered, df2_clean, df_filtered)
    st.dataframe(table3_applicator, use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("## Hospital-Doctor Summary")
    hospital_doctor_table = build_hospital_doctor_summary(df2_filtered, df_filtered)
    if hospital_doctor_table is not None:
        st.dataframe(hospital_doctor_table, use_container_width=True, hide_index=True)
    else:
        st.info("No 'Hospital Name' / 'Doctor' columns found in the call data.")

    st.divider()

    # ------------------------------------------------------------------
    # ------------------- MONTH-ON-MONTH COMPARISON -----------------------
    # ------------------------------------------------------------------
    # *** FIX: MoM comparison is now completely independent of the sidebar
    # filters (Applicator/City/Month/Call List/Quarter). Previously it built
    # its two month-snapshots from df2_filtered / df_filtered, so if the
    # person had e.g. a specific Month selected in the filters above, one (or
    # both) of the two comparison months would have zero rows in the filtered
    # data -> the comparison came out all NaN. It now always starts from the
    # full, unfiltered df2_full / df_full, so it works the same no matter what
    # the sidebar filters are set to.
    st.markdown("## Month-on-Month Comparison")
    st.caption(
        "Pick two months (e.g. April vs June) to see the % change for every table above — "
        "same formulas as Table 1 / City / Applicator / Hospital-Doctor, just diffed between the two months. "
        "This comparison always uses the full dataset and ignores the filters above."
    )

    if len(month_options) >= 2:
        cmp_col1, cmp_col2 = st.columns(2)
        with cmp_col1:
            month_a = st.selectbox("Month A (baseline)", month_options, index=0, key="cmp_month_a")
        with cmp_col2:
            month_b = st.selectbox(
                "Month B (compare to)", month_options,
                index=1 if len(month_options) > 1 else 0, key="cmp_month_b",
            )

        if st.button("Generate Month-on-Month Comparison", type="primary"):
            def build_month_snapshot(month_label):
                """
                Deliberately starts from df2_full / df_full (NOT df2_filtered /
                df_filtered) so this comparison is always computed from the
                complete dataset, regardless of what's selected in the sidebar
                filters above.
                """
                raw_month = df2_full[df2_full["Month"].astype(str) == month_label].copy()
                patients_month = None
                if df_full is not None and "Month" in df_full.columns:
                    patients_month = df_full[df_full["Month"].astype(str) == month_label].copy()
                clean_month = get_clean_contact_df2(raw_month)
                clean_month = add_connectivity_status(clean_month)
                clean_month = apply_rating_shift(clean_month)
                return raw_month, clean_month, patients_month

            raw_a, clean_a, patients_a = build_month_snapshot(month_a)
            raw_b, clean_b, patients_b = build_month_snapshot(month_b)

            perf_a = build_performance_summary(clean_a, patients_a)
            perf_b = build_performance_summary(clean_b, patients_b)
            perf_cmp = build_performance_comparison(perf_a, perf_b)

            city_a = build_table2_city_nps(raw_a, clean_a, patients_a)
            city_b = build_table2_city_nps(raw_b, clean_b, patients_b)
            city_cmp = build_generic_comparison(
                city_a, city_b, key_cols=["City"],
                metric_cols=["Unique Patients", "Missing Contact", "Total Not Connected",
                             "Total Connected", "Positive Feedback", "Negative Feedback", "General Query"],
            )

            app_a = build_table3_applicator_defaulter(raw_a, clean_a, patients_a)
            app_b = build_table3_applicator_defaulter(raw_b, clean_b, patients_b)
            app_cmp = build_generic_comparison(
                app_a, app_b, key_cols=["Applicators"],
                metric_cols=["Unique Patients", "Total Dressings", "Missing Number", "Not Connected",
                             "Connected Calls", "Positive Feedback", "Negative Feedback",
                             "General Feedback", "NPS Given", "Avg NPS Score"],
            )

            hd_a = build_hospital_doctor_summary(raw_a, patients_a)
            hd_b = build_hospital_doctor_summary(raw_b, patients_b)
            hd_cmp = build_generic_comparison(
                hd_a, hd_b, key_cols=["Hospital Name", "Doctor"],
                metric_cols=["Missing Contact", "Total Unique Contact"],
                rename_map={"Total Unique Contact": "Total Unique Patients"},
            ) if hd_a is not None and hd_b is not None else None

            st.session_state["cmp_results"] = {
                "label": f"{month_a} vs {month_b}",
                "Table 1 - Performance Comparison": perf_cmp,
                "City Comparison": city_cmp,
                "Applicator Comparison": app_cmp,
                "Hospital-Doctor Comparison": hd_cmp,
            }

        if "cmp_results" in st.session_state:
            cmp_results = st.session_state["cmp_results"]
            st.caption(f"Showing: **{cmp_results['label']}**")

            if cmp_results["Table 1 - Performance Comparison"] is not None:
                st.markdown("**Table 1 — Overall Performance Comparison**")
                st.dataframe(cmp_results["Table 1 - Performance Comparison"], use_container_width=True, hide_index=True)

            if cmp_results["City Comparison"] is not None:
                st.markdown("**City Comparison**")
                st.dataframe(cmp_results["City Comparison"], use_container_width=True, hide_index=True)

            if cmp_results["Applicator Comparison"] is not None:
                st.markdown("**Applicator Comparison**")
                st.dataframe(cmp_results["Applicator Comparison"], use_container_width=True, hide_index=True)

            if cmp_results["Hospital-Doctor Comparison"] is not None:
                st.markdown("**Hospital-Doctor Comparison**")
                st.dataframe(cmp_results["Hospital-Doctor Comparison"], use_container_width=True, hide_index=True)
    else:
        st.info("Need at least two distinct months in the data to build a Month-on-Month comparison.")

    st.divider()

    results = {}

    try:
        results["Overall Performance Summary (India)"] = india_summary
    except Exception:
        pass

    try:
        results["Month-Wise Trend"] = monthly_trend
    except Exception:
        pass

    try:
        results["Table 2 - City NPS Breakdown"] = table2_city
    except Exception:
        pass

    try:
        results["Table 3 - Applicator Defaulter Ranking"] = table3_applicator
    except Exception:
        pass

    try:
        results["City Summary (full)"] = build_city_summary(df2_filtered, df2_clean, df_filtered)
    except Exception as e:
        st.warning(f"City Summary could not be built: {e}")

    try:
        results["Applicator Summary (full)"] = build_applicator_summary(df2_filtered, df2_clean, df_filtered)
    except Exception as e:
        st.warning(f"Applicator Summary could not be built: {e}")

    try:
        hd = build_hospital_doctor_summary(df2_filtered, df_filtered)
        if hd is not None:
            results["Hospital-Doctor Summary"] = hd
    except Exception as e:
        st.warning(f"Hospital-Doctor Summary could not be built: {e}")

    try:
        nps_table = build_nps_category_table(df2_clean)
        if nps_table is not None:
            results["NPS Categories"] = nps_table
    except Exception as e:
        st.warning(f"NPS Category table could not be built: {e}")

    try:
        conn_table = build_connectivity_breakdown(df2_clean)
        if conn_table is not None:
            results["Connectivity Breakdown"] = conn_table
    except Exception as e:
        st.warning(f"Connectivity Breakdown could not be built: {e}")

    try:
        nc_table = build_not_connected_breakdown(df2_clean)
        if nc_table is not None:
            results["Not Connected Breakdown"] = nc_table
    except Exception as e:
        st.warning(f"Not Connected Breakdown could not be built: {e}")

    try:
        gen_table = build_general_breakdown(df2_clean)
        if gen_table is not None:
            results["General Feedback Breakdown"] = gen_table
    except Exception as e:
        st.warning(f"General Feedback Breakdown could not be built: {e}")

    try:
        qtr_table = build_quarterly_kit_summary(df_filtered)
        if qtr_table is not None:
            results["Quarterly Kit Progress"] = qtr_table
    except Exception as e:
        st.warning(f"Quarterly Kit Progress could not be built: {e}")

    try:
        qtr_app_table = build_quarterly_kit_by_applicator(df_filtered)
        if qtr_app_table is not None:
            results["Quarterly Kit by Applicator"] = qtr_app_table
    except Exception as e:
        st.warning(f"Quarterly Kit by Applicator could not be built: {e}")

    if "cmp_results" in st.session_state:
        cmp_results = st.session_state["cmp_results"]
        for name, table in cmp_results.items():
            if name == "label" or table is None:
                continue
            results[f"MoM {name}"] = table

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        for name, table in results.items():
            if table is None:
                continue
            sheet_name = name[:31]
            table.to_excel(writer, sheet_name=sheet_name, index=False)
    excel_buffer.seek(0)

    st.download_button(
        label="Download Report (Excel, current filters applied)",
        data=excel_buffer,
        file_name="report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
