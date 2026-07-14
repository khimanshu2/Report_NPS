"""
Call Center + Patient Data Automation Report
==============================================
Upload the two data files (Call Data = df2, Patient/Dressing Data = df)
and get every summary table automatically, plus one combined Excel
download with each table as a separate sheet.
"""

import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Call & Patient Report Automation", layout="wide")

# ------------------------------------------------------------------
# ---------------------------- CONSTANTS ----------------------------
# ------------------------------------------------------------------

NOT_CONNECTED_LIST = [
    'Not Connected', 'Busy / asked to call later', 'DNP 1', 'DNP 2', 'DNP 3',
    'Continue switched off', 'Invalid Number', 'Number not active'
]
POSITIVE_LIST = ['Connected – Feedback Positive', 'Everything is good, no issue at all']
NEGATIVE_LIST = ['Connected – Feedback Negative', 'Overall negative feedback or dissatisfaction']
LANGUAGE_LIST = ['Connected - Language Barrier', 'Telgu', 'Malyalam', 'Kannada', 'Tamil']
EXCLUDED_LIST = POSITIVE_LIST + NEGATIVE_LIST

# Lowercase versions for robust, case-insensitive matching against the actual data
# (so "DNP 1", "dnp 1", "Dnp1" etc. all match the same way).
NOT_CONNECTED_LOWER = {x.strip().lower() for x in NOT_CONNECTED_LIST}
POSITIVE_LOWER = {x.strip().lower() for x in POSITIVE_LIST}
NEGATIVE_LOWER = {x.strip().lower() for x in NEGATIVE_LIST}
LANGUAGE_LOWER = {x.strip().lower() for x in LANGUAGE_LIST}
EXCLUDED_LOWER = POSITIVE_LOWER | NEGATIVE_LOWER


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

    # Encoding fix + clean text
    if "Sub Disposition" in df2.columns:
        df2["Sub Disposition"] = (
            df2["Sub Disposition"].astype(str).str.strip()
            .str.replace("â€“", "–", regex=False)
        )
        df2["Sub Disposition"] = df2["Sub Disposition"].replace("Nan", pd.NA)

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

    rename_map = {}
    if col_applicator: rename_map[col_applicator] = "Applicators"
    if col_patient: rename_map[col_patient] = "Patient Name"
    if col_kit: rename_map[col_kit] = "Total Kit"
    if col_hospital: rename_map[col_hospital] = "Hospital Name"
    if col_doctor: rename_map[col_doctor] = "Doctor"
    if col_city: rename_map[col_city] = "City"

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

    return df


def get_clean_contact_df2(df2):
    """df2 rows where Contact is present and not the text 'NA'."""
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
        lambda x: "Not Connected" if x in NOT_CONNECTED_LIST else "Connected"
    )
    return df2_clean


def apply_rating_shift(df2_clean):
    """
    Business rule: agar kisi record mein NPS Score diya gaya hai (rating hai),
    to us record ko 'Removal' list mein count karo, chahe uska original
    Call List 'Running' ho. Original data change nahi hota, sirf ek naya
    'Effective Call List' column banta hai jisse saari tables calculate hoti hain.
    """
    df2_clean = df2_clean.copy()
    if "Call List" not in df2_clean.columns:
        return df2_clean

    df2_clean["Effective Call List"] = df2_clean["Call List"]
    if "NPS SCORE" in df2_clean.columns:
        rated_mask = df2_clean["NPS SCORE"] > 0
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
    gen_total = gen_data.groupby("City").size().reset_index(name="General Feedback")

    nps_given = df2_clean[df2_clean["NPS SCORE"] > 0].groupby("City").size().reset_index(name="NPS Given")
    avg_nps = df2_clean[df2_clean["NPS SCORE"] > 0].groupby("City")["NPS SCORE"].mean().round(2).reset_index(name="Avg NPS Score")

    table = conn_total.merge(nc_total, on="City", how="outer") \
        .merge(missing_contact, on="City", how="left") \
        .merge(pos_total, on="City", how="left") \
        .merge(neg_total, on="City", how="left") \
        .merge(gen_total, on="City", how="left") \
        .merge(nps_given, on="City", how="left") \
        .merge(avg_nps, on="City", how="left")

    # ---- Running / Removal breakdown (uses the shifted "Effective Call List") ----
    if "Effective Call List" in df2_clean.columns:
        running_total = df2_clean[df2_clean["Effective Call List"] == "Running"].groupby("City").size().reset_index(name="Running")
        removal_total = df2_clean[df2_clean["Effective Call List"] == "Removal"].groupby("City").size().reset_index(name="Removal")
        table = table.merge(running_total, on="City", how="left")
        table = table.merge(removal_total, on="City", how="left")

    fill_cols = ["Total Connected", "Total Not Connected", "Missing Contact",
                 "Positive Feedback", "Negative Feedback", "General Feedback", "NPS Given",
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
        "General Feedback": gen_data.shape[0],
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

    # ---- Running / Removal breakdown (uses the shifted "Effective Call List") ----
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

    if df is not None and "Hospital Name" in df.columns and "Doctor" in df.columns:
        unique_patients = df.groupby(["Hospital Name", "Doctor"], dropna=False)["Patient Name"].nunique().reset_index(name="Total Unique Patients")
        unique_patients["Hospital Name"] = unique_patients["Hospital Name"].fillna("Not Mentioned")
        unique_patients["Doctor"] = unique_patients["Doctor"].fillna("Not Mentioned")
        table = table.merge(unique_patients, on=["Hospital Name", "Doctor"], how="outer")
        table["Total Unique Patients"] = table["Total Unique Patients"].fillna(0).astype(int)
        for c in ["Total Contact", "Missing Contact", "Total Unique Contact"]:
            table[c] = table[c].fillna(0).astype(int)

    table = table.sort_values(by="Total Contact", ascending=False).reset_index(drop=True)

    grand_total = {
        "Hospital Name": "Grand Total",
        "Doctor": "",
        "Total Contact": df2_raw.shape[0],
        "Missing Contact": ((df2_raw["Contact"].isna()) | (df2_raw["Contact"].astype(str).str.strip().str.upper() == "NA")).sum(),
        "Total Unique Contact": df2_raw["Contact"].nunique(),
    }
    if df is not None:
        grand_total["Total Unique Patients"] = df["Patient Name"].nunique()

    table = pd.concat([table, pd.DataFrame([grand_total])], ignore_index=True)
    return table


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
    """Total Connectivity / Positive / Negative / General with Running / Removal / Overall."""
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
    """
    Pivot table for 'General Feedback' — i.e. Connected calls whose Sub Disposition
    is NOT one of the Not Connected / Positive / Negative categories (things like
    'Connected – Refused', 'Connected – Machine Issue Feedback', language-barrier
    dispositions, etc). Same METRIC x Running/Removal/Overall shape as the
    Not Connected breakdown.
    """
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


# ------------------------------------------------------------------
# ------------------------------ UI ----------------------------------
# ------------------------------------------------------------------

st.title("📊 Call & Patient Data — Automated Report")
st.caption("Upload your files once, get every report instantly. Use the filters below to drill into any Applicator, City, Month, etc. — just like an Excel filter.")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1️⃣ Call Data (df2)")
    file_df2 = st.file_uploader("Upload call/telecalling data (CSV or Excel)", type=["csv", "xlsx", "xls"], key="df2")
with col2:
    st.subheader("2️⃣ Patient / Dressing Data (df) — optional")
    file_df = st.file_uploader("Upload patient/dressing data (CSV or Excel)", type=["csv", "xlsx", "xls"], key="df")

if st.button("🚀 Load Data", type="primary", disabled=(file_df2 is None)):
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

    # ------------------------------------------------------------------
    # ---------------------- EXCEL-STYLE FILTERS --------------------------
    # ------------------------------------------------------------------
    st.markdown("### 🔍 Filters")
    st.caption("Leave a filter empty to include everyone / everything for that field.")

    f1, f2, f3, f4 = st.columns(4)

    applicator_options = sorted(df2_full["Applicator"].dropna().unique()) if "Applicator" in df2_full.columns else []
    city_options = sorted(df2_full["City"].dropna().unique()) if "City" in df2_full.columns else []
    month_options = sorted(df2_full["Month"].dropna().astype(str).unique()) if "Month" in df2_full.columns else []
    calllist_options = sorted(df2_full["Call List"].dropna().unique()) if "Call List" in df2_full.columns else []

    with f1:
        sel_applicator = st.multiselect("Applicator", applicator_options, key="filter_applicator")
    with f2:
        sel_city = st.multiselect("City", city_options, key="filter_city")
    with f3:
        sel_month = st.multiselect("Month", month_options, key="filter_month")
    with f4:
        sel_calllist = st.multiselect("Call List", calllist_options, key="filter_calllist")

    # ---- Apply filters to df2 ----
    df2_filtered = df2_full.copy()
    if sel_applicator:
        df2_filtered = df2_filtered[df2_filtered["Applicator"].isin(sel_applicator)]
    if sel_city:
        df2_filtered = df2_filtered[df2_filtered["City"].isin(sel_city)]
    if sel_month:
        df2_filtered = df2_filtered[df2_filtered["Month"].astype(str).isin(sel_month)]
    if sel_calllist:
        df2_filtered = df2_filtered[df2_filtered["Call List"].isin(sel_calllist)]

    # ---- Apply the same Applicator / City filters to df (patient data), if present ----
    df_filtered = df_full
    if df_full is not None:
        df_filtered = df_full.copy()
        if sel_applicator and "Applicators" in df_filtered.columns:
            df_filtered = df_filtered[df_filtered["Applicators"].isin(sel_applicator)]
        if sel_city and "City" in df_filtered.columns:
            df_filtered = df_filtered[df_filtered["City"].isin(sel_city)]

    st.caption(f"Showing **{len(df2_filtered)}** of **{len(df2_full)}** call records after filters.")

    # ------------------------------------------------------------------
    # ------------------------- BUILD REPORTS ------------------------------
    # ------------------------------------------------------------------
    df2_clean = get_clean_contact_df2(df2_filtered)
    df2_clean = add_connectivity_status(df2_clean)
    df2_clean = apply_rating_shift(df2_clean)  # NPS-rated calls count as "Removal"

    results = {}

    try:
        results["City Summary"] = build_city_summary(df2_filtered, df2_clean, df_filtered)
    except Exception as e:
        st.warning(f"City Summary could not be built: {e}")

    try:
        results["Applicator Summary"] = build_applicator_summary(df2_filtered, df2_clean, df_filtered)
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

    st.success(f"✅ {len(results)} table(s) ready.")

    for name, table in results.items():
        st.subheader(name)
        st.dataframe(table, use_container_width=True)

    # Build combined Excel for download (reflects current filters)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        for name, table in results.items():
            sheet_name = name[:31]  # Excel sheet name limit
            table.to_excel(writer, sheet_name=sheet_name, index=False)
    excel_buffer.seek(0)

    st.download_button(
        label="⬇️ Download Report (Excel, current filters applied)",
        data=excel_buffer,
        file_name="report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )