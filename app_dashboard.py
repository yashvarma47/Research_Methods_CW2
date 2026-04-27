# Libraries used for file paths, Excel loading, cleaning and plotting.
from pathlib import Path
import glob
import os
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Basic page setup for the Streamlit dashboard.
st.set_page_config(page_title="NHS Hospital Admissions Dashboard", layout="wide")

# Custom CSS for a clean white dashboard style.
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0E1117;
        color: #F5F5F5;
    }

    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }

    .metric-card {
        background: #1B1F2A;
        border: 1px solid #303747;
        border-radius: 14px;
        padding: 16px 18px;
        box-shadow: 0 1px 8px rgba(0,0,0,0.35);
        color: #F5F5F5;
    }

    h1, h2, h3, h4, h5, h6, p, label {
        color: #F5F5F5 !important;
    }

    [data-testid="stSidebar"] {
        background-color: #111827 !important;
    }

    [data-testid="stSidebar"] * {
        color: #F5F5F5 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Shared Plotly template so all charts follow the dark dashboard style.
PLOTLY_TEMPLATE = "plotly_dark"

# NHS workbook sheet names are not exactly the same every year, so aliases are needed.
SHEET_ALIASES = {
    "summary": ["Primary Diagnosis Summary", "Primary diagnosis - summary"],
    "3char": ["Primary Diagnosis 3 Character", "Primary diagnosis - 3 char"],
    "4char": ["Primary Diagnosis 4 Character", "Primary diagnosis - 4 char"],
}

# Standard age column names used inside the dashboard.
AGE_CANONICAL = {
    "age_0": ["Age 0", "Age 0 (FCE)"],
    "age_1_4": ["Age 1-4", "Age 1-4 (FCE)"],
    "age_5_9": ["Age 5-9", "Age 5-9 (FCE)"],
    "age_10_14": ["Age 10-14", "Age 10-14 (FCE)"],
    "age_15": ["Age 15", "Age 15 (FCE)"],
    "age_16": ["Age 16", "Age 16 (FCE)"],
    "age_17": ["Age 17", "Age 17 (FCE)"],
    "age_18": ["Age 18", "Age 18 (FCE)"],
    "age_19": ["Age 19", "Age 19 (FCE)"],
    "age_20_24": ["Age 20-24", "Age 20-24 (FCE)"],
    "age_25_29": ["Age 25-29", "Age 25-29 (FCE)"],
    "age_30_34": ["Age 30-34", "Age 30-34 (FCE)"],
    "age_35_39": ["Age 35-39", "Age 35-39 (FCE)"],
    "age_40_44": ["Age 40-44", "Age 40-44 (FCE)"],
    "age_45_49": ["Age 45-49", "Age 45-49 (FCE)"],
    "age_50_54": ["Age 50-54", "Age 50-54 (FCE)"],
    "age_55_59": ["Age 55-59", "Age 55-59 (FCE)"],
    "age_60_64": ["Age 60-64", "Age 60-64 (FCE)"],
    "age_65_69": ["Age 65-69", "Age 65-69 (FCE)"],
    "age_70_74": ["Age 70-74", "Age 70-74 (FCE)"],
    "age_75_79": ["Age 75-79", "Age 75-79 (FCE)"],
    "age_80_84": ["Age 80-84", "Age 80-84 (FCE)"],
    "age_85_89": ["Age 85-89", "Age 85-89 (FCE)"],
    "age_90_plus": ["Age 90+", "Age 90+ (FCE)"],
}

# User-friendly labels shown in the sidebar instead of raw column names.
AGE_LABELS = {
    "age_0": "0",
    "age_1_4": "1 to 4",
    "age_5_9": "5 to 9",
    "age_10_14": "10 to 14",
    "age_15": "15",
    "age_16": "16",
    "age_17": "17",
    "age_18": "18",
    "age_19": "19",
    "age_20_24": "20 to 24",
    "age_25_29": "25 to 29",
    "age_30_34": "30 to 34",
    "age_35_39": "35 to 39",
    "age_40_44": "40 to 44",
    "age_45_49": "45 to 49",
    "age_50_54": "50 to 54",
    "age_55_59": "55 to 59",
    "age_60_64": "60 to 64",
    "age_65_69": "65 to 69",
    "age_70_74": "70 to 74",
    "age_75_79": "75 to 79",
    "age_80_84": "80 to 84",
    "age_85_89": "85 to 89",
    "age_90_plus": "90+",
}


def clean_text(value):
    """Clean Excel cell text so headers and labels are easier to work with."""
    # Blank Excel cells are returned as None so they can be skipped later.
    if pd.isna(value):
        return None
    # Remove line breaks and repeated spaces because NHS headers vary across years.
    value = str(value).replace("\n", " ").strip()
    value = re.sub(r"\s+", " ", value)
    return value if value else None


def dedupe(items):
    """Make repeated column names unique after reading Excel headers."""
    # Some tables have repeated names, so a counter is added to duplicates.
    seen = {}
    result = []
    for item in items:
        base = item or "unnamed"
        n = seen.get(base, 0)
        result.append(base if n == 0 else f"{base}_{n}")
        seen[base] = n + 1
    return result


def parse_year_label(path_str):
    """Get the financial year label from the NHS workbook file name."""
    # Example file name contains 2020-21, which becomes the displayed year label.
    name = os.path.basename(path_str)
    match = re.search(r"(20\d{2})-(\d{2})", name)
    if not match:
        raise ValueError(f"Cannot parse year from {name}")
    return f"{match.group(1)}-{match.group(2)}"


def year_start(year_label):
    """Return the first calendar year so the financial years sort properly."""
    return int(year_label[:4])


def resolve_sheet(xl, abstraction):
    """Find the correct worksheet name for the selected diagnosis level."""
    # The same sheet has slightly different names in older and newer workbooks.
    for candidate in SHEET_ALIASES[abstraction]:
        if candidate in xl.sheet_names:
            return candidate
    raise ValueError(f"Missing sheet for {abstraction}: {xl.sheet_names}")


def parse_sheet(path_str, abstraction):
    """Read one worksheet and keep only valid diagnosis rows and metric columns."""
    # Open the workbook and select the correct diagnosis sheet.
    xl = pd.ExcelFile(path_str)
    sheet_name = resolve_sheet(xl, abstraction)
    raw = pd.read_excel(path_str, sheet_name=sheet_name, header=None)

    # Find the real table header. NHS files have title notes above the data table.
    header_idx = None
    metric_start = None
    for i, row in raw.iterrows():
        mask = row.astype(str).str.contains("Finished consultant episodes", case=False, na=False)
        if mask.any():
            header_idx = i
            metric_start = int(mask[mask].index[0])
            break
    if header_idx is None:
        raise ValueError(f"Could not locate metric header in {path_str} {sheet_name}")

    # Clean metric column names and start reading rows below the header area.
    metric_headers = dedupe([clean_text(v) or "" for v in raw.iloc[header_idx, metric_start:].tolist()])
    records = []
    data = raw.iloc[header_idx + 2 :].copy()

    # Convert each valid spreadsheet row into one clean Python dictionary.
    for _, row in data.iterrows():
        leading = [clean_text(row[c]) for c in range(metric_start) if clean_text(row[c]) is not None]
        if not leading:
            continue
        code = leading[0]
        description = leading[1] if len(leading) > 1 else None

        # Keep only ICD-like codes, which removes notes, copyright text and blank rows.
        valid_patterns = {
            "summary": r"^[A-Z][0-9]{2}(?:-[A-Z]?[0-9]{2})?$",
            "3char": r"^[A-Z][0-9]{2}$",
            "4char": r"^[A-Z][0-9]{2}\.[0-9A-Z]$",
        }
        if code != "Total" and re.match(valid_patterns[abstraction], str(code)) is None:
            continue

        # Store core identifiers first, then add all numeric measures.
        rec = {
            "year": parse_year_label(path_str),
            "year_start": year_start(parse_year_label(path_str)),
            "abstraction_level": abstraction,
            "code": str(code).strip(),
            "description": description,
        }
        for col_idx, col_name in zip(range(metric_start, raw.shape[1]), metric_headers):
            rec[col_name] = row[col_idx]
        records.append(rec)

    return pd.DataFrame(records)


def coalesce_numeric(df, candidates):
    """Use the first matching numeric column from a list of possible NHS column names."""
    # This handles cases where a column name changed slightly between years.
    existing = [c for c in candidates if c in df.columns]
    if not existing:
        return pd.Series(np.nan, index=df.index)
    result = pd.to_numeric(df[existing[0]], errors="coerce")
    for col in existing[1:]:
        result = result.fillna(pd.to_numeric(df[col], errors="coerce"))
    return result


def is_valid_3char_code(value):
    """Check that a value is a proper ICD 3-character code like A00 or J18."""
    # Only codes with one letter and two digits should be used for hierarchy matching.
    if value is None or pd.isna(value):
        return False
    return re.match(r"^[A-Z][0-9]{2}$", str(value).strip().upper()) is not None


def code_to_ord(code3):
    """Turn a 3-character ICD code into a sortable number for hierarchy matching."""
    code3 = str(code3).strip().upper()
    if not is_valid_3char_code(code3):
        return None
    return ord(code3[0]) * 1000 + int(code3[1:3])


def summary_range_to_bounds(summary_code):
    """Convert summary code ranges like A00-A09 into start and end positions."""
    summary_code = str(summary_code).strip().upper()
    if "-" not in summary_code:
        start = summary_code
        end = summary_code
    else:
        start, end = summary_code.split("-", 1)
        if len(end) == 2:
            end = f"{start[0]}{end}"

    start_ord = code_to_ord(start)
    end_ord = code_to_ord(end)
    if start_ord is None or end_ord is None:
        return None, None
    return start_ord, end_ord


def attach_hierarchy(df):
    """Link summary, 3-character, and 4-character diagnosis levels together."""
    # Work on a copy so the original loaded data is not accidentally changed.
    df = df.copy()
    df["summary_code"] = np.where(df["abstraction_level"] == "summary", df["code"], None)
    df["summary_desc"] = np.where(df["abstraction_level"] == "summary", df["description"], None)

    # 4-character codes belong under their first 3 characters, for example M54.5 -> M54.
    df["code_3char"] = None
    df.loc[df["abstraction_level"] == "3char", "code_3char"] = df.loc[df["abstraction_level"] == "3char", "code"]
    df.loc[df["abstraction_level"] == "4char", "code_3char"] = df.loc[df["abstraction_level"] == "4char", "code"].astype(str).str[:3]
    df["code_3char"] = df["code_3char"].where(df["code_3char"].apply(is_valid_3char_code), None)

    # Build a lookup table that maps each 3-character code to its summary range.
    summaries = df[df["abstraction_level"] == "summary"][["year", "code", "description"]].dropna(subset=["code"])
    map_rows = []

    for year, g in summaries.groupby("year"):
        ranges = []
        for _, row in g.iterrows():
            lo, hi = summary_range_to_bounds(row["code"])
            if lo is not None and hi is not None:
                ranges.append((lo, hi, row["code"], row["description"]))

        year_3chars = sorted(
            c for c in df[(df["year"] == year) & (df["abstraction_level"] == "3char")]["code"].dropna().unique()
            if is_valid_3char_code(c)
        )

        for c3 in year_3chars:
            key = code_to_ord(c3)
            if key is None:
                continue

            parent_code = None
            parent_desc = None
            for lo, hi, scode, sdesc in ranges:
                if lo <= key <= hi:
                    parent_code = scode
                    parent_desc = sdesc
                    break

            map_rows.append({
                "year": year,
                "code_3char": c3,
                "summary_code_map": parent_code,
                "summary_desc_map": parent_desc,
            })

    map_df = pd.DataFrame(map_rows)

    # Add the summary parent back onto the main table.
    if not map_df.empty:
        df = df.merge(map_df, on=["year", "code_3char"], how="left")
        df["summary_code"] = df["summary_code"].fillna(df["summary_code_map"])
        df["summary_desc"] = df["summary_desc"].fillna(df["summary_desc_map"])
        df.drop(columns=["summary_code_map", "summary_desc_map"], inplace=True)

    # Add readable 3-character labels so the treemap hierarchy is understandable.
    threechar_lookup = df[df["abstraction_level"] == "3char"][["year", "code", "description"]].rename(
        columns={"code": "code_3char", "description": "desc_3char"}
    )
    threechar_lookup = threechar_lookup[threechar_lookup["code_3char"].apply(is_valid_3char_code)]
    df = df.merge(threechar_lookup.drop_duplicates(), on=["year", "code_3char"], how="left")

    # Final labels combine code and description for clear chart text.
    df["summary_label"] = np.where(
        df["summary_code"].notna(),
        df["summary_code"].astype(str) + " " + df["summary_desc"].fillna(""),
        None,
    )
    df["label"] = np.where(
        df["description"].notna(),
        df["code"].astype(str) + " " + df["description"].astype(str),
        df["code"].astype(str),
    )
    df["label_3char"] = np.where(
        df["code_3char"].notna() & df["desc_3char"].notna(),
        df["code_3char"].astype(str) + " " + df["desc_3char"].astype(str),
        df["code_3char"],
    )
    return df


@st.cache_data(show_spinner=True)
def load_dataset(data_dir_str):
    """Load all NHS Excel files once, clean them, and return dashboard-ready tables."""
    # Fast path: if a preprocessed pickle exists, load it instead of reading all Excel files.
    # This makes local runs and Streamlit Cloud cold starts much quicker after the cache file is created.
    cache_path = Path(__file__).resolve().parent / "nhs_dashboard_processed.pkl"
    if cache_path.exists():
        try:
            cached = pd.read_pickle(cache_path)
            return cached["tidy"], cached["age_long"]
        except Exception:
            pass

    # Find all yearly .xlsx files in the selected data folder.
    data_dir = Path(data_dir_str)
    files = sorted(glob.glob(str(data_dir / "hosp-epis-stat-admi-diag-20*-tab*.xlsx")))
    if not files:
        raise FileNotFoundError("No matching xlsx files found")

    # Load summary, 3-character and 4-character sheets for every year.
    parsed = []
    for f in files:
        for abstraction in ["summary", "3char", "4char"]:
            parsed.append(parse_sheet(f, abstraction))
    raw = pd.concat(parsed, ignore_index=True)

    # Create one consistent table with standard column names used by the dashboard.
    tidy = raw[["year", "year_start", "abstraction_level", "code", "description"]].copy()
    tidy["fce"] = coalesce_numeric(raw, ["Finished consultant episodes", "Finished Consultant Episodes"])
    tidy["fae"] = coalesce_numeric(raw, ["Finished Admission Episodes", "Admissions"])
    tidy["male_fce"] = coalesce_numeric(raw, ["Male (FCE)", "Male"])
    tidy["female_fce"] = coalesce_numeric(raw, ["Female (FCE)", "Female"])
    tidy["gender_unknown_fce"] = coalesce_numeric(raw, ["Gender Unknown (FCE)", "Gender Unknown", "Gender unknown"])
    tidy["emergency_fae"] = coalesce_numeric(raw, ["Emergency (FAE)", "Emergency"])
    tidy["waiting_list_fae"] = coalesce_numeric(raw, ["Waiting list (FAE)", "Waiting list", "Waiting List"])
    tidy["planned_fae"] = coalesce_numeric(raw, ["Planned (FAE)", "Planned"])
    tidy["other_fae"] = coalesce_numeric(raw, ["Other (FAE)", "Other Admission Method"])
    tidy["mean_wait_days"] = coalesce_numeric(raw, ["Mean time waited (Days)", "Mean time waited"])
    tidy["median_wait_days"] = coalesce_numeric(raw, ["Median time waited (Days)", "Median time waited"])
    tidy["mean_los_days"] = coalesce_numeric(raw, ["Mean length of stay (Days)", "Mean length of stay"])
    tidy["median_los_days"] = coalesce_numeric(raw, ["Median length of stay (Days)", "Median length of stay"])
    tidy["mean_age"] = coalesce_numeric(raw, ["Mean age (Years)", "Mean age"])
    tidy["day_case_fce"] = coalesce_numeric(raw, ["Day case (FCE)", "Day case"])
    tidy["bed_days"] = coalesce_numeric(raw, ["FCE bed days", "Bed Days"])

    # Add all age columns using the same internal names.
    for canonical, candidates in AGE_CANONICAL.items():
        tidy[canonical] = coalesce_numeric(raw, candidates)

    # Add hierarchy fields and remove the overall Total row.
    tidy = attach_hierarchy(tidy)
    tidy = tidy[tidy["code"] != "Total"].copy()
    # Derived measures used by the charts.
    tidy["emergency_share"] = np.where(tidy["fae"] > 0, tidy["emergency_fae"] / tidy["fae"], np.nan)
    tidy["planned_share"] = np.where(tidy["fae"] > 0, tidy["planned_fae"] / tidy["fae"], np.nan)
    tidy["waiting_share"] = np.where(tidy["fae"] > 0, tidy["waiting_list_fae"] / tidy["fae"], np.nan)
    tidy["bed_days_per_fce"] = np.where(tidy["fce"] > 0, tidy["bed_days"] / tidy["fce"], np.nan)

    # Reshape age columns into long format for the violin plot.
    age_cols = list(AGE_CANONICAL.keys())
    age_long = tidy.melt(
        id_vars=["year", "year_start", "abstraction_level", "code", "description", "label", "summary_label", "label_3char", "fce", "fae"],
        value_vars=age_cols,
        var_name="age_group",
        value_name="age_fce",
    )
    age_long["age_group_label"] = age_long["age_group"].map(AGE_LABELS)
    age_long["age_share_of_category"] = np.where(age_long["fce"] > 0, age_long["age_fce"] / age_long["fce"], np.nan)

    # Save the processed tables for faster future starts.
    # If the file cannot be written, the dashboard still works normally.
    try:
        pd.to_pickle({"tidy": tidy, "age_long": age_long}, cache_path)
    except Exception:
        pass

    return tidy, age_long


def filter_base(df, years, abstraction):
    """Apply the year and abstraction level filters from the sidebar."""
    return df[(df["year"].isin(years)) & (df["abstraction_level"] == abstraction)].copy()


def add_selected_age_metric(df, selected_age_cols):
    """Add one combined age burden column based on the selected age groups."""
    if not selected_age_cols:
        df["selected_age_fce"] = df["fce"]
    else:
        df["selected_age_fce"] = df[selected_age_cols].sum(axis=1, skipna=True)
    return df


def metric_series_name(metric_option):
    """Match the sidebar metric name to the actual dataframe column."""
    mapping = {
        "Finished Consultant Episodes": "fce",
        "Finished Admission Episodes": "fae",
        "Emergency Admissions": "emergency_fae",
        "Bed Days": "bed_days",
        "Selected Age FCE": "selected_age_fce",
    }
    return mapping[metric_option]


def format_int(x):
    """Format large numbers with commas for the KPI cards."""
    return f"{int(round(x)):,}" if pd.notna(x) else "NA"


def make_short_label(value, max_len=42):
    """Shorten long diagnosis labels so charts stay readable."""
    if value is None or pd.isna(value):
        return "Unknown"
    value = str(value).strip()
    return value if len(value) <= max_len else value[: max_len - 1] + "…"


# Main dashboard title.
st.title("NHS Hospital Admissions Research Dashboard")
st.caption("Interactive analysis of NHS hospital admissions using the consistent 2012-13 to 2023-24 Excel releases with detailed age breakdowns across summary, 3-character, and 4-character diagnosis levels.")

# Sidebar contains all user controls for filtering and selecting metrics.
with st.sidebar:
    st.header("Controls")
    # Default to the data folder placed next to this app file.
    local_data_path = Path(__file__).resolve().parent / "NHS Hospital Admissions"
    default_path = str(local_data_path if local_data_path.exists() else Path.cwd())
    data_dir = st.text_input("Data folder", value=default_path)
    tidy, age_long = load_dataset(data_dir)

    available_years = sorted(tidy["year"].unique())
    selected_years = st.multiselect("Years", available_years, default=available_years)
    abstraction_label = st.selectbox("Abstraction level", ["summary", "3char", "4char"], index=0, format_func=lambda x: {"summary":"Summary diagnosis", "3char":"3-character diagnosis", "4char":"4-character diagnosis"}[x])
    selected_age_cols = st.multiselect("Age categories", list(AGE_LABELS.keys()), default=["age_65_69", "age_70_74", "age_75_79", "age_80_84", "age_85_89", "age_90_plus"], format_func=lambda x: AGE_LABELS[x])
    metric_option = st.selectbox("Treemap size metric", ["Finished Consultant Episodes", "Finished Admission Episodes", "Emergency Admissions", "Bed Days", "Selected Age FCE"], index=4)
    top_n = st.slider("Top N categories", 5, 40, 15)

# Apply sidebar filters before building the visualizations.
base = filter_base(tidy, selected_years, abstraction_label)
base = add_selected_age_metric(base, selected_age_cols)
metric_col = metric_series_name(metric_option)

if base.empty:
    st.warning("No records found for the selected filters")
    st.stop()

# Most charts use the latest selected year as the current snapshot.
latest_year = max(selected_years, key=year_start)
latest = base[base["year"] == latest_year].copy()
latest = latest[latest[metric_col].fillna(0) > 0].copy()

# KPI cards give a quick summary of the selected dataset.
k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f"<div class='metric-card'><b>Selected abstraction</b><br>{abstraction_label}<br><br><b>Latest year</b><br>{latest_year}</div>", unsafe_allow_html=True)
with k2:
    st.markdown(f"<div class='metric-card'><b>Total FCE</b><br>{format_int(latest['fce'].sum())}<br><br><b>Total FAE</b><br>{format_int(latest['fae'].sum())}</div>", unsafe_allow_html=True)
with k3:
    emergency_rate = latest["emergency_fae"].sum() / latest["fae"].sum() if latest["fae"].sum() > 0 else np.nan
    st.markdown(f"<div class='metric-card'><b>Total emergency admissions</b><br>{format_int(latest['emergency_fae'].sum())}<br><br><b>Emergency share</b><br>{emergency_rate:.1%}</div>", unsafe_allow_html=True)
with k4:
    st.markdown(f"<div class='metric-card'><b>Selected age burden</b><br>{format_int(latest['selected_age_fce'].sum())}<br><br><b>Mean stay</b><br>{latest['mean_los_days'].mean():.2f} days</div>", unsafe_allow_html=True)

# Chart 1: Treemap for hierarchical overview.
st.subheader("1. Hierarchical burden overview")
tree_df = latest.copy()
tree_df = tree_df.replace([np.inf, -np.inf], np.nan)
tree_df[metric_col] = pd.to_numeric(tree_df[metric_col], errors="coerce").fillna(0)
tree_df["emergency_share"] = pd.to_numeric(tree_df["emergency_share"], errors="coerce").clip(lower=0, upper=1).fillna(0)
tree_df = tree_df[tree_df[metric_col] > 0].copy()

# Keep short labels inside the boxes, but keep full labels for hover.
tree_df["short_label"] = tree_df["label"].apply(lambda x: make_short_label(x, 48))
tree_df["full_label"] = tree_df["label"].fillna(tree_df["code"]).fillna("Unknown category")
tree_df["summary_full_label"] = tree_df["summary_label"].fillna("Unmapped summary")
tree_df["summary_short_label"] = tree_df["summary_full_label"].apply(lambda x: make_short_label(x, 48))
tree_df["three_full_label"] = tree_df["label_3char"].fillna(tree_df["code_3char"]).fillna("Unmapped 3-character group")
tree_df["three_short_label"] = tree_df["three_full_label"].apply(lambda x: make_short_label(x, 48))

if tree_df.empty:
    st.warning("No positive values are available for the selected treemap metric and filters.")
else:
    try:
        # Build the treemap manually rather than using px.treemap.
        # This lets the visible box label stay short while the hover label is always full.
        nodes = {}

        def weighted_share(frame):
            values = pd.to_numeric(frame[metric_col], errors="coerce").fillna(0)
            shares = pd.to_numeric(frame["emergency_share"], errors="coerce").fillna(0)
            total = values.sum()
            if total <= 0:
                return 0
            return float((shares * values).sum() / total)

        def add_node(node_id, parent_id, visible_label, full_label, value, colour):
            nodes[node_id] = {
                "id": node_id,
                "parent": parent_id,
                "label": visible_label,
                "full": full_label,
                "value": float(value) if pd.notna(value) else 0,
                "colour": float(colour) if pd.notna(colour) else 0,
            }

        root_id = "root"
        add_node(
            root_id,
            "",
            "All diagnoses",
            "All diagnoses",
            tree_df[metric_col].sum(),
            weighted_share(tree_df),
        )

        if abstraction_label == "summary":
            for _, row in tree_df.iterrows():
                leaf_id = f"leaf|{row['code']}|{row['full_label']}"
                add_node(
                    leaf_id,
                    root_id,
                    row["short_label"],
                    row["full_label"],
                    row[metric_col],
                    row["emergency_share"],
                )

        elif abstraction_label == "3char":
            for summary_full, summary_group in tree_df.groupby("summary_full_label", dropna=False):
                summary_id = f"summary|{summary_full}"
                add_node(
                    summary_id,
                    root_id,
                    make_short_label(summary_full, 48),
                    summary_full,
                    summary_group[metric_col].sum(),
                    weighted_share(summary_group),
                )

                for _, row in summary_group.iterrows():
                    leaf_id = f"leaf|{row['code']}|{row['full_label']}"
                    add_node(
                        leaf_id,
                        summary_id,
                        row["short_label"],
                        row["full_label"],
                        row[metric_col],
                        row["emergency_share"],
                    )

        else:
            for summary_full, summary_group in tree_df.groupby("summary_full_label", dropna=False):
                summary_id = f"summary|{summary_full}"
                add_node(
                    summary_id,
                    root_id,
                    make_short_label(summary_full, 48),
                    summary_full,
                    summary_group[metric_col].sum(),
                    weighted_share(summary_group),
                )

                for three_full, three_group in summary_group.groupby("three_full_label", dropna=False):
                    three_id = f"three|{summary_full}|{three_full}"
                    add_node(
                        three_id,
                        summary_id,
                        make_short_label(three_full, 48),
                        three_full,
                        three_group[metric_col].sum(),
                        weighted_share(three_group),
                    )

                    for _, row in three_group.iterrows():
                        leaf_id = f"leaf|{row['code']}|{row['full_label']}"
                        add_node(
                            leaf_id,
                            three_id,
                            row["short_label"],
                            row["full_label"],
                            row[metric_col],
                            row["emergency_share"],
                        )

        node_list = list(nodes.values())

        fig_tree = go.Figure(
            go.Treemap(
                ids=[n["id"] for n in node_list],
                labels=[n["label"] for n in node_list],
                parents=[n["parent"] for n in node_list],
                values=[n["value"] for n in node_list],
                branchvalues="total",
                customdata=np.array([[n["full"], n["colour"]] for n in node_list], dtype=object),
                marker=dict(
                    colors=[n["colour"] for n in node_list],
                    colorscale="RdBu_r",
                    cmin=0,
                    cmax=1,
                    line=dict(width=1, color="white"),
                    colorbar=dict(title="Emergency share"),
                ),
                textinfo="label",
                textfont=dict(size=15),
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Burden: %{value:,.0f}<br>"
                    "Emergency share: %{customdata[1]:.2%}"
                    "<extra></extra>"
                ),
            )
        )

        fig_tree.update_layout(
            template=PLOTLY_TEMPLATE,
            height=740,
            margin=dict(t=30, l=10, r=10, b=10),
            uniformtext=dict(minsize=11, mode="hide"),
        )

        st.plotly_chart(fig_tree, use_container_width=True)

    except Exception as exc:
        # Fallback keeps the dashboard usable even if Plotly cannot build a treemap.
        st.warning("The treemap could not be drawn for the current filter combination because the hierarchy is incomplete or inconsistent. A ranked fallback view is shown instead.")
        fallback = tree_df.nlargest(min(top_n, len(tree_df)), metric_col).copy()
        fallback["fallback_label"] = fallback["label"].apply(lambda x: make_short_label(x, 72))

        fig_fallback = px.bar(
            fallback,
            x=metric_col,
            y="fallback_label",
            orientation="h",
            color="emergency_share",
            color_continuous_scale="RdBu_r",
            template=PLOTLY_TEMPLATE,
            custom_data=["label", "emergency_share"],
        )

        fig_fallback.update_traces(
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Burden: %{x:,.0f}<br>"
                "Emergency share: %{customdata[1]:.2%}"
                "<extra></extra>"
            )
        )

        fig_fallback.update_layout(
            height=max(440, 45 * len(fallback) + 120),
            margin=dict(l=160, r=30, t=30, b=30),
            yaxis_title="",
            xaxis_title=metric_option,
        )

        st.plotly_chart(fig_fallback, use_container_width=True)

st.subheader("2. Change vs pre COVID baseline")

baseline_year = "2019-20"
lockdown_year = "2020-21"

# The same heatmap can show admissions change or emergency share change.
heatmap_metric = st.selectbox(
    "Heatmap metric",
    [
        "Admissions % change from 2019-20 baseline",
        "Emergency share percentage point change from 2019-20 baseline",
    ],
    index=0,
)

# Fixed research set for the lockdown anomaly story. Labels are deliberately concise
# so the heatmap remains readable and rows do not duplicate when NHS wording changes across years.
focus_codes = [
    "U00-U49", "O10-O16", "F20-F29", "F10-F19", "F50-F59", "I20-I25",
    "I10-I15", "C81-C96", "C15-C26", "N17-N19", "K00-K14", "K55-K64",
    "M00-M25", "M40-M54", "H65-H75", "H10-H13", "H25-H28", "J20-J22",
    "J00-J06", "J30-J39", "J09-J18", "A00-A09", "E65-E68", "S00-S09",
]

focus_label_map = {
    "U00-U49": "U00-U49 COVID-19 codes",
    "O10-O16": "O10-O16 Pregnancy hypertension",
    "F20-F29": "F20-F29 Schizophrenia and psychosis",
    "F10-F19": "F10-F19 Alcohol and substance disorders",
    "F50-F59": "F50-F59 Eating and behavioural syndromes",
    "I20-I25": "I20-I25 Ischaemic heart disease",
    "I10-I15": "I10-I15 Hypertensive diseases",
    "C81-C96": "C81-C96 Blood and lymphatic cancers",
    "C15-C26": "C15-C26 Digestive organ cancers",
    "N17-N19": "N17-N19 Renal failure",
    "K00-K14": "K00-K14 Oral cavity and salivary diseases",
    "K55-K64": "K55-K64 Other intestinal diseases",
    "M00-M25": "M00-M25 Arthropathies",
    "M40-M54": "M40-M54 Dorsopathies and back pain",
    "H65-H75": "H65-H75 Middle ear and mastoid diseases",
    "H10-H13": "H10-H13 Conjunctivitis and related eye disorders",
    "H25-H28": "H25-H28 Cataracts and lens disorders",
    "J20-J22": "J20-J22 Acute lower respiratory infections",
    "J00-J06": "J00-J06 Acute upper respiratory infections",
    "J30-J39": "J30-J39 Other upper respiratory diseases",
    "J09-J18": "J09-J18 Influenza and pneumonia",
    "A00-A09": "A00-A09 Intestinal infectious diseases",
    "E65-E68": "E65-E68 Obesity and hyperalimentation",
    "S00-S09": "S00-S09 Head injuries",
}

# A balanced diverging palette: red for decrease, warm neutral near baseline, blue for increase.
# The palette is intentionally muted and publication friendly on a white dashboard.
baseline_colour_scale = [
    [0.00, "#7A0026"],
    [0.16, "#B12A3B"],
    [0.34, "#E9856C"],
    [0.50, "#F2EFE8"],
    [0.66, "#8AC1D4"],
    [0.84, "#2166A5"],
    [1.00, "#08306B"],
]

try:
    # Use summary level for this heatmap because the main findings are summary codes.
    heat_source = tidy[
        (tidy["abstraction_level"] == "summary")
        & (tidy["year"].isin(available_years))
        & (tidy["code"].isin(focus_codes))
    ].copy()

    if heat_source.empty:
        st.warning("No summary diagnosis records are available for the selected baseline heatmap categories.")
    else:
        # Collapse by code and year before plotting. This prevents duplicate rows when the
        # category wording changes slightly across NHS releases.
        heat_source = heat_source.groupby(["code", "year", "year_start"], as_index=False).agg(
            fce=("fce", "sum"),
            fae=("fae", "sum"),
            emergency_fae=("emergency_fae", "sum"),
        )
        heat_source["emergency_share"] = np.where(
            heat_source["fae"] > 0,
            heat_source["emergency_fae"] / heat_source["fae"],
            np.nan,
        )
        heat_source["category_order"] = heat_source["code"].apply(lambda x: focus_codes.index(x) if x in focus_codes else 999)
        heat_source["category_label"] = heat_source["code"].map(focus_label_map).fillna(heat_source["code"])

        # Pull out 2019-20 values to use as the baseline for each category.
        baseline = heat_source[heat_source["year"] == baseline_year][[
            "code", "fce", "fae", "emergency_share"
        ]].rename(columns={
            "fce": "baseline_fce",
            "fae": "baseline_fae",
            "emergency_share": "baseline_emergency_share",
        })

        heat_calc = heat_source.merge(baseline, on="code", how="left")

        # Calculate the cell value depending on which metric the user selected.
        if heatmap_metric.startswith("Admissions"):
            heat_calc["heat_value"] = np.where(
                heat_calc["baseline_fae"] > 0,
                (heat_calc["fae"] - heat_calc["baseline_fae"]) / heat_calc["baseline_fae"] * 100,
                np.nan,
            )
            color_title = "% change vs 2019-20 baseline"
            hover_value_label = "Change vs baseline"
            zmin, zmax = -80, 200
        else:
            heat_calc["heat_value"] = (
                heat_calc["emergency_share"] - heat_calc["baseline_emergency_share"]
            ) * 100
            color_title = "Emergency share change, percentage points"
            hover_value_label = "Emergency share change"
            zmin, zmax = -20, 20

        # The baseline year is the reference point, therefore it is neutral by definition.
        heat_calc.loc[heat_calc["year"] == baseline_year, "heat_value"] = 0
        heat_calc["heat_value"] = heat_calc["heat_value"].replace([np.inf, -np.inf], np.nan)
        heat_calc = heat_calc.sort_values(["category_order", "year_start"])

        y_order = [focus_label_map[c] for c in focus_codes if c in set(heat_calc["code"])]
        x_order = [yr for yr in available_years if yr in set(heat_calc["year"])]
        x_order = sorted(x_order, key=year_start)

        # Convert long data into a matrix that the heatmap can plot.
        heat_matrix = heat_calc.pivot_table(
            index="category_label",
            columns="year",
            values="heat_value",
            aggfunc="mean",
            sort=False,
        ).reindex(index=y_order, columns=x_order)

        x_vals = list(range(len(x_order)))
        y_vals = list(range(len(y_order)))
        z_values = heat_matrix.to_numpy(dtype=float)

        # Custom hover data mirrors the z matrix.
        custom_records = []
        heat_lookup = heat_calc.set_index(["category_label", "year"])
        for category in y_order:
            row_records = []
            for yr in x_order:
                if (category, yr) in heat_lookup.index:
                    r = heat_lookup.loc[(category, yr)]
                    if isinstance(r, pd.DataFrame):
                        r = r.iloc[0]
                    row_records.append([
                        r.get("code", ""),
                        yr,
                        r.get("fae", np.nan),
                        r.get("baseline_fae", np.nan),
                        r.get("emergency_share", np.nan),
                        r.get("baseline_emergency_share", np.nan),
                    ])
                else:
                    row_records.append(["", yr, np.nan, np.nan, np.nan, np.nan])
            custom_records.append(row_records)

        # Build the heatmap manually so gaps and cell spacing are easier to control.
        fig_heat = go.Figure()

        # Grey base layer makes unavailable cells explicit instead of leaving holes.
        fig_heat.add_trace(go.Heatmap(
            z=np.zeros_like(z_values, dtype=float),
            x=x_vals,
            y=y_vals,
            colorscale=[[0, "#E6E8EB"], [1, "#E6E8EB"]],
            showscale=False,
            hoverinfo="skip",
            xgap=2,
            ygap=2,
        ))

        fig_heat.add_trace(go.Heatmap(
            z=z_values,
            x=x_vals,
            y=y_vals,
            coloraxis="coloraxis",
            customdata=np.array(custom_records, dtype=object),
            hoverongaps=False,
            xgap=2,
            ygap=2,
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Year: %{customdata[1]}<br>"
                "Code: %{customdata[0]}<br>"
                "Admissions: %{customdata[2]:,.0f}<br>"
                "2019-20 baseline admissions: %{customdata[3]:,.0f}<br>"
                f"{hover_value_label}: %{{z:.1f}}<br>"
                "Emergency share: %{customdata[4]:.1%}<br>"
                "Baseline emergency share: %{customdata[5]:.1%}"
                "<extra></extra>"
            ),
        ))

        # Highlight the lockdown year so it is easy to locate during the video.
        if lockdown_year in x_order:
            lockdown_index = x_order.index(lockdown_year)
            fig_heat.add_shape(
                type="rect",
                x0=lockdown_index - 0.5,
                x1=lockdown_index + 0.5,
                y0=-0.5,
                y1=len(y_order) - 0.5,
                line=dict(color="#F2A900", width=3),
                fillcolor="rgba(0,0,0,0)",
                layer="above",
            )
            fig_heat.add_annotation(
                x=lockdown_index,
                y=-0.9,
                text="Lockdown year 2020-21",
                showarrow=False,
                font=dict(size=11, color="#F2A900"),
                bgcolor="rgba(14,17,23,0.92)",
                bordercolor="#F2A900",
                borderwidth=1,
            )

        # Layout settings keep labels black and readable on the white background.
        fig_heat.update_layout(
            height=790,
            margin=dict(l=275, r=65, t=35, b=90),
            paper_bgcolor="#0E1117",
            plot_bgcolor="#0E1117",
            font=dict(color="white"),
            coloraxis=dict(
                colorscale=baseline_colour_scale,
                cmin=zmin,
                cmax=zmax,
                colorbar=dict(
                    title=dict(text=color_title, font=dict(color="white")),
                    tickfont=dict(color="white"),
                    len=0.78,
                ),
            ),
            xaxis=dict(
                title=dict(text="Financial year", font=dict(color="white")),
                tickmode="array",
                tickvals=x_vals,
                ticktext=x_order,
                tickangle=-35,
                tickfont=dict(color="white"),
                showgrid=False,
                zeroline=False,
            ),
            yaxis=dict(
                title=dict(text="Primary diagnosis summary category", font=dict(color="white")),
                tickmode="array",
                tickvals=y_vals,
                ticktext=y_order,
                tickfont=dict(size=11, color="white"),
                autorange="reversed",
                showgrid=False,
                zeroline=False,
            ),
        )

        st.plotly_chart(fig_heat, use_container_width=True)

        # Small automatic insight line below the heatmap.
        admissions_view = heat_calc[heat_calc["year"] == lockdown_year].copy()
        if heatmap_metric.startswith("Admissions"):
            f_row = admissions_view[admissions_view["code"] == "F50-F59"]
            m_row = admissions_view[admissions_view["code"] == "M00-M25"]
            if not f_row.empty and not m_row.empty:
                st.info(
                    f"Lockdown signal: F50-F59 changed by {f_row['heat_value'].iloc[0]:.1f}% versus 2019-20, "
                    f"while M00-M25 changed by {m_row['heat_value'].iloc[0]:.1f}%."
                )
        else:
            m_row = admissions_view[admissions_view["code"] == "M00-M25"]
            if not m_row.empty:
                st.info(
                    f"Emergency pressure signal: M00-M25 changed by {m_row['heat_value'].iloc[0]:.1f} percentage points in emergency share during 2020-21 versus 2019-20."
                )
except Exception as exc:
    # Do not crash the whole dashboard if this section has a data issue.
    st.warning("The baseline heatmap could not be drawn for the current data combination. Try refreshing the dashboard or checking that the 2019-20 and 2020-21 summary workbooks are present.")

# Chart 3: Radar chart for comparing categories across several measures.
st.subheader("3. Emergency burden profile radar chart")
radar_source = latest[latest["fae"] > 0].nlargest(min(top_n, 8), metric_col).copy()
radar_choices = radar_source["label"].tolist()
default_radar = radar_choices[: min(5, len(radar_choices))]
selected_radar = st.multiselect("Choose up to 5 diagnosis categories for radar comparison", radar_choices, default=default_radar)
radar_df = radar_source[radar_source["label"].isin(selected_radar)].copy()
radar_metrics = ["selected_age_fce", "emergency_share", "mean_wait_days", "mean_los_days", "mean_age"]
radar_names = {
    "selected_age_fce": "Selected age burden",
    "emergency_share": "Emergency share",
    "mean_wait_days": "Mean wait",
    "mean_los_days": "Mean stay",
    "mean_age": "Mean age",
}
if not radar_df.empty:
    # Normalise values to 0-1 because the radar axes use different units.
    norm = radar_df.copy()
    for m in radar_metrics:
        mmin = norm[m].min()
        mmax = norm[m].max()
        norm[m] = 0.5 if pd.isna(mmin) or pd.isna(mmax) or mmin == mmax else (norm[m] - mmin) / (mmax - mmin)
    fig_radar = go.Figure()
    theta = [radar_names[m] for m in radar_metrics]
    for _, row in norm.iterrows():
        r = [row[m] for m in radar_metrics]
        fig_radar.add_trace(go.Scatterpolar(r=r + [r[0]], theta=theta + [theta[0]], fill="toself", name=row["label"]))
    fig_radar.update_layout(template=PLOTLY_TEMPLATE, height=600, polar=dict(radialaxis=dict(visible=True, range=[0, 1])))
    st.plotly_chart(fig_radar, use_container_width=True)

# Chart 4: Violin plot for age-group distribution.
st.subheader("4. Age disparity violin plot")
age_view = age_long[(age_long["year"] == latest_year) & (age_long["abstraction_level"] == abstraction_label)].copy()
focus_labels = latest.nlargest(top_n, metric_col)["label"].tolist()
age_view = age_view[age_view["label"].isin(focus_labels)]
age_view = age_view[age_view["age_fce"].fillna(0) > 0]
# Violin shape shows distribution across selected diagnosis categories.
fig_violin = px.violin(
    age_view,
    x="age_group_label",
    y="age_share_of_category",
    color="age_group_label",
    box=True,
    points=False,
    template=PLOTLY_TEMPLATE,
)
fig_violin.update_layout(height=650, xaxis_title="Age group", yaxis_title="Share of category FCE")
st.plotly_chart(fig_violin, use_container_width=True)
if not age_view.empty:
    median_age_share = age_view.groupby("age_group_label")["age_share_of_category"].median().sort_values(ascending=False)
    if not median_age_share.empty:
        st.info(f"Current filtered insight: {median_age_share.index[0]} has the highest median share across the selected diagnosis categories.")

# Chart 5: Parallel coordinates for multivariable comparison.
st.subheader("5. Parallel coordinates for multivariate comparison")
parallel_df = latest.nlargest(top_n, metric_col).copy()
parallel_df["color_value"] = parallel_df[metric_col]
# Each line is a diagnosis category and each axis is a different variable.
fig_parallel = px.parallel_coordinates(
    parallel_df,
    dimensions=["fce", "fae", "emergency_fae", "mean_wait_days", "mean_los_days", "mean_age", "selected_age_fce"],
    color="color_value",
    labels={
        "fce": "FCE",
        "fae": "FAE",
        "emergency_fae": "Emergency FAE",
        "mean_wait_days": "Mean wait",
        "mean_los_days": "Mean stay",
        "mean_age": "Mean age",
        "selected_age_fce": "Selected age FCE",
    },
    color_continuous_scale=px.colors.sequential.Blues,
)
fig_parallel.update_layout(
    template=PLOTLY_TEMPLATE,
    height=650,
    margin=dict(l=120, r=35, t=35, b=25),
)
st.plotly_chart(fig_parallel, use_container_width=True)

# Chart 6: Supporting stacked chart for admission route composition.
st.subheader("6. Admission route composition")
stack_years = sorted(selected_years, key=year_start)
stack_default_index = stack_years.index(latest_year) if latest_year in stack_years else len(stack_years) - 1
stack_year = st.selectbox("Admission route year", stack_years, index=stack_default_index)

# Use one selected year so labels stay readable and do not overlap.
stack_source = tidy[(tidy["abstraction_level"] == abstraction_label) & (tidy["year"] == stack_year)].copy()
if not selected_radar:
    stack_labels = latest.nlargest(6, metric_col)["label"].tolist()
else:
    stack_labels = selected_radar[:6]
stack_source = stack_source[stack_source["label"].isin(stack_labels)].copy()
stack_source["short_stack_label"] = stack_source["label"].apply(lambda x: make_short_label(x, 64))

# Convert admission route columns into long format for stacked plotting.
stack_long = stack_source.melt(
    id_vars=["year", "label", "short_stack_label"],
    value_vars=["emergency_fae", "waiting_list_fae", "planned_fae", "other_fae"],
    var_name="admission_type",
    value_name="value",
)
stack_long["admission_type"] = stack_long["admission_type"].map({
    "emergency_fae": "Emergency",
    "waiting_list_fae": "Waiting list",
    "planned_fae": "Planned",
    "other_fae": "Other",
})
stack_long["value"] = pd.to_numeric(stack_long["value"], errors="coerce").fillna(0)
stack_long = stack_long[stack_long["value"] >= 0].copy()

if stack_long.empty:
    st.info("No admission route composition values are available for the selected filters.")
else:
    # Sort smaller categories at the bottom and larger categories at the top.
    order = stack_long.groupby("short_stack_label")["value"].sum().sort_values(ascending=True).index.tolist()
    fig_stack = px.bar(
        stack_long,
        x="value",
        y="short_stack_label",
        color="admission_type",
        orientation="h",
        category_orders={"short_stack_label": order},
        barmode="stack",
        template=PLOTLY_TEMPLATE,
        custom_data=["label", "admission_type", "value"],
    )
    fig_stack.update_traces(
        hovertemplate="<b>%{customdata[0]}</b><br>Admission route: %{customdata[1]}<br>Value: %{customdata[2]:,.0f}<extra></extra>"
    )
    fig_stack.update_layout(
        height=max(430, 85 * len(order) + 120),
        margin=dict(l=210, r=30, t=30, b=40),
        legend_title="Admission route",
        xaxis_title="Admissions",
        yaxis_title="",
        yaxis=dict(tickfont=dict(size=12)),
    )
    st.plotly_chart(fig_stack, use_container_width=True)
