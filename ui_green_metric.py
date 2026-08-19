import streamlit as st
import pandas as pd
from pathlib import Path
import re
import html

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="UI GreenMetric Dashboard",
    page_icon="🌱",
    layout="wide",
)

# ============================================================
# FILE PATH
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "data" / "IIUM_Sustainability_Data_Submission_Tracker.xlsx"
HUB_FILE = BASE_DIR / "data" / "sustainability_data_hub.xlsx"

SHEET_NAME = "UI GreenMetric"

# ============================================================
# RANKINGS (placeholder data — replace with real numbers later)
# ============================================================
# Edit COUNTRY_LABEL and the numbers below once real rankings are
# available. Keys are matched case-insensitively against your section
# names; any section without a matching key just won't show a box.

COUNTRY_LABEL = "MALAYSIA RANKINGS"

WORLD_RANKINGS = {
    "overall": 19,
    "setting and infrastructure": 86,
    "energy and climate change": 111,
    "waste": 13,
    "water": 11,
    "transportation": 118,
    "education and research": 15,
    "governance and digitalization": 60,
}

COUNTRY_RANKINGS = {
    "overall": 1,
    "setting and infrastructure": 10,
    "energy and climate change": 15,
    "waste": 1,
    "water": 1,
    "transportation": 9,
    "education and research": 8,
    "governance and digitalization": 3,
}

# Which SDGs each GreenMetric category maps to, per the official UI
# GreenMetric guideline diagram. Matched by keyword against the section
# name (case-insensitive), so it doesn't rely on exact spelling.
SDG_MAP = [
    (("governance", "digitalization"), [4, 5, 8, 10, 16, 17]),
    (("waste",), [3, 6, 11, 12, 14, 15, 17]),
    (("water",), [6, 11, 12, 13, 14, 15, 17]),
    (("energy", "climate"), [7, 9, 11, 12, 13, 17]),
    (("transportation",), [3, 9, 11, 13, 15, 17]),
    (("education", "research"), [4, 5, 8, 10, 12, 15, 17]),
    (("setting", "infrastructure"), [3, 9, 10, 13, 15, 16, 17]),
]


def sdgs_for_section(section_name):
    name_lower = section_name.strip().lower()
    for keywords, sdg_list in SDG_MAP:
        if any(kw in name_lower for kw in keywords):
            return sdg_list
    return []


def sdg_icon_url(n):
    # Official UN SDG icons, hosted publicly on Wikimedia Commons
    # (public domain / PD-UN-doc). Free to use per UN branding guidelines:
    # https://www.un.org/sustainabledevelopment
    return f"https://commons.wikimedia.org/wiki/Special:FilePath/Sustainable%20Development%20Goal%20{n}.png"

# ============================================================
# CUSTOM CSS
# ============================================================

st.html("""
<style>
    .dashboard-title {
        font-size: 2.6rem;
        font-weight: 800;
        color: #14532d;
        margin-bottom: 0.2rem;
    }

    .dashboard-subtitle {
        font-size: 1.05rem;
        color: #4b5563;
        margin-bottom: 1.5rem;
    }

    /* Top 3 summary boxes — dark green */
    .summary-box-dark {
        background: #004700;
        border-radius: 16px;
        padding: 1.4rem;
        text-align: center;
        min-height: 130px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }

    .summary-label-dark {
        font-size: 0.9rem;
        font-weight: 600;
        color: #bbf7d0;
        text-transform: uppercase;
        letter-spacing: 0.03em;
        margin-bottom: 0.4rem;
    }

    .summary-value-dark {
        font-size: 2.2rem;
        font-weight: 800;
        color: #ffffff;
    }

    /* Per-section overview boxes — equal size, white */
    .section-box {
        background: white;
        border-radius: 14px;
        padding: 1.1rem;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        text-align: center;
        min-height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }

    .section-box-label {
        font-size: 0.85rem;
        font-weight: 700;
        color: #6b7280;
        text-transform: uppercase;
        letter-spacing: 0.02em;
        margin-bottom: 0.5rem;
        line-height: 1.3;
    }

    .section-box-value {
        font-size: 1.4rem;
        font-weight: 800;
        color: #166534;
    }

    /* Ranking boxes */
    .rank-box-overall {
        background: #004700;
        border-radius: 14px;
        padding: 1.4rem;
        text-align: center;
        min-height: 110px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }

    .rank-box-overall .rank-value {
        font-size: 2rem;
        font-weight: 800;
        color: #ffffff;
    }

    .rank-box-overall .rank-label {
        font-size: 0.8rem;
        font-weight: 700;
        color: #bbf7d0;
        text-transform: uppercase;
        margin-top: 0.3rem;
    }

    .rank-box {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 1.1rem;
        text-align: center;
        min-height: 110px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }

    .rank-box .rank-value {
        font-size: 1.7rem;
        font-weight: 800;
        color: #14532d;
    }

    .rank-box .rank-label {
        font-size: 0.78rem;
        font-weight: 700;
        color: #6b7280;
        text-transform: uppercase;
        margin-top: 0.3rem;
        line-height: 1.25;
    }

    .rankings-heading {
        font-size: 1.15rem;
        font-weight: 700;
        color: #1f2937;
        margin-top: 1.8rem;
        margin-bottom: 0.8rem;
    }

    /* Question & answer cards inside each tab */
    .qa-card {
        background: white;
        border-radius: 12px;
        padding: 1rem;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        min-height: 170px;
        margin-bottom: 1rem;
    }

    .qa-code {
        font-size: 0.78rem;
        font-weight: 700;
        color: #15803d;
        margin-bottom: 0.3rem;
    }

    .qa-question {
        font-size: 0.88rem;
        font-weight: 600;
        color: #1f2937;
        line-height: 1.4;
        margin-bottom: 0.5rem;
    }

    .qa-answer-label {
        display: none;
    }

    .qa-answer {
        font-size: 1.05rem;
        color: #008000;
        line-height: 1.4;
    }

    /* Add breathing room between tab labels */
    button[data-baseweb="tab"] {
        margin-right: 1.4rem !important;
    }

    .sdg-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.6rem;
        margin: 0.8rem 0 1.4rem 0;
    }

    .sdg-row img {
        width: 58px;
        height: 58px;
        border-radius: 8px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15);
    }

    .sdg-heading {
        font-size: 0.85rem;
        font-weight: 700;
        color: #6b7280;
        text-transform: uppercase;
        letter-spacing: 0.02em;
        margin-top: 0.3rem;
    }
</style>
""")

# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_greenmetric_data():
    df = pd.read_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME,
        header=3,
    )

    # Keep only the four relevant columns
    df = df.iloc[:, :4].copy()
    df.columns = ["No", "Indicator Code", "Question", "Answer"]

    return df


df = load_greenmetric_data()


@st.cache_data
def load_hub_data():
    buildings = pd.read_excel(HUB_FILE, sheet_name="buildings")
    electricity = pd.read_excel(HUB_FILE, sheet_name="electricity_usage")
    water = pd.read_excel(HUB_FILE, sheet_name="water_usage")

    electricity = electricity.merge(buildings, on="building_id", how="left")
    water = water.merge(buildings, on="building_id", how="left")

    for frame in (electricity, water):
        frame["date"] = pd.to_datetime(
            frame["year"].astype(str) + "-" + frame["month"].astype(str) + "-01"
        )

    return buildings, electricity, water


buildings_df, electricity_df, water_df = load_hub_data()

# ============================================================
# PREPARE DATA
# ============================================================

# Section rows have a title in "No" but no Question or Answer
section_mask = (
    df["No"].notna()
    & df["Question"].isna()
    & df["Answer"].isna()
)

# Fill section name down
current_section = None
sections = []

for _, row in df.iterrows():
    if (
        pd.notna(row["No"])
        and pd.isna(row["Question"])
        and pd.isna(row["Answer"])
    ):
        current_section = str(row["No"]).strip()

    sections.append(current_section)

df["Section"] = sections

# Remove section rows
indicator_df = df[~section_mask].copy()

# Ordered list of the sections, in the order they first appear
section_order = list(dict.fromkeys(indicator_df["Section"].dropna().tolist()))

# ============================================================
# FORMATTING HELPERS
# ============================================================

def clean_text(value):
    """Convert Excel values into clean readable text."""
    if pd.isna(value):
        return "No data provided"

    text = str(value).strip()

    if not text:
        return "No data provided"

    # Fix Excel-style inequality symbols
    text = text.replace("<=", "≤")
    text = text.replace(">=", "≥")

    # Replace multiple spaces
    text = re.sub(r"\s+", " ", text)

    return text


def format_answer(value):
    """
    Make answers more readable without changing the actual meaning.
    """
    if pd.isna(value):
        return "No data provided"

    text = str(value).strip()

    if text == "":
        return "No data provided"

    # Replace inequalities
    text = text.replace("<=", "≤")
    text = text.replace(">=", "≥")

    # Convert plain whole-number values such as 25942 -> 25,942
    if re.fullmatch(r"-?\d+", text):
        try:
            number = int(text)
            return f"{number:,}"
        except ValueError:
            pass

    # Convert decimal numbers such as 25000.5 -> 25,000.5
    if re.fullmatch(r"-?\d+\.\d+", text):
        try:
            number = float(text)
            return f"{number:,.2f}".rstrip("0").rstrip(".")
        except ValueError:
            pass

    return text


def qa_card_html(row):
    question = html.escape(clean_text(row["Question"]))
    answer = html.escape(format_answer(row["Answer"]))

    indicator_code = (
        clean_text(row["Indicator Code"])
        if pd.notna(row["Indicator Code"])
        else clean_text(row["No"])
    )
    indicator_code = html.escape(indicator_code)

    return f"""
    <div class="qa-card">
        <div class="qa-code">{indicator_code}</div>
        <div class="qa-question">{question}</div>
        <div class="qa-answer-label">Answer</div>
        <div class="qa-answer">{answer}</div>
    </div>
    """


def rank_lookup(mapping, section_name):
    return mapping.get(section_name.strip().lower())


# ============================================================
# HEADER (title only — no sidebar filters)
# ============================================================

st.markdown(
    '<div class="dashboard-title">🌱 UI GreenMetric Dashboard</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="dashboard-subtitle">UI GreenMetric questionnaire data for IIUM</div>',
    unsafe_allow_html=True,
)

# ============================================================
# TABS
# ============================================================

tab_labels = [name.upper() for name in section_order]
tabs = st.tabs(["OVERVIEW"] + tab_labels)

# ---- Overview tab ----
with tabs[0]:
    total_indicators = len(indicator_df)
    completed_answers = indicator_df["Answer"].notna().sum()
    completion_rate = (
        (completed_answers / total_indicators) * 100
        if total_indicators
        else 0
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            f'<div class="summary-box-dark">'
            f'<div class="summary-label-dark">Total Indicators</div>'
            f'<div class="summary-value-dark">{total_indicators}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    with col2:
        st.markdown(
            f'<div class="summary-box-dark">'
            f'<div class="summary-label-dark">Answered</div>'
            f'<div class="summary-value-dark">{completed_answers}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    with col3:
        st.markdown(
            f'<div class="summary-box-dark">'
            f'<div class="summary-label-dark">Completion</div>'
            f'<div class="summary-value-dark">{completion_rate:.0f}%</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # Per-section boxes, equal size, side by side in one row
    section_cols = st.columns(len(section_order))

    for col, section_name in zip(section_cols, section_order):
        section_data = indicator_df[indicator_df["Section"] == section_name]
        sub_total = len(section_data)
        sub_answered = section_data["Answer"].notna().sum()

        with col:
            st.markdown(
                f'<div class="section-box">'
                f'<div class="section-box-label">{html.escape(section_name.upper())}</div>'
                f'<div class="section-box-value">{sub_answered}/{sub_total}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    # ---- Rankings ----
    st.markdown(f'<div class="rankings-heading">🏆 UI GreenMetric Rankings — World</div>', unsafe_allow_html=True)

    world_overall = WORLD_RANKINGS.get("overall")
    rank_cols = st.columns(len(section_order) + 1)

    with rank_cols[0]:
        st.markdown(
            f'<div class="rank-box-overall">'
            f'<div class="rank-value">#{world_overall}</div>'
            f'<div class="rank-label">Overall</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    for col, section_name in zip(rank_cols[1:], section_order):
        rank = rank_lookup(WORLD_RANKINGS, section_name)
        rank_display = f"#{rank}" if rank is not None else "—"
        with col:
            st.markdown(
                f'<div class="rank-box">'
                f'<div class="rank-value">{rank_display}</div>'
                f'<div class="rank-label">{html.escape(section_name.upper())}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.markdown(f'<div class="rankings-heading">🏆 UI GreenMetric Rankings — {COUNTRY_LABEL}</div>', unsafe_allow_html=True)

    country_overall = COUNTRY_RANKINGS.get("overall")
    rank_cols2 = st.columns(len(section_order) + 1)

    with rank_cols2[0]:
        st.markdown(
            f'<div class="rank-box-overall">'
            f'<div class="rank-value">#{country_overall}</div>'
            f'<div class="rank-label">Overall</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    for col, section_name in zip(rank_cols2[1:], section_order):
        rank = rank_lookup(COUNTRY_RANKINGS, section_name)
        rank_display = f"#{rank}" if rank is not None else "—"
        with col:
            st.markdown(
                f'<div class="rank-box">'
                f'<div class="rank-value">{rank_display}</div>'
                f'<div class="rank-label">{html.escape(section_name.upper())}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

# ---- One tab per section: questions/answers, plus matching charts ----

def render_electricity_section():
    st.markdown("---")
    st.markdown("### 📊 Electricity Consumption Data")

    filt_col1, filt_col2 = st.columns(2)
    with filt_col1:
        campus_options = ["All Campuses"] + sorted(buildings_df["campus"].dropna().unique().tolist())
        selected_campus = st.selectbox("Campus", campus_options, key="elec_campus")
    with filt_col2:
        building_options = ["All Buildings"] + sorted(buildings_df["building_name"].dropna().unique().tolist())
        selected_building = st.selectbox("Building", building_options, key="elec_building")

    elec_filtered = electricity_df.copy()
    if selected_campus != "All Campuses":
        elec_filtered = elec_filtered[elec_filtered["campus"] == selected_campus]
    if selected_building != "All Buildings":
        elec_filtered = elec_filtered[elec_filtered["building_name"] == selected_building]

    total_kwh = elec_filtered["kwh_consumed"].sum()
    total_renewable = elec_filtered["renewable_kwh"].sum()
    renewable_pct = (total_renewable / total_kwh * 100) if total_kwh else 0

    m1, m2 = st.columns(2)
    with m1:
        st.markdown(
            f'<div class="summary-box-dark">'
            f'<div class="summary-label-dark">Total Electricity</div>'
            f'<div class="summary-value-dark">{total_kwh:,.0f} kWh</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
    with m2:
        st.markdown(
            f'<div class="summary-box-dark">'
            f'<div class="summary-label-dark">Renewable Share</div>'
            f'<div class="summary-value-dark">{renewable_pct:.1f}%</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("**Monthly Electricity Consumption (kWh)**")
    elec_by_month = (
        elec_filtered.groupby("date")[["kwh_consumed", "renewable_kwh"]]
        .sum()
        .sort_index()
    )
    elec_by_month.columns = ["Total kWh", "Renewable kWh"]
    st.line_chart(elec_by_month)

    st.markdown("**Electricity Use by Campus**")
    elec_by_campus = elec_filtered.groupby("campus")["kwh_consumed"].sum().sort_values(ascending=False)
    st.bar_chart(elec_by_campus)


def render_water_section():
    st.markdown("---")
    st.markdown("### 📊 Water Consumption Data")

    filt_col1, filt_col2 = st.columns(2)
    with filt_col1:
        campus_options = ["All Campuses"] + sorted(buildings_df["campus"].dropna().unique().tolist())
        selected_campus = st.selectbox("Campus", campus_options, key="water_campus")
    with filt_col2:
        building_options = ["All Buildings"] + sorted(buildings_df["building_name"].dropna().unique().tolist())
        selected_building = st.selectbox("Building", building_options, key="water_building")

    water_filtered = water_df.copy()
    if selected_campus != "All Campuses":
        water_filtered = water_filtered[water_filtered["campus"] == selected_campus]
    if selected_building != "All Buildings":
        water_filtered = water_filtered[water_filtered["building_name"] == selected_building]

    total_water = water_filtered["water_m3"].sum()

    st.markdown(
        f'<div class="summary-box-dark">'
        f'<div class="summary-label-dark">Total Water Use</div>'
        f'<div class="summary-value-dark">{total_water:,.0f} m³</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("**Monthly Water Usage (m³)**")
    water_by_month = water_filtered.groupby("date")["water_m3"].sum().sort_index()
    water_by_month = water_by_month.rename("Water (m³)")
    st.line_chart(water_by_month)

    st.markdown("**Water Use by Campus**")
    water_by_campus = water_filtered.groupby("campus")["water_m3"].sum().sort_values(ascending=False)
    st.bar_chart(water_by_campus)


for i, section_name in enumerate(section_order):
    with tabs[i + 1]:
        section_data = indicator_df[indicator_df["Section"] == section_name]

        total_section = len(section_data)
        answered_section = section_data["Answer"].notna().sum()

        st.caption(f"{total_section} indicators • {answered_section} answered")

        sdg_list = sdgs_for_section(section_name)
        if sdg_list:
            st.markdown('<div class="sdg-heading">Related SDGs</div>', unsafe_allow_html=True)
            icons_html = "".join(
                f'<img src="{sdg_icon_url(n)}" alt="SDG {n}" title="SDG {n}">'
                for n in sdg_list
            )
            st.markdown(f'<div class="sdg-row">{icons_html}</div>', unsafe_allow_html=True)

        rows = list(section_data.iterrows())
        for start in range(0, len(rows), 4):
            chunk = rows[start:start + 4]
            cols = st.columns(4)
            for col, (_, row) in zip(cols, chunk):
                with col:
                    st.markdown(qa_card_html(row), unsafe_allow_html=True)

        name_lower = section_name.strip().lower()
        if "energy" in name_lower or "climate" in name_lower:
            render_electricity_section()
        elif "water" in name_lower:
            render_water_section()