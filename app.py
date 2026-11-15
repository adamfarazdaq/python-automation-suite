"""Streamlit front end for the spreadsheet cleaning and reporting pipeline."""

from __future__ import annotations

import re
from html import escape
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from cleaner import clean_dataframe, dataframe_to_excel_bytes, standardize_headers
from report_generator import generate_pdf_report
from sample_data import generate_sample_data

DEMO_LABEL = "Use Built-in Messy Sample Data (1-Click Demo)"
UPLOAD_LABEL = "Upload Your Excel/CSV File (.xlsx, .csv)"
MAX_UPLOAD_MB = 20

st.set_page_config(
    page_title="DataClean Pro | Spreadsheet Automation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      :root {
        --ink: #112b36;
        --muted: #5e737c;
        --line: #dfe8ea;
        --teal: #0e7c86;
        --teal-dark: #0b626b;
        --mint: #e8f4f2;
        --paper: #ffffff;
      }
      html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      }
      .stApp {
        background: linear-gradient(180deg, #f6faf9 0%, #fbfcfc 32%, #ffffff 100%);
      }
      [data-testid="stHeader"] { background: transparent; }
      .block-container {
        max-width: 1240px;
        padding-top: 1.8rem;
        padding-bottom: 2.5rem;
      }
      [data-testid="stSidebar"] {
        background: #f3f8f7;
        border-right: 1px solid var(--line);
      }
      [data-testid="stSidebar"] .block-container { padding-top: 1.4rem; }
      .brand {
        display: flex;
        align-items: center;
        gap: 11px;
        padding: 2px 0 20px 0;
      }
      .brand-mark {
        width: 38px;
        height: 38px;
        display: grid;
        place-items: center;
        border-radius: 11px;
        background: var(--ink);
        color: #fff;
        font-size: 0.78rem;
        font-weight: 800;
        letter-spacing: 0.04em;
      }
      .brand strong { display: block; color: var(--ink); font-size: 0.98rem; line-height: 1.2; }
      .brand span { display: block; color: var(--muted); font-size: 0.72rem; margin-top: 2px; }
      .side-label {
        color: var(--muted);
        font-size: 0.67rem;
        font-weight: 700;
        letter-spacing: 0.11em;
        text-transform: uppercase;
        margin: 0.2rem 0 0.65rem 0;
      }
      .pipeline {
        margin: 0.2rem 0 0 0;
        padding: 0;
        list-style: none;
      }
      .pipeline li {
        color: var(--ink);
        font-size: 0.82rem;
        line-height: 1.45;
        padding: 0 0 0 18px;
        position: relative;
        margin-bottom: 0.68rem;
      }
      .pipeline li::before {
        content: "";
        position: absolute;
        left: 0;
        top: 6px;
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: var(--teal);
        box-shadow: 0 0 0 3px rgba(14, 124, 134, 0.10);
      }
      .hero {
        position: relative;
        overflow: hidden;
        border: 1px solid var(--line);
        border-radius: 24px;
        padding: 2.4rem 2.6rem 2.2rem 2.6rem;
        background:
          radial-gradient(circle at 92% 8%, rgba(14, 124, 134, 0.13) 0%, rgba(14, 124, 134, 0) 38%),
          radial-gradient(circle at 78% 105%, rgba(196, 143, 34, 0.10) 0%, rgba(196, 143, 34, 0) 34%),
          #ffffff;
        box-shadow: 0 16px 44px rgba(17, 43, 54, 0.055);
        margin-bottom: 1.35rem;
      }
      .hero::after {
        content: "";
        position: absolute;
        right: -80px;
        top: -110px;
        width: 300px;
        height: 300px;
        border: 1px solid rgba(14, 124, 134, 0.16);
        border-radius: 50%;
      }
      .eyebrow {
        display: flex;
        align-items: center;
        gap: 8px;
        color: var(--teal);
        font-size: 0.69rem;
        font-weight: 800;
        letter-spacing: 0.13em;
        text-transform: uppercase;
      }
      .eyebrow .dot {
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: var(--teal);
        box-shadow: 0 0 0 4px rgba(14, 124, 134, 0.11);
      }
      .hero h1 {
        color: var(--ink);
        font-size: clamp(2rem, 4vw, 3.05rem);
        line-height: 1.08;
        letter-spacing: -0.035em;
        font-weight: 800;
        margin: 0.7rem 0 0.85rem 0;
        max-width: 810px;
      }
      .hero h1 .accent { color: var(--teal); }
      .hero p {
        color: var(--muted);
        font-size: 1rem;
        line-height: 1.6;
        max-width: 730px;
        margin: 0;
      }
      .chips { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 1.35rem; }
      .chips span {
        border: 1px solid var(--line);
        background: rgba(255, 255, 255, 0.72);
        color: #35505a;
        border-radius: 999px;
        padding: 5px 11px;
        font-size: 0.71rem;
        font-weight: 600;
      }
      .source-line {
        display: flex;
        align-items: center;
        gap: 8px;
        color: var(--muted);
        font-size: 0.8rem;
        margin: -0.2rem 0 1.1rem 0.35rem;
      }
      .source-line .live-dot {
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: #3f9c72;
        box-shadow: 0 0 0 4px rgba(63, 156, 114, 0.11);
      }
      .section-head {
        display: flex;
        align-items: center;
        gap: 10px;
        color: var(--ink);
        font-size: 1.08rem;
        font-weight: 750;
        letter-spacing: -0.015em;
        margin: 1.9rem 0 0.85rem 0;
      }
      .section-head .num {
        display: inline-grid;
        place-items: center;
        width: 25px;
        height: 25px;
        border-radius: 8px;
        background: var(--mint);
        color: var(--teal);
        font-size: 0.68rem;
        font-weight: 800;
        letter-spacing: 0.03em;
      }
      [data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid var(--line);
        border-radius: 16px;
        padding: 0.9rem 1rem;
        box-shadow: 0 8px 24px rgba(17, 43, 54, 0.04);
        min-height: 116px;
      }
      [data-testid="stMetricLabel"] p {
        color: var(--muted);
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.055em;
        text-transform: uppercase;
      }
      [data-testid="stMetricValue"] { color: var(--ink); font-weight: 750; }
      [data-testid="stMetricDelta"] { color: var(--muted); font-size: 0.72rem; }
      .fact-bar {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 1px;
        background: var(--line);
        border: 1px solid var(--line);
        border-radius: 14px;
        overflow: hidden;
        margin: 0.9rem 0 0.2rem 0;
      }
      .fact { background: #ffffff; padding: 0.8rem 1rem; }
      .fact span {
        display: block;
        color: var(--muted);
        font-size: 0.66rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.065em;
        margin-bottom: 3px;
      }
      .fact strong { color: var(--ink); font-size: 0.98rem; font-weight: 700; }
      .preview-head {
        display: flex;
        align-items: center;
        gap: 8px;
        color: var(--ink);
        font-size: 0.86rem;
        font-weight: 700;
        margin-bottom: 1px;
      }
      .preview-head .tag {
        font-size: 0.6rem;
        font-weight: 700;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        border-radius: 999px;
        padding: 3px 7px;
      }
      .tag.raw { background: #fbf2e5; color: #95631d; }
      .tag.clean { background: var(--mint); color: var(--teal-dark); }
      .empty-state {
        border: 1px dashed #b9cfce;
        border-radius: 20px;
        background: #f9fcfb;
        padding: 3.2rem 2rem;
        text-align: center;
        margin-top: 1rem;
      }
      .empty-state .icon {
        width: 52px;
        height: 52px;
        display: grid;
        place-items: center;
        margin: 0 auto 1rem auto;
        border-radius: 15px;
        background: var(--mint);
        color: var(--teal);
        font-size: 1.4rem;
        font-weight: 800;
      }
      .empty-state h3 { color: var(--ink); margin: 0 0 0.4rem 0; font-size: 1.08rem; }
      .empty-state p { color: var(--muted); margin: 0 auto; max-width: 470px; font-size: 0.88rem; line-height: 1.55; }
      div[data-testid="stDownloadButton"] button {
        border-radius: 10px;
        font-weight: 650;
        min-height: 2.65rem;
      }
      div[data-testid="stDownloadButton"] button[kind="primary"] {
        background: var(--teal);
        border-color: var(--teal);
      }
      div[data-testid="stDownloadButton"] button[kind="primary"]:hover {
        background: var(--teal-dark);
        border-color: var(--teal-dark);
      }
      .app-footer {
        border-top: 1px solid var(--line);
        margin-top: 2.6rem;
        padding-top: 1rem;
        color: var(--muted);
        font-size: 0.72rem;
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        flex-wrap: wrap;
      }
      @media (max-width: 760px) {
        .hero { padding: 1.7rem 1.4rem; border-radius: 18px; }
        .hero h1 { font-size: 2rem; }
        .fact-bar { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        .block-container { padding-top: 1rem; }
      }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_sample() -> pd.DataFrame:
    return generate_sample_data()


@st.cache_data(show_spinner="Building the formatted workbook...")
def build_excel(df: pd.DataFrame) -> bytes:
    return dataframe_to_excel_bytes(df)


@st.cache_data(show_spinner="Composing the executive summary...")
def build_pdf(df: pd.DataFrame, stats: dict, dataset_name: str) -> bytes:
    return generate_pdf_report(df, stats, dataset_name)


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Read a CSV or XLSX upload defensively, including accented encodings."""
    payload = uploaded_file.getvalue()
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        last_error: Exception | None = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(BytesIO(payload), encoding=encoding, sep=None, engine="python")
            except UnicodeDecodeError as exc:
                last_error = exc
            except pd.errors.ParserError as exc:
                raise ValueError("The CSV structure could not be parsed.") from exc
        raise ValueError("The CSV encoding could not be detected.") from last_error

    return pd.read_excel(BytesIO(payload), engine="openpyxl")


def preview_column_config(df: pd.DataFrame) -> dict:
    config: dict = {}
    hints = ("price", "amount", "total", "revenue", "cost", "fee", "salary", "budget")
    for column in df.columns:
        lowered = str(column).lower()
        if "supplier" in lowered:
            width = 150
        elif "product" in lowered:
            width = 115
        elif "date" in lowered:
            width = 90
        elif any(hint in lowered for hint in hints):
            width = 115
        elif "id" in lowered:
            width = 95
        else:
            width = 90

        if not pd.api.types.is_numeric_dtype(df[column]):
            config[column] = st.column_config.TextColumn(width=width)
        elif any(hint in lowered for hint in hints):
            config[column] = st.column_config.NumberColumn(format="dollar", width=width)
        elif str(column) == "Quantity":
            config[column] = st.column_config.NumberColumn(format="localized", width=width)
        else:
            config[column] = st.column_config.NumberColumn(format="localized", width=width)
    return config


def file_stem(name: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9_-]+", "_", name).strip("_")
    return stem[:60] or "cleaned_data"


def preview_frames(raw_df: pd.DataFrame, cleaned_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Pick a readable set of matching columns for the side-by-side preview."""
    raw_names = dict(zip(standardize_headers(raw_df.columns), raw_df.columns))
    preferred = ["Order ID", "Order Date", "Supplier", "Total Amount"]
    selected = [column for column in preferred if column in cleaned_df.columns and column in raw_names]

    for column in cleaned_df.columns:
        if len(selected) == 4:
            break
        if column not in selected and column in raw_names:
            selected.append(column)

    raw_columns = [raw_names[column] for column in selected]
    return raw_df[raw_columns].head(100), cleaned_df[selected].head(100)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(
        """
        <div class="brand">
          <div class="brand-mark">DC</div>
          <div><strong>DataClean Pro</strong><span>Spreadsheet automation</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown('<p class="side-label">Data source</p>', unsafe_allow_html=True)
    mode = st.radio("Data source", [DEMO_LABEL, UPLOAD_LABEL], label_visibility="collapsed")

    uploaded_file = None
    if mode == UPLOAD_LABEL:
        uploaded_file = st.file_uploader(
            "Choose an Excel or CSV file",
            type=["xlsx", "csv"],
            help="The first sheet is used for Excel files. Maximum size: 20 MB.",
        )

    st.divider()
    st.markdown('<p class="side-label">What happens</p>', unsafe_allow_html=True)
    st.markdown(
        """
        <ul class="pipeline">
          <li>Headers and whitespace are standardized</li>
          <li>Exact duplicates are removed</li>
          <li>Dates become YYYY-MM-DD</li>
          <li>Currency text becomes numbers</li>
          <li>Missing values are repaired</li>
          <li>Excel and PDF reports are built</li>
        </ul>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Files are processed in memory and never stored by this app.")

# ---------------------------------------------------------------------------
# Hero and data source
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
      <div class="eyebrow"><span class="dot"></span>Python automation suite</div>
      <h1>Messy spreadsheets in.<br><span class="accent">Decision-ready reports out.</span></h1>
      <p>Clean supplier orders, standardize every field, and download a formatted Excel workbook
      plus a one-page executive PDF — without touching a formula.</p>
      <div class="chips">
        <span>Excel + CSV</span>
        <span>In-memory processing</span>
        <span>1-click demo</span>
        <span>Client-ready PDF</span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if mode == UPLOAD_LABEL and uploaded_file is None:
    st.markdown(
        """
        <div class="empty-state">
          <div class="icon">+</div>
          <h3>Add a spreadsheet to begin</h3>
          <p>Drop an <strong>.xlsx</strong> or <strong>.csv</strong> file in the sidebar. The tool will clean it,
          show a before-and-after preview, and prepare both downloads.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

if mode == UPLOAD_LABEL:
    if uploaded_file.size > MAX_UPLOAD_MB * 1024 * 1024:
        st.error(f"That file is larger than {MAX_UPLOAD_MB} MB. Please upload a smaller file.")
        st.stop()
    try:
        raw_df = read_uploaded_file(uploaded_file)
    except Exception:  # noqa: BLE001 - third-party file parsers raise varied errors
        st.error("We could not read that file. Check that it is a valid .xlsx or .csv and try again.")
        st.stop()

    dataset_name = Path(uploaded_file.name).stem.replace("_", " ").strip() or "Uploaded dataset"
    source_text = f"Uploaded file · {uploaded_file.name}"
else:
    raw_df = load_sample()
    dataset_name = "Built-in pharmacy sample"
    source_text = "Built-in demo · 120 unique supplier orders with 10 exact duplicates"

st.markdown(
    f'<div class="source-line"><span class="live-dot"></span>{escape(source_text)}</div>',
    unsafe_allow_html=True,
)

try:
    with st.spinner("Cleaning data..."):
        cleaned_df, stats = clean_dataframe(raw_df)
except ValueError as exc:
    st.error(str(exc))
    st.stop()
except Exception:  # noqa: BLE001 - keep demo failures user-friendly
    st.error("The cleaning step failed. Please check the file and try again.")
    st.stop()

# ---------------------------------------------------------------------------
# Metrics
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="section-head"><span class="num">01</span>Results overview</div>',
    unsafe_allow_html=True,
)

metric_columns = st.columns(4, gap="medium")
with metric_columns[0]:
    st.metric(
        "Rows processed",
        f"{stats['rows_before']:,}",
        f"{stats['rows_after']:,} clean rows",
        delta_color="off",
        delta_arrow="off",
        help="Rows read from the source, including duplicates that were removed.",
    )
with metric_columns[1]:
    st.metric(
        "Duplicates removed",
        f"{stats['duplicates_removed']:,}",
        "exact matches",
        delta_color="off",
        delta_arrow="off",
        help="Exact duplicate rows removed after whitespace normalization.",
    )
with metric_columns[2]:
    st.metric(
        "Missing values fixed",
        f"{stats['missing_values_fixed']:,}",
        f"{stats['missing_values_remaining']:,} date gaps left",
        delta_color="off",
        delta_arrow="off",
        help="Blank cells filled with a sensible value. Unparseable dates stay blank.",
    )
with metric_columns[3]:
    st.metric(
        "Formatting issues fixed",
        f"{stats['formatting_issues_fixed']:,}",
        "dates, currency, text",
        delta_color="off",
        delta_arrow="off",
        help="Cells whose whitespace, casing, date format, or currency format changed.",
    )

if stats.get("total_amount") is not None:
    top_category = escape(str(stats.get("top_category") or "—"))
    st.markdown(
        f"""
        <div class="fact-bar">
          <div class="fact"><span>Total order value</span><strong>${stats['total_amount']:,.2f}</strong></div>
          <div class="fact"><span>Average order value</span><strong>${stats['average_order_value']:,.2f}</strong></div>
          <div class="fact"><span>Total units</span><strong>{stats.get('total_units') or 0:,}</strong></div>
          <div class="fact"><span>Top category</span><strong>{top_category}</strong></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# Before / after preview
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="section-head"><span class="num">02</span>Before &amp; after</div>',
    unsafe_allow_html=True,
)

raw_column, clean_column = st.columns(2, gap="medium")
raw_preview_df, cleaned_preview_df = preview_frames(raw_df, cleaned_df)
with raw_column:
    st.markdown(
        '<div class="preview-head">Raw input <span class="tag raw">Unchanged</span></div>',
        unsafe_allow_html=True,
    )
    st.caption(
        f"{len(raw_df):,} rows · {len(raw_preview_df.columns)} of {len(raw_df.columns)} columns · first 100 shown"
    )
    raw_preview = raw_preview_df.map(lambda value: "" if pd.isna(value) else str(value))
    st.dataframe(
        raw_preview,
        width="stretch",
        height=365,
        hide_index=True,
        column_config=preview_column_config(raw_preview),
    )

with clean_column:
    st.markdown(
        '<div class="preview-head">Cleaned output <span class="tag clean">Standardized</span></div>',
        unsafe_allow_html=True,
    )
    st.caption(
        f"{len(cleaned_df):,} rows · {len(cleaned_preview_df.columns)} of {len(cleaned_df.columns)} columns · first 100 shown"
    )
    cleaned_preview = cleaned_preview_df.copy()
    for column in cleaned_preview.columns:
        if not pd.api.types.is_numeric_dtype(cleaned_preview[column]):
            cleaned_preview[column] = cleaned_preview[column].fillna("—")
    st.dataframe(
        cleaned_preview,
        width="stretch",
        height=365,
        hide_index=True,
        column_config=preview_column_config(cleaned_preview),
    )

# ---------------------------------------------------------------------------
# Downloads
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="section-head"><span class="num">03</span>Download deliverables</div>',
    unsafe_allow_html=True,
)

excel_bytes: bytes | None = None
pdf_bytes: bytes | None = None
try:
    excel_bytes = build_excel(cleaned_df)
except Exception:  # noqa: BLE001 - show a recoverable UI error
    excel_bytes = None
try:
    pdf_bytes = build_pdf(cleaned_df, stats, dataset_name)
except Exception:  # noqa: BLE001 - show a recoverable UI error
    pdf_bytes = None

safe_stem = file_stem(dataset_name)
download_columns = st.columns(2, gap="medium")

with download_columns[0], st.container(border=True):
    st.markdown("**Cleaned Excel workbook**")
    st.caption("Styled headers, auto-sized columns, filters, and numeric formats.")
    if excel_bytes is not None:
        st.download_button(
            "Download Cleaned Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"{safe_stem}_cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width="stretch",
        )
        st.caption(f"{len(excel_bytes) / 1024:,.0f} KB · generated in memory")
    else:
        st.error("The Excel workbook could not be generated. Try a smaller file.")

with download_columns[1], st.container(border=True):
    st.markdown("**Executive summary PDF**")
    st.caption("One page: quality metrics, business totals, and a data preview.")
    if pdf_bytes is not None:
        st.download_button(
            "Download Executive Summary (.pdf)",
            data=pdf_bytes,
            file_name=f"{safe_stem}_summary.pdf",
            mime="application/pdf",
            width="stretch",
        )
        st.caption(f"{len(pdf_bytes) / 1024:,.0f} KB · one-page A4 report")
    else:
        st.error("The PDF report could not be generated. Try a smaller file.")

st.markdown(
    """
    <div class="app-footer">
      <span>DataClean Pro · Python, pandas, Streamlit, openpyxl, fpdf2</span>
      <span>No uploads persisted · No API keys required</span>
    </div>
    """,
    unsafe_allow_html=True,
)
