"""Core cleaning pipeline and formatted Excel export.

The public API is intentionally small:

* :func:`clean_dataframe` returns a cleaned DataFrame plus a stats dictionary.
* :func:`dataframe_to_excel_bytes` returns a styled ``.xlsx`` workbook as bytes.
"""

from __future__ import annotations

import math
import numbers
import re
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

CANONICAL_HEADERS = {
    "orderid": "Order ID",
    "ordernumber": "Order ID",
    "ponumber": "Order ID",
    "purchaseorder": "Order ID",
    "orderdate": "Order Date",
    "purchasedate": "Order Date",
    "invoicedate": "Order Date",
    "supplier": "Supplier",
    "suppliername": "Supplier",
    "vendor": "Supplier",
    "vendorname": "Supplier",
    "product": "Product",
    "productname": "Product",
    "item": "Product",
    "itemname": "Product",
    "description": "Product",
    "category": "Category",
    "productcategory": "Category",
    "qty": "Quantity",
    "quantity": "Quantity",
    "units": "Quantity",
    "unitssold": "Quantity",
    "unitprice": "Unit Price",
    "price": "Unit Price",
    "unitcost": "Unit Price",
    "totalamount": "Total Amount",
    "ordertotal": "Total Amount",
    "grandtotal": "Total Amount",
    "total": "Total Amount",
    "amount": "Total Amount",
    "revenue": "Revenue",
    "status": "Status",
    "deliverystatus": "Status",
    "orderstatus": "Status",
    "region": "Region",
    "warehouseregion": "Region",
    "territory": "Region",
}

DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%d.%m.%Y",
    "%b %d, %Y",
    "%B %d, %Y",
    "%d %b %Y",
    "%d %B %Y",
    "%m/%d/%Y",
    "%m-%d-%Y",
    "%Y%m%d",
)

PROPER_TOKENS = {
    "gmbh": "GmbH",
    "sarl": "SARL",
    "sas": "SAS",
    "inc": "Inc.",
    "inc.": "Inc.",
    "ltd": "Ltd.",
    "ltd.": "Ltd.",
    "llc": "LLC",
    "co": "Co.",
    "co.": "Co.",
    "id": "ID",
    "sku": "SKU",
    "url": "URL",
    "api": "API",
    "po": "PO",
    "usa": "USA",
    "uk": "UK",
    "eu": "EU",
    "otc": "OTC",
    "rx": "Rx",
    "iv": "IV",
    "iu": "IU",
    "d3": "D3",
    "purecare": "PureCare",
}

MONEY_HINTS = ("price", "amount", "total", "cost", "revenue", "sales", "fee", "balance", "spend", "budget", "salary")
CATEGORY_HINTS = (
    "supplier",
    "vendor",
    "category",
    "status",
    "region",
    "country",
    "city",
    "segment",
    "channel",
    "method",
    "payment",
    "priority",
    "type",
)


def _is_missing(value: Any) -> bool:
    if value is None or value is pd.NA or value is pd.NaT:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return bool(isinstance(value, str) and not value.strip())


def _canonical_key(name: Any) -> str:
    text = str(name).replace("\ufeff", "").strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def _smart_title(text: str) -> str:
    titled = text.title()
    parts = re.split(r"(\s+)", titled)
    output = []
    for part in parts:
        if not part or part.isspace():
            output.append(part)
            continue
        mapped = PROPER_TOKENS.get(part.lower())
        if mapped:
            output.append(mapped)
            continue
        stripped = part.lower().rstrip(".")
        mapped = PROPER_TOKENS.get(stripped)
        if mapped:
            output.append(mapped + ("." if part.endswith(".") and not mapped.endswith(".") else ""))
        else:
            output.append(part)
    return "".join(output)


def standardize_headers(columns) -> list[str]:
    """Return readable, unique column names for arbitrary input headers."""
    cleaned: list[str] = []
    used: dict[str, int] = {}

    for index, column in enumerate(columns, start=1):
        key = _canonical_key(column)
        if key in CANONICAL_HEADERS:
            name = CANONICAL_HEADERS[key]
        else:
            base = re.sub(r"[_\-.]+", " ", str(column)).replace("\ufeff", "")
            base = re.sub(r"\s+", " ", base).strip()
            name = _smart_title(base) if base else f"Column {index}"

        count = used.get(name, 0) + 1
        used[name] = count
        cleaned.append(name if count == 1 else f"{name} ({count})")

    return cleaned


def _is_date_column(name: str) -> bool:
    return "date" in name.lower()


def _is_identifier(name: str) -> bool:
    tokens = re.split(r"[^a-z0-9]+", name.lower())
    return any(token in {"id", "code", "sku", "ref", "reference", "number"} for token in tokens)


def _has_money_hint(name: str) -> bool:
    lowered = name.lower()
    return any(hint in lowered for hint in MONEY_HINTS)


def _is_categorical(name: str) -> bool:
    lowered = name.lower()
    return any(hint in lowered for hint in CATEGORY_HINTS)


def _parse_number(value: Any) -> float:
    """Parse numbers such as ``$1,250.50``, ``USD 980``, or ``(120.00)``."""
    if _is_missing(value):
        return float("nan")
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, numbers.Real):
        return float(value)

    text = str(value).strip().replace("\u00a0", " ").replace("\u2212", "-")
    if not text or text.lower() in {"n/a", "na", "none", "null", "-"}:
        return float("nan")

    negative = text.startswith("(") and text.endswith(")")
    text = re.sub(r"[^\d,.\-+]", "", text)
    if not text:
        return float("nan")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    try:
        number = float(text)
    except ValueError:
        return float("nan")
    return -number if negative else number


def _parse_date(value: Any) -> date | None:
    """Parse common date formats; ambiguous slash dates are treated as day-first."""
    if _is_missing(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, numbers.Real) and not isinstance(value, bool):
        serial = float(value)
        if 1 <= serial <= 80000:
            return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
        return None

    text = re.sub(r"\s+", " ", str(value)).strip()
    if not text:
        return None

    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        pass

    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    try:
        parsed = pd.to_datetime(text, errors="coerce")
    except (TypeError, ValueError):
        return None
    return None if pd.isna(parsed) else parsed.date()


def _normalise_cell(value: Any) -> Any:
    if isinstance(value, str):
        cleaned = re.sub(r"\s+", " ", value.replace("\ufeff", "")).strip()
        return cleaned if cleaned else pd.NA
    return value


def _looks_numeric(series: pd.Series, name: str) -> bool:
    if _is_date_column(name) or _is_identifier(name):
        return False
    if pd.api.types.is_numeric_dtype(series):
        return True

    non_missing = series.dropna()
    if non_missing.empty:
        return False
    if _has_money_hint(name):
        return True

    sample = non_missing.head(100)
    return float(sample.map(_parse_number).notna().mean()) >= 0.9


def _all_whole(series: pd.Series) -> bool:
    if series.empty:
        return False
    return bool(((series.dropna() % 1).abs() < 1e-9).all())


def _display(value: Any) -> str:
    if _is_missing(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, numbers.Real) and not isinstance(value, bool):
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number)
    return str(value)


def _count_formatting_fixes(before: pd.DataFrame, after: pd.DataFrame) -> int:
    fixes = 0
    for column in after.columns:
        if column not in before.columns:
            continue
        for old_value, new_value in zip(before[column], after[column]):
            if _is_missing(old_value):
                continue
            if _display(old_value) != _display(new_value):
                fixes += 1
    return fixes


def clean_dataframe(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Clean a raw spreadsheet and return ``(cleaned_dataframe, stats)``.

    The function is defensive by design: unknown columns are preserved, empty
    rows/columns are dropped, and missing values are imputed rather than raising.
    """
    if not isinstance(raw, pd.DataFrame):
        raise TypeError("clean_dataframe expects a pandas DataFrame")
    if raw.empty:
        raise ValueError("The file contains no data rows.")

    work = raw.dropna(how="all").dropna(axis=1, how="all").copy()
    if work.empty:
        raise ValueError("The file contains no data rows.")

    work.columns = standardize_headers(work.columns)
    rows_before = len(work)
    original_rows = work.copy()

    # Normalise whitespace first so duplicates separated only by spacing disappear.
    work = work.map(_normalise_cell)
    keep = ~work.duplicated(keep="first")
    work = work.loc[keep].reset_index(drop=True)
    original_rows = original_rows.loc[keep].reset_index(drop=True)

    duplicates_removed = rows_before - len(work)
    missing_before = int(work.isna().sum().sum())

    numeric_series: dict[str, pd.Series] = {}

    for column in work.columns:
        name = str(column)
        if _is_date_column(name):
            parsed_dates = work[column].map(_parse_date)
            work[column] = parsed_dates.map(lambda value: value.isoformat() if isinstance(value, date) else pd.NA)
        elif _looks_numeric(work[column], name):
            numeric_series[column] = work[column].map(_parse_number)
        else:
            text = work[column].map(lambda value: str(value) if not _is_missing(value) else pd.NA)
            if _is_identifier(name):
                text = text.str.upper()
            elif _is_categorical(name):
                text = text.map(lambda value: _smart_title(value) if isinstance(value, str) else value)
            work[column] = text.fillna("Unknown")

    # Recover missing totals from quantity x unit price before imputing either one.
    required_totals = {"Quantity", "Unit Price", "Total Amount"}
    if required_totals.issubset(numeric_series):
        total = numeric_series["Total Amount"]
        computed = numeric_series["Quantity"] * numeric_series["Unit Price"]
        numeric_series["Total Amount"] = total.where(total.notna(), computed)

    for column, series in numeric_series.items():
        if column == "Total Amount":
            fill_value = 0.0
        else:
            median = series.median(skipna=True)
            fill_value = float(median) if pd.notna(median) else 0.0

        filled = series.fillna(fill_value)
        if column == "Quantity" or _all_whole(filled):
            filled = filled.round().astype("int64")
        work[column] = filled

    missing_after = int(work.isna().sum().sum())
    missing_fixed = max(0, missing_before - missing_after)
    formatting_fixed = _count_formatting_fixes(original_rows, work)

    stats: dict[str, Any] = {
        "rows_before": rows_before,
        "rows_after": len(work),
        "columns_cleaned": len(work.columns),
        "duplicates_removed": duplicates_removed,
        "missing_values_fixed": missing_fixed,
        "missing_values_remaining": missing_after,
        "formatting_issues_fixed": formatting_fixed,
        "issues_fixed": missing_fixed + formatting_fixed,
        "total_amount": None,
        "average_order_value": None,
        "total_units": None,
        "top_category": None,
        "date_start": None,
        "date_end": None,
    }

    if "Total Amount" in work.columns and pd.api.types.is_numeric_dtype(work["Total Amount"]):
        stats["total_amount"] = round(float(work["Total Amount"].sum()), 2)
        stats["average_order_value"] = round(float(work["Total Amount"].mean()), 2)
    if "Quantity" in work.columns and pd.api.types.is_numeric_dtype(work["Quantity"]):
        stats["total_units"] = int(work["Quantity"].sum())
    if "Category" in work.columns:
        categories = work["Category"].dropna().astype(str)
        categories = categories[~categories.str.lower().eq("unknown")]
        if not categories.empty:
            stats["top_category"] = categories.value_counts().index[0]
    if "Order Date" in work.columns:
        dates = work["Order Date"].dropna().astype(str)
        if not dates.empty:
            stats["date_start"] = min(dates)
            stats["date_end"] = max(dates)

    return work, stats


def _excel_safe(value: Any) -> Any:
    """Escape spreadsheet formula prefixes in user-supplied text."""
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Cleaned Data") -> bytes:
    """Write ``df`` to a styled ``.xlsx`` workbook and return the bytes."""
    safe_df = df.copy()
    for column in safe_df.columns:
        if safe_df[column].dtype == object or pd.api.types.is_string_dtype(safe_df[column]):
            safe_df[column] = safe_df[column].map(_excel_safe)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        safe_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.book[sheet_name]

        header_fill = PatternFill("solid", fgColor="0E4C5C")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        band_fill = PatternFill("solid", fgColor="F3F8F8")
        thin_border = Border(bottom=Side(style="thin", color="E0EAEC"))

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        worksheet.row_dimensions[1].height = 26
        worksheet.freeze_panes = "A2"
        worksheet.sheet_view.showGridLines = False
        worksheet.sheet_properties.tabColor = "0E7C86"
        if worksheet.max_row > 1:
            worksheet.auto_filter.ref = worksheet.dimensions

        apply_banding = worksheet.max_row <= 5000
        for row_index, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                if apply_banding and row_index % 2 == 0:
                    cell.fill = band_fill

        for column_index, column in enumerate(safe_df.columns, start=1):
            letter = get_column_letter(column_index)
            longest = len(str(column))
            for row_index in range(1, min(worksheet.max_row, 200) + 1):
                value = worksheet.cell(row=row_index, column=column_index).value
                if value is not None:
                    longest = max(longest, len(str(value)))
            worksheet.column_dimensions[letter].width = min(max(longest + 2, 12), 45)

            lowered = str(column).lower()
            if _has_money_hint(lowered) and pd.api.types.is_numeric_dtype(safe_df[column]):
                number_format = "$#,##0.00"
            elif column == "Quantity" and pd.api.types.is_numeric_dtype(safe_df[column]):
                number_format = "#,##0"
            elif pd.api.types.is_numeric_dtype(safe_df[column]):
                number_format = "#,##0.00"
            else:
                continue

            for row in worksheet.iter_rows(min_row=2, min_col=column_index, max_col=column_index):
                if isinstance(row[0].value, (int, float)):
                    row[0].number_format = number_format
                    row[0].alignment = Alignment(horizontal="right", vertical="center")

    return buffer.getvalue()
