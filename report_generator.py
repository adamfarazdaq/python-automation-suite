"""One-page executive PDF summary built with fpdf2.

The report uses a system Unicode font when available (Arial, Segoe UI, DejaVu,
or Liberation Sans) and falls back to a transliterated core font otherwise.
Text is filtered against the font's glyph map, so accented or unsupported
characters can never trigger an fpdf2 encoding error.
"""

from __future__ import annotations

import os
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from fontTools.ttLib import TTFont
from fpdf import FPDF

NAVY = (12, 56, 69)
TEAL = (14, 124, 134)
INK = (31, 42, 51)
MUTED = (94, 115, 124)
LINE = (223, 232, 234)
AMBER = (196, 143, 34)
PALE = (247, 250, 250)
WHITE = (255, 255, 255)
LIGHT_TEXT = (199, 224, 221)

PREFERRED_COLUMNS = [
    "Order ID",
    "Order Date",
    "Supplier",
    "Product",
    "Category",
    "Quantity",
    "Unit Price",
    "Total Amount",
    "Status",
    "Region",
]
COLUMN_WEIGHTS = {
    "Order ID": 24,
    "Order Date": 20,
    "Supplier": 30,
    "Product": 38,
    "Category": 24,
    "Quantity": 16,
    "Unit Price": 21,
    "Total Amount": 25,
    "Status": 20,
    "Region": 20,
}


def _font_pairs() -> list[tuple[Path, Path]]:
    windows_fonts = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    return [
        (windows_fonts / "arial.ttf", windows_fonts / "arialbd.ttf"),
        (windows_fonts / "segoeui.ttf", windows_fonts / "seguisb.ttf"),
        (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"), Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
        (Path("/usr/share/fonts/dejavu/DejaVuSans.ttf"), Path("/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf")),
        (Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"), Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf")),
        (
            Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
            Path("/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
        ),
        (Path("/Library/Fonts/Arial.ttf"), Path("/Library/Fonts/Arial Bold.ttf")),
    ]


def _font_codepoints(path: Path) -> set[int] | None:
    try:
        font = TTFont(str(path), lazy=True)
        try:
            return set(font.getBestCmap() or {})
        finally:
            font.close()
    except Exception:  # noqa: BLE001 - font parsing failures should fall back safely
        return None


def _configure_font(pdf: FPDF) -> tuple[str, set[int] | None, bool]:
    for regular, bold in _font_pairs():
        if not regular.exists():
            continue
        bold_path = bold if bold.exists() else regular
        try:
            pdf.add_font("ReportSans", "", str(regular))
            pdf.add_font("ReportSans", "B", str(bold_path))
        except Exception:  # noqa: BLE001, S112 - try the next font family
            continue

        regular_points = _font_codepoints(regular)
        bold_points = _font_codepoints(bold_path)
        if regular_points is not None and bold_points is not None:
            supported = regular_points & bold_points
        else:
            supported = regular_points
        return "ReportSans", supported, True

    return "Helvetica", None, False


def _pdf_safe(value: Any, supported: set[int] | None, unicode_font: bool) -> str:
    text = "" if value is None else str(value)
    text = re.sub(r"\s+", " ", text).strip()

    if not unicode_font:
        replacements = {
            "\u2014": "-",
            "\u2013": "-",
            "\u20ac": "EUR",
            "\u2022": "-",
            "\u2026": "...",
            "\u2019": "'",
            "\u2018": "'",
            "\u201c": '"',
            "\u201d": '"',
        }
        for source, target in replacements.items():
            text = text.replace(source, target)
        text = unicodedata.normalize("NFKD", text)
        return text.encode("latin-1", "ignore").decode("latin-1")

    text = "".join(character if ord(character) >= 32 else " " for character in text)
    if supported is not None:
        text = "".join(character if ord(character) in supported else "?" for character in text)
    return text


def _compact_money(value: float | None) -> str:
    if value is None:
        return "-"
    absolute = abs(value)
    if absolute >= 1_000_000_000:
        return f"${value / 1_000_000_000:,.2f}B"
    if absolute >= 1_000_000:
        return f"${value / 1_000_000:,.2f}M"
    if absolute >= 10_000:
        return f"${value / 1_000:,.1f}K"
    return f"${value:,.2f}"


def _money(value: float | None) -> str:
    return "-" if value is None else f"${value:,.2f}"


def _short_date(value: str | None) -> str:
    if not value:
        return "-"
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%b %d, %Y")
    except ValueError:
        return value


def _date_coverage(start: str | None, end: str | None) -> str:
    if not start or not end:
        return "-"
    try:
        start_date = datetime.strptime(start, "%Y-%m-%d")
        end_date = datetime.strptime(end, "%Y-%m-%d")
    except ValueError:
        return f"{start} - {end}"
    if start_date.year == end_date.year:
        return f"{start_date:%b %d} - {end_date:%b %d, %Y}"
    return f"{_short_date(start)} - {_short_date(end)}"


def _is_missing(value: Any) -> bool:
    if value is None or value is pd.NA or value is pd.NaT:
        return True
    return isinstance(value, float) and pd.isna(value)


def _format_cell(column: str, value: Any) -> str:
    if _is_missing(value):
        return "-"
    lowered = column.lower()
    if any(hint in lowered for hint in ("price", "amount", "total", "revenue", "cost", "fee")):
        try:
            return f"${float(value):,.2f}"
        except (TypeError, ValueError):
            return str(value)
    if column == "Quantity":
        try:
            return f"{int(value):,}"
        except (TypeError, ValueError):
            return str(value)
    if isinstance(value, float):
        return f"{value:,.2f}"
    return str(value)


def _table_columns(df: pd.DataFrame) -> list[str]:
    selected = [column for column in PREFERRED_COLUMNS if column in df.columns][:6]
    if len(selected) < 6:
        for column in df.columns:
            if column not in selected:
                selected.append(column)
            if len(selected) == 6:
                break
    return selected


class _ExecutiveReport(FPDF):
    def __init__(self, dataset_name: str, stats: dict[str, Any]) -> None:
        super().__init__(orientation="P", unit="mm", format="A4")
        self.stats = stats
        self.dataset_name = dataset_name
        self.set_auto_page_break(auto=False)
        self.set_margins(14, 14, 14)
        self.alias_nb_pages()
        self.font_family, self.supported, self.unicode_font = _configure_font(self)
        self.set_title(self.t("Executive Summary - Data Cleaning"))
        self.set_author(self.t("DataClean Pro"))
        self.set_creator(self.t("DataClean Pro"))
        self.set_subject(self.t("Cleaned spreadsheet summary"))

    def t(self, value: Any) -> str:
        return _pdf_safe(value, self.supported, self.unicode_font)

    def fit(self, value: Any, width: float) -> str:
        text = self.t(value)
        if self.get_string_width(text) <= width:
            return text
        while text and self.get_string_width(text + "...") > width:
            text = text[:-1]
        return text.rstrip() + "..."

    def heading(self, text: str, y: float) -> None:
        self.set_xy(14, y)
        self.set_font(self.font_family, "B", 8)
        self.set_text_color(*TEAL)
        self.cell(182, 5, self.t(text), new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(*LINE)
        self.line(14, y + 6, 196, y + 6)

    def draw_header(self) -> None:
        self.set_fill_color(*NAVY)
        self.rect(0, 0, 210, 35, "F")
        self.set_fill_color(*TEAL)
        self.rect(0, 35, 210, 1.4, "F")

        self.set_xy(14, 8)
        self.set_font(self.font_family, "B", 6.8)
        self.set_text_color(*LIGHT_TEXT)
        self.cell(100, 4, self.t("DATA QUALITY REPORT"))

        self.set_xy(14, 13)
        self.set_font(self.font_family, "B", 21)
        self.set_text_color(*WHITE)
        self.cell(120, 10, self.t("Executive Summary"))

        self.set_xy(14, 24)
        self.set_font(self.font_family, "", 8)
        self.set_text_color(*LIGHT_TEXT)
        subtitle = f"{self.dataset_name}  |  Generated {datetime.now():%B %d, %Y}"
        self.cell(125, 5, self.fit(subtitle, 124))

        badge_x, badge_y, badge_w, badge_h = 145, 10.5, 51, 16
        self.set_fill_color(24, 78, 91)
        self.rect(badge_x, badge_y, badge_w, badge_h, "F", round_corners=True, corner_radius=2)
        self.set_xy(badge_x, badge_y + 3)
        self.set_font(self.font_family, "B", 6.3)
        self.set_text_color(127, 212, 201)
        self.cell(badge_w, 4, self.t("STATUS"), align="C")
        self.set_xy(badge_x, badge_y + 7.2)
        self.set_font(self.font_family, "B", 9.5)
        self.set_text_color(*WHITE)
        self.cell(badge_w, 5, self.t("Clean & ready"), align="C")

    def draw_cards(self) -> None:
        stats = self.stats
        missing = int(stats.get("missing_values_fixed", 0))
        formatting = int(stats.get("formatting_issues_fixed", 0))
        total = stats.get("total_amount")
        average = stats.get("average_order_value")
        cards = [
            ("ROWS CLEANED", f"{stats.get('rows_after', 0):,}", f"of {stats.get('rows_before', 0):,} processed", TEAL),
            ("DUPLICATES", f"{stats.get('duplicates_removed', 0):,}", "exact rows removed", AMBER),
            ("ISSUES FIXED", f"{missing + formatting:,}", f"{missing:,} missing  |  {formatting:,} format", TEAL),
            ("ORDER VALUE", _compact_money(total), f"avg {_compact_money(average)}", (58, 125, 86)),
        ]

        card_y, card_h, gap = 44, 26, 3
        card_w = (182 - gap * 3) / 4
        for index, (label, value, footnote, accent) in enumerate(cards):
            x = 14 + index * (card_w + gap)
            self.set_fill_color(*PALE)
            self.set_draw_color(*LINE)
            self.rect(x, card_y, card_w, card_h, "DF", round_corners=True, corner_radius=2)
            self.set_fill_color(*accent)
            self.rect(x + 4.5, card_y + 5, 1.8, card_h - 10, "F", round_corners=True, corner_radius=0.8)

            self.set_xy(x + 9, card_y + 4.6)
            self.set_font(self.font_family, "B", 5.9)
            self.set_text_color(*MUTED)
            self.cell(card_w - 11, 4, self.fit(label, card_w - 11))

            self.set_xy(x + 9, card_y + 9.2)
            self.set_font(self.font_family, "B", 14.5)
            self.set_text_color(*NAVY)
            self.cell(card_w - 11, 7, self.fit(value, card_w - 11))

            self.set_xy(x + 9, card_y + 18.2)
            self.set_font(self.font_family, "", 5.9)
            self.set_text_color(*MUTED)
            self.cell(card_w - 11, 4, self.fit(footnote, card_w - 11))

    def metric_row(self, x: float, y: float, label: str, value: str, width: float) -> float:
        self.set_xy(x, y)
        self.set_font(self.font_family, "", 7.5)
        self.set_text_color(*MUTED)
        self.cell(width - 42, 7.4, self.fit(label, width - 44))
        self.set_xy(x + width - 42, y)
        self.set_font(self.font_family, "B", 8)
        self.set_text_color(*INK)
        self.cell(42, 7.4, self.fit(value, 41), align="R")
        self.set_draw_color(*LINE)
        self.line(x, y + 7.6, x + width, y + 7.6)
        return y + 8.6

    def draw_overview(self) -> None:
        stats = self.stats
        self.heading("AT A GLANCE", 78)
        left_x, right_x, width = 14, 108, 88

        self.set_xy(left_x, 88)
        self.set_font(self.font_family, "B", 6.5)
        self.set_text_color(*TEAL)
        self.cell(width, 4, self.t("DATA QUALITY"))
        self.set_xy(right_x, 88)
        self.set_font(self.font_family, "B", 6.5)
        self.cell(width, 4, self.t("KEY METRICS"))

        y = 94
        left_rows = [
            ("Rows processed", f"{stats.get('rows_before', 0):,}"),
            ("Rows after cleaning", f"{stats.get('rows_after', 0):,}"),
            ("Duplicates removed", f"{stats.get('duplicates_removed', 0):,}"),
            ("Missing values repaired", f"{stats.get('missing_values_fixed', 0):,}"),
            ("Formatting fixes", f"{stats.get('formatting_issues_fixed', 0):,}"),
        ]
        for label, value in left_rows:
            y = self.metric_row(left_x, y, label, value, width)

        y = 94
        right_rows = [
            ("Total order value", _money(stats.get("total_amount"))),
            ("Average order value", _money(stats.get("average_order_value"))),
            ("Total units", f"{stats['total_units']:,}" if stats.get("total_units") is not None else "-"),
            ("Top category", stats.get("top_category") or "-"),
            ("Date coverage", _date_coverage(stats.get("date_start"), stats.get("date_end"))),
        ]
        for label, value in right_rows:
            y = self.metric_row(right_x, y, label, value, width)

    def draw_table(self, df: pd.DataFrame) -> None:
        self.heading("CLEANED DATA PREVIEW", 140)
        columns = _table_columns(df)
        preview = df.head(12)

        self.set_xy(14, 148)
        self.set_font(self.font_family, "", 6.3)
        self.set_text_color(*MUTED)
        caption = f"First {len(preview)} of {len(df):,} rows  |  {len(df.columns)} cleaned columns"
        self.cell(182, 4, self.t(caption), align="R")

        if not columns:
            return

        weights = [COLUMN_WEIGHTS.get(column, max(16, min(40, len(str(column)) + 8))) for column in columns]
        total_weight = sum(weights)
        widths = [weight / total_weight * 182 for weight in weights]

        header_y, header_h, row_h = 154, 7.2, 6.5
        self.set_xy(14, header_y)
        self.set_fill_color(*NAVY)
        self.set_font(self.font_family, "B", 6.2)
        self.set_text_color(*WHITE)
        for column, width in zip(columns, widths):
            self.cell(width, header_h, self.fit(str(column).upper(), width - 2.4), border="B", fill=True, align="C")

        for row_index, (_, row) in enumerate(preview.iterrows()):
            y = header_y + header_h + row_index * row_h
            self.set_xy(14, y)
            fill = row_index % 2 == 1
            if fill:
                self.set_fill_color(*PALE)
            self.set_draw_color(*LINE)
            self.set_font(self.font_family, "", 6.3)
            self.set_text_color(*INK)
            for column, width, value in zip(columns, widths, row.tolist()):
                text = _format_cell(str(column), value)
                alignment = "R" if isinstance(value, (int, float)) and not isinstance(value, bool) else "L"
                self.cell(width, row_h, self.fit(text, width - 2.6), border="B", fill=fill, align=alignment)

    def draw_notes(self) -> None:
        stats = self.stats
        self.set_draw_color(*LINE)
        self.line(14, 238, 196, 238)
        self.set_xy(14, 242)
        self.set_font(self.font_family, "B", 6.8)
        self.set_text_color(*TEAL)
        self.cell(182, 4, self.t("CLEANING ACTIONS APPLIED"))

        self.set_xy(14, 247)
        self.set_font(self.font_family, "", 6.5)
        self.set_text_color(*MUTED)
        actions = (
            "Headers standardized, whitespace trimmed, exact duplicates removed, dates normalized to YYYY-MM-DD, "
            "currency text converted to numbers, and missing values repaired. "
            f"{stats.get('missing_values_remaining', 0):,} date gaps remain blank."
        )
        self.multi_cell(182, 3.6, self.t(actions), new_x="LMARGIN", new_y="NEXT")

        self.set_xy(14, 257)
        self.set_font(self.font_family, "B", 6.5)
        self.set_text_color(*TEAL)
        self.cell(182, 4, self.t("NOTES"))
        self.set_xy(14, 261)
        self.set_font(self.font_family, "", 6.2)
        self.set_text_color(*MUTED)
        notes = (
            "Numeric gaps are filled with the column median; unrecoverable totals are set to zero. "
            "Ambiguous slash dates are interpreted day-first. Amounts are displayed in the source currency (USD in the built-in demo)."
        )
        self.multi_cell(182, 3.4, self.t(notes), new_x="LMARGIN", new_y="NEXT")

    def draw_footer(self) -> None:
        self.set_draw_color(*LINE)
        self.line(14, 281, 196, 281)
        self.set_xy(14, 283)
        self.set_font(self.font_family, "", 6.8)
        self.set_text_color(*MUTED)
        self.cell(120, 5, self.t("Generated by DataClean Pro  |  Processed in memory"))
        self.set_xy(134, 283)
        self.cell(62, 5, self.t(f"Page {self.page_no()} of {{nb}}"), align="R")


def generate_pdf_report(df: pd.DataFrame, stats: dict[str, Any], dataset_name: str = "Supplier orders") -> bytes:
    """Build and return a one-page executive summary PDF as bytes."""
    report = _ExecutiveReport(dataset_name=dataset_name, stats=stats)
    report.add_page()
    report.draw_header()
    report.draw_cards()
    report.draw_overview()
    report.draw_table(df)
    report.draw_notes()
    report.draw_footer()
    return bytes(report.output())
