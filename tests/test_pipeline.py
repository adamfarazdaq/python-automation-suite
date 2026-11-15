"""End-to-end checks for the sample -> clean -> Excel/PDF pipeline."""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import load_workbook

from cleaner import clean_dataframe, dataframe_to_excel_bytes
from report_generator import generate_pdf_report
from sample_data import generate_sample_data


class PipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.raw = generate_sample_data()
        cls.cleaned, cls.stats = clean_dataframe(cls.raw)

    def test_sample_is_deliberately_messy(self) -> None:
        self.assertGreater(self.raw.duplicated().sum(), 0)
        self.assertIn("  order id  ", self.raw.columns)
        self.assertGreater(self.raw.isna().sum().sum(), 0)
        self.assertTrue(self.raw["unit price ($)"].astype(str).str.contains(r"\$").any())
        self.assertTrue(self.raw["order_date"].astype(str).str.contains(r"[/,-]").any())

    def test_cleaning_contract(self) -> None:
        self.assertGreater(self.stats["duplicates_removed"], 0)
        self.assertGreater(self.stats["missing_values_fixed"], 0)
        self.assertGreater(self.stats["formatting_issues_fixed"], 0)
        self.assertFalse(self.cleaned.duplicated().any())
        self.assertIn("Order Date", self.cleaned.columns)
        self.assertNotIn("  order id  ", self.cleaned.columns)

        dates = self.cleaned["Order Date"].dropna()
        self.assertTrue(dates.str.fullmatch(r"\d{4}-\d{2}-\d{2}").all())
        self.assertTrue(all(not self.cleaned[column].isna().any() for column in ("Supplier", "Category", "Status")))

        supplier_text = " ".join(self.cleaned["Supplier"].dropna().astype(str))
        self.assertTrue(any(ord(character) > 127 for character in supplier_text))

    def test_excel_and_pdf_outputs(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            output_dir = Path(temporary_directory)
            xlsx_path = output_dir / "cleaned.xlsx"
            pdf_path = output_dir / "summary.pdf"

            xlsx_path.write_bytes(dataframe_to_excel_bytes(self.cleaned))
            pdf_path.write_bytes(generate_pdf_report(self.cleaned, self.stats, "Pipeline test"))

            self.assertGreater(xlsx_path.stat().st_size, 4_000)
            self.assertTrue(pdf_path.read_bytes().startswith(b"%PDF"))
            self.assertGreater(pdf_path.stat().st_size, 2_000)

            workbook = load_workbook(xlsx_path)
            worksheet = workbook.active
            self.assertEqual(worksheet["A1"].value, "Order ID")
            self.assertTrue(worksheet["A1"].font.bold)
            self.assertEqual(worksheet.freeze_panes, "A2")

            # The temporary context removes both generated files when it exits.
            self.assertTrue(xlsx_path.exists())
            self.assertTrue(pdf_path.exists())

        self.assertFalse(xlsx_path.exists())
        self.assertFalse(pdf_path.exists())

    def test_pdf_handles_unicode_and_unsupported_glyphs(self) -> None:
        unicode_data = self.cleaned.copy()
        unicode_data.loc[unicode_data.index[0], "Supplier"] = "Müller Pharma GmbH — Café € 漢字 😀"
        unicode_data.loc[unicode_data.index[0], "Product"] = "Vitamine D3 – 1000 IU"

        payload = generate_pdf_report(unicode_data, self.stats, "Unicode test — café 😀")
        self.assertTrue(payload.startswith(b"%PDF"))
        self.assertGreater(len(payload), 2_000)

        # Simulate a host with no usable TrueType font (core Helvetica fallback).
        with patch("report_generator._font_pairs", return_value=[]):
            fallback_payload = generate_pdf_report(unicode_data, self.stats, "Unicode test — café 😀")
        self.assertTrue(fallback_payload.startswith(b"%PDF"))
        self.assertGreater(len(fallback_payload), 2_000)


if __name__ == "__main__":
    unittest.main(verbosity=2)
