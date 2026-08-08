"""Quick check: write_report must not emit empty inlineStr cells."""
from __future__ import annotations

import re
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

import main as mm


class WriteReportExcelTests(unittest.TestCase):
    def test_no_empty_inline_str_and_opens(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            hits = [
                mm.Hit(
                    "тройка",
                    "https://example.com/a  ",
                    "2026-08-01",
                    "Title\x01x",
                    datetime(2026, 8, 8, 12, 0, 0),
                )
            ]
            path = mm.write_report(
                hits,
                Path(td),
                datetime(2026, 8, 8, 12, 0, 0),
                3.0,
                unavailable_sites=["https://bad.example  "],
                commented_sites=["https://old.example"],
            )
            with zipfile.ZipFile(path) as zf:
                sheet = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
            # Broken pattern that triggers Excel repair:
            self.assertIsNone(
                re.search(r't="inlineStr"></c>', sheet),
                msg="empty inlineStr cells break Excel",
            )
            self.assertIn("HYPERLINK", sheet)
            # Trailing spaces must be stripped from link targets
            self.assertNotIn("example.com/a  ", sheet)
            wb = __import__("openpyxl").load_workbook(path)
            self.assertGreaterEqual(wb.active.max_row, 3)


if __name__ == "__main__":
    unittest.main()
