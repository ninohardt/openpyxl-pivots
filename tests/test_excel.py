"""Opt-in Windows/desktop Excel smoke test. Never runs in ordinary CI."""
from datetime import date
import os
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from test_regressions import build, NS


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("OPENPYXL_PIVOTS_EXCEL") == "1",
                     "requires Windows desktop Excel and OPENPYXL_PIVOTS_EXCEL=1")
class ExcelTests(unittest.TestCase):
    def test_open_refresh_save_blank_axis_matrix(self):
        import win32com.client

        excel = win32com.client.DispatchEx("Excel.Application")
        try:
            excel.Visible = False
            excel.DisplayAlerts = False
            print(f"Desktop Excel {excel.Version}, build {excel.Build}")
            with TemporaryDirectory() as directory:
                for member in ("North", 1, 1.5, True, date(2025, 1, 1)):
                    for axis in ("row", "col"):
                        with self.subTest(member=member, axis=axis):
                            rows = [(member, "X", 10), (None, "X", 20)]
                            if axis == "col":
                                rows = [(c, r, v) for r, c, v in rows]
                            wb, _, _ = build(rows, refresh_on_load=False)
                            path = Path(directory) / "generated.xlsx"
                            saved = Path(directory) / "excel-saved.xlsx"
                            if saved.exists():
                                saved.unlink()
                            wb.save(path)
                            book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                                        CorruptLoad=0, AddToMru=False)
                            try:
                                pivots = book.Worksheets("Pivot").PivotTables()
                                self.assertEqual(pivots.Count, 1)
                                pivot = pivots.Item(1)
                                self.assertEqual(pivot.TableRange2.Value[-1][-1], 30)
                                self.assertTrue(pivot.RefreshTable())
                                self.assertEqual(pivot.TableRange2.Value[-1][-1], 30)
                                book.SaveAs(str(saved), FileFormat=51)
                            finally:
                                book.Close(SaveChanges=False)
                            with ZipFile(saved) as archive:
                                root = ET.fromstring(archive.read("xl/workbook.xml"))
                                for recovery in root.findall("x:fileRecoveryPr", NS):
                                    self.assertNotIn(recovery.get("repairLoad"), ("1", "true"))
        finally:
            excel.Quit()
