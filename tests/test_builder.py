from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

from openpyxl_pivots import PivotBuildError, add_pivot_table


def make_workbook():
    wb = Workbook()
    data = wb.active
    data.title = "Data"
    data.append(["Region", "Year", "Revenue"])
    for row in [
        ("North", 2025, 120),
        ("South", 2025, 80),
        ("North", 2026, 150),
        ("South", 2026, 95),
    ]:
        data.append(row)
    return wb


class BuilderTests(unittest.TestCase):
    def test_creates_complete_package_and_round_trips(self):
        wb = make_workbook()
        target = wb.create_sheet("Pivot")
        pivot = add_pivot_table(
            target,
            source="Data!A1:C5",
            destination="A3",
            name="SalesPivot",
            row="Region",
            column="Year",
            value="Revenue",
        )
        self.assertEqual(pivot.name, "SalesPivot")
        self.assertEqual(target["A3"].value, "Sum of Revenue")
        self.assertEqual(target["B3"].value, "Year")
        self.assertEqual(target["B5"].value, 120)
        self.assertEqual(target["C5"].value, 150)
        self.assertEqual(target["D7"].value, 445)

        with TemporaryDirectory() as directory:
            output = Path(directory) / "pivot.xlsx"
            wb.save(output)
            with ZipFile(output) as archive:
                names = set(archive.namelist())
                self.assertIn("xl/pivotTables/pivotTable1.xml", names)
                self.assertIn("xl/pivotCache/pivotCacheDefinition1.xml", names)
                self.assertIn("xl/pivotCache/pivotCacheRecords1.xml", names)
                self.assertIn(
                    b'refreshOnLoad="1"',
                    archive.read("xl/pivotCache/pivotCacheDefinition1.xml"),
                )
                self.assertIn(
                    b'name="SalesPivot"', archive.read("xl/pivotTables/pivotTable1.xml")
                )
                self.assertIn(
                    b'<location ref="A3:D7" firstHeaderRow="1" firstDataRow="2"',
                    archive.read("xl/pivotTables/pivotTable1.xml"),
                )
                self.assertIn(
                    b'name="SalesPivot" cacheId="1"',
                    archive.read("xl/pivotTables/pivotTable1.xml"),
                )
                self.assertIn(
                    b'<i t="data" r="0" i="0"><x/></i>',
                    archive.read("xl/pivotTables/pivotTable1.xml"),
                )
                cache_xml = archive.read("xl/pivotCache/pivotCacheDefinition1.xml")
                self.assertNotIn(b'saveData=', cache_xml)
                self.assertIn(b'<cacheField name="Region" numFmtId="0">', cache_xml)

            reloaded = load_workbook(output)
            self.assertEqual(len(reloaded["Pivot"]._pivots), 1)
            self.assertEqual(reloaded["Pivot"]._pivots[0].name, "SalesPivot")
            self.assertEqual(reloaded["Pivot"]._pivots[0].cache.recordCount, 4)

    def test_row_only_average(self):
        wb = make_workbook()
        target = wb.create_sheet("Pivot")
        add_pivot_table(
            target,
            source=(wb["Data"], "A1:C5"),
            destination="B2",
            name="AveragePivot",
            row="Region",
            value="Revenue",
            aggregation="average",
        )
        self.assertEqual(target["C3"].value, 135)
        self.assertEqual(target["C4"].value, 87.5)
        with TemporaryDirectory() as directory:
            output = Path(directory) / "average.xlsx"
            wb.save(output)
            pivot = load_workbook(output)["Pivot"]._pivots[0]
            self.assertEqual(pivot.dataFields[0].subtotal, "average")

    def test_rejects_non_numeric_sum(self):
        wb = make_workbook()
        target = wb.create_sheet("Pivot")
        with self.assertRaisesRegex(PivotBuildError, "requires numeric"):
            add_pivot_table(
                target,
                source="Data!A1:C5",
                destination="A1",
                name="BadPivot",
                row="Year",
                value="Region",
                aggregation="sum",
            )


if __name__ == "__main__":
    unittest.main()
