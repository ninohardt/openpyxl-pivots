from openpyxl import Workbook

from openpyxl_pivots import add_pivot_table


wb = Workbook()
data = wb.active
data.title = "Data"
data.append(["Region", "Year", "Revenue"])
for row in [
    ("North", 2025, 120),
    ("South", 2025, 80),
    ("North", 2026, 150),
    ("South", 2026, 95),
    ("West", 2025, 110),
    ("West", 2026, 125),
]:
    data.append(row)

pivot = wb.create_sheet("Pivot")
add_pivot_table(
    pivot,
    source="Data!A1:C7",
    destination="A3",
    name="SalesPivot",
    row="Region",
    column="Year",
    value="Revenue",
    aggregation="sum",
)
wb.save("sales-pivot.xlsx")
