# openpyxl-pivots

Experimental native creation of ordinary worksheet-backed PivotTables with
openpyxl 3.1.5.

`openpyxl` can preserve existing PivotTables, but it does not expose a public
API for creating them. `openpyxl-pivots` builds the required object graph and
uses openpyxl's existing OOXML serializers and package writer.

> [!WARNING]
> This is an alpha package with a deliberately narrow feature set. Keep a
> backup of important workbooks and validate generated files in the versions
> of Excel used by your application.

## Installation

Directly from GitHub:

```console
python -m pip install git+https://github.com/ninohardt/openpyxl-pivots.git
```

From a source checkout:

```console
python -m pip install .
```

For development:

```console
python -m pip install -e ".[dev]"
pytest
```

This MVP intentionally supports one row field, zero or one column field,
and one data field. It writes the PivotTable definition, pivot cache definition,
cache records, package relationships, and a materialized result. The cache is
marked `refreshOnLoad`, so Excel can rebuild the view from the source range.

## Example

```python
from openpyxl import Workbook
from openpyxl_pivots import add_pivot_table

wb = Workbook()
data = wb.active
data.title = "Data"
data.append(["Region", "Year", "Revenue"])
data.append(["North", 2025, 120])
data.append(["South", 2025, 80])
data.append(["North", 2026, 150])

pivot_sheet = wb.create_sheet("Pivot")
add_pivot_table(
    pivot_sheet,
    source="Data!A1:C4",
    destination="A3",
    name="SalesPivot",
    row="Region",
    column="Year",
    value="Revenue",
    aggregation="sum",
)
wb.save("sales-pivot.xlsx")
```

You can optionally install a worksheet convenience method:

```python
from openpyxl_pivots import install

install()
pivot_sheet.add_pivot_table(
    source="Data!A1:C4",
    destination="A3",
    name="SalesPivot",
    row="Region",
    value="Revenue",
)
```

## Current limits

- one row field
- at most one column field
- one value field
- aggregations: `sum`, `count`, `countNums`, `average`, `min`, `max`
- worksheet ranges only; no external, OLAP, or Data Model sources
- no grouping, calculated fields, slicers, timelines, or PivotCharts
- formulas in source cells are cached as their formula strings unless the
  workbook was loaded with cached values; use literal source data for now

## Validation status

Version 0.2.1 is covered by package-level tests, structurally round-tripped by
openpyxl, smoke-tested with LibreOffice, and manually opened successfully in
desktop Microsoft Excel without a repair warning. The XML now matches the
two-header-row geometry, positive cache identifiers, and axis-item encoding
observed in an Excel-normalized workbook. Text axis fields containing blanks
are emitted with Excel-compatible shared-item flags; this was validated both on
a minimal reproducer and on a 35,380-row workbook in desktop Excel on Windows.

This validation is evidence for the supported example, not a general Excel
compatibility guarantee. Broader field combinations, data types, and Excel
versions still need coverage.

Large, wide source ranges can currently take significant time to build because
cache construction scales per source cell. Profile representative workloads
before using the package in latency-sensitive request/response paths.

## Project direction

The package is an incubation layer for PivotTable creation. A focused proposal
for eventual openpyxl integration is available in
[`docs/upstream-proposal.md`](docs/upstream-proposal.md). The extension remains
useful as an experimental high-level builder and as a potential compatibility
layer even if full creation support is not accepted upstream.

## License

MIT
