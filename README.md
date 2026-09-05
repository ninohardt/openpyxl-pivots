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
- literal strings, finite numbers, booleans, and timezone-naive dates/datetimes
- formulas and Excel error cells are rejected with a cell-specific error;
  openpyxl does not calculate formulas. Use literal data or load an already
  calculated workbook with `data_only=True` (missing formula caches appear blank)
- source and target must be normal worksheets in the same workbook
- output must fit within Excel's sheet limits and avoid occupied or merged cells

## Validation status

Version 0.2.2 is covered by package-level tests, structurally round-tripped by
openpyxl, smoke-tested with LibreOffice, and manually opened successfully in
desktop Microsoft Excel without a repair warning. The XML now matches the
two-header-row geometry, positive cache identifiers, and axis-item encoding
observed in an Excel-normalized workbook. Text, date, and boolean axis fields
containing blanks are emitted with Excel-compatible shared-item flags. This was
validated across 14 field shapes and on a 35,380-row workbook in desktop Excel
on Windows.

This validation is evidence for the supported example, not a general Excel
compatibility guarantee. Broader field combinations, data types, and Excel
versions still need coverage.

The unreleased changes have automated regression coverage for numeric blanks,
equivalent numeric/date grouping keys, all six aggregations, invalid inputs,
multiple pivots, and repeated save/load cycles. These changes still need desktop
Excel validation; the 0.2.2 validation above does not certify newer code.

Cache indexing now uses keyed lookups instead of repeated scans of distinct
values. Construction still retains all source cells/cache records in memory,
and rendering costs grow with the number of row groups times column groups.
Profile representative workloads before using this in a request/response path.
See [testing and benchmarks](docs/testing.md) for the repeatable benchmark and
the opt-in Windows/desktop Excel test.

## Project direction

The package is an incubation layer for PivotTable creation. A focused proposal
for eventual openpyxl integration is available in
[`docs/upstream-proposal.md`](docs/upstream-proposal.md). The extension remains
useful as an experimental high-level builder and as a potential compatibility
layer even if full creation support is not accepted upstream.

## License

MIT
