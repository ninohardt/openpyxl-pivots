# Proposal: native API for creating PivotTables

## Summary

Would the openpyxl maintainers consider accepting a deliberately limited API
for creating ordinary worksheet-backed PivotTables?

openpyxl already contains the OOXML models, serializers, relationship handling,
and package writer needed to preserve PivotTables. The missing layer is a public
builder that derives a consistent cache and table definition from worksheet data.

An external proof of concept now creates a PivotTable with:

- one row field;
- zero or one column field;
- one value field;
- `sum`, `count`, `countNums`, `average`, `min`, and `max` aggregation;
- a worksheet-range source;
- cached records and a materialized initial result; and
- `refreshOnLoad` so Excel can rebuild the view.

The generated package round-trips through openpyxl, has been smoke-tested in
LibreOffice, and opens successfully in desktop Microsoft Excel without a repair
warning for the supported test case.

## Possible API

```python
ws.add_pivot_table(
    source="Data!A1:C100",
    destination="A3",
    name="SalesPivot",
    row="Region",
    column="Year",
    value="Revenue",
    aggregation="sum",
)
```

The first upstream version could intentionally reject unsupported combinations
instead of attempting to expose the complete PivotTable schema.

## Questions for maintainers

1. Is creation support within openpyxl's intended scope?
2. If not, would maintainers accept stable lower-level hooks needed by an
   extension package?
3. If it is in scope, should the initial API live on `Worksheet`, in a dedicated
   builder module, or at a lower OOXML level?
4. Which Excel versions and workbook fixtures would be required for acceptance?

## Proposed initial acceptance boundary

- Worksheet-range sources only.
- One row field, optional one column field, and one value field.
- No filters, grouping, calculated fields, slicers, timelines, PivotCharts,
  OLAP sources, external sources, or Data Model support.
- Explicit validation errors for every unsupported request.
- Tests for XML fragments, package relationships, multiple caches, save/load
  round trips, and representative strings, numbers, dates, and blanks.
- Documentation that generated files should be validated in the target Excel
  environment while creation support remains experimental.

## Work still required before a merge request

- Integrate the public API without relying on private worksheet attributes.
- Exercise multiple PivotTables and cache identifiers in one workbook.
- Add broader source-type and missing-value fixtures.
- Validate generated XML against the published OOXML schema where practical.
- Run the full openpyxl test matrix and documentation build.

The proof-of-concept package can remain an incubation and compatibility layer
while the scope and API are discussed.
