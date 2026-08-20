# Changelog

All notable changes to this project are documented here.

## 0.2.0 - 2026-08-20

- Correct the PivotTable location for Excel's two-header-row layout.
- Use positive pivot-cache identifiers.
- Match Excel's axis-item encoding and simplify cache-field attributes.
- Add Microsoft Excel validation following the corrections.

## 0.1.0 - 2026-08-20

- Add the initial worksheet-backed PivotTable builder.
- Support one row field, an optional column field, one value field, and six
  aggregation modes.
- Materialize the initial result and mark the cache for refresh on load.
