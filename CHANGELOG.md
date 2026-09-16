# Changelog

All notable changes to this project are documented here.

## Unreleased

- Keep independent pivot caches distinct across repeated openpyxl save/load
  cycles by recording per-pivot builder provenance in `refreshedBy`.
- Use consistent numeric/date keys for both cache indexing and aggregation;
  equivalent `int`/`float` and `date`/`datetime` labels no longer lose values.
- Replace quadratic cache scans and repeated column-total scans with lookups.
- Correct numeric-plus-blank cache flags, declare long text, and avoid combining
  numeric and date bounds in mixed caches.
- Reject malformed ranges, cross-workbook sources, merged/occupied output,
  non-finite numbers, timezone-aware dates, formulas, and Excel errors explicitly.
  Formula/error rejection replaces the previous conversion to cache text.
- Preserve quoted sheet names and literal formula-like text labels.
- Expand semantic XML, aggregation, package-relationship, and failure tests;
  exercise both XML backends and installed distribution artifacts in CI.
- Add an opt-in desktop Excel open/refresh/save test and synthetic benchmark.

## 0.2.2 - 2026-08-21

- Route date fields containing blanks through the mixed-item cache path so the
  generated cache declares both `containsBlank` and `containsSemiMixedTypes`.
- Mark boolean shared items as strings, matching the attributes Microsoft Excel
  requires when a boolean field also contains blanks.
- Add regression coverage for date-plus-blank and boolean-plus-blank axis fields.

## 0.2.1 - 2026-08-21

- Mark text-or-blank shared-item caches as semi-mixed so Microsoft Excel can
  open PivotTables whose row or column field contains blank values.
- Raise a descriptive `PivotBuildError` when `row`, `column`, or `value` is not
  a single field name instead of leaking an internal unhashable-type error.
- Add regression coverage for blank categorical axis values.

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
