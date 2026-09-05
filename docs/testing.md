# Testing and performance

## Automated tests

```console
python -m pip install -e ".[dev]"
python -m unittest discover -s tests -v
pytest
```

The tests assert semantic XML attributes, cache record indices, package
relationship targets, and materialized aggregates with independently specified
expected results. They cover save/load/save cycles with multiple independent
pivots, including adding a pivot after loading. CI tests Python 3.9–3.13 with
both stdlib XML and lxml, then installs and tests the wheel and sdist.

The cache comparison test bounds equality operations instead of wall-clock time
to catch the quadratic algorithm without a machine-dependent timing threshold.

## Desktop Excel (opt-in)

On Windows with desktop Excel installed, run in PowerShell:

```powershell
python -m pip install -e . pywin32
$env:OPENPYXL_PIVOTS_EXCEL = "1"
python -m unittest discover -s tests -p test_excel.py -v
```

This launches a separate Excel instance and generates ten synthetic workbooks:
text, integer, fractional number, boolean, and date fields with blanks, on both
axes. Each file must open with normal load (no repair fallback), keep its native
pivot and expected grand total before and after refresh, and save without a
`repairLoad` marker. The Excel version/build is printed. The process quits even
if a test fails. No user workbooks are opened.

The test is skipped unless explicitly enabled; a green ordinary CI run is **not**
desktop Excel validation. This smoke test is also not exhaustive: manually
check repair dialogs/logs and the full visible layout in supported Excel versions.

## Benchmark

```console
PYTHONPATH=src python benchmarks/benchmark_builder.py --rows 2000 --columns 11 --repeats 3
```

The script reports median build and save times separately. Workbook setup is
excluded. It uses ten row labels, four column labels, and distinct strings in
every additional source column, exposing the former cache-indexing bottleneck.
Use the same script, runtime, arguments, and machine for both revisions.

This is synthetic data, not the previously reported 35,380-row survey dataset.
Results should not be interpreted as a speedup guarantee for that dataset.

Measured locally on Python 3.12.13 / openpyxl 3.1.5, with 2,000 rows and 11
columns (three repetitions, median):

| Revision | Build | Save |
| --- | ---: | ---: |
| `60203a9` baseline | 8.347 s | 0.338 s |
| Review changes | 0.229 s | 0.370 s |

The build improvement is about 36× for this high-cardinality workload. A single
35,380-row/11-column run of the review changes took 4.989 s to build and 7.254 s
to save; no matching baseline was timed at that size.

## Cache compatibility notes

The original 0.2.2 date/blank and boolean/blank corrections are retained. New
regressions also check numeric/blank fields on both axes. Blank-only and inline
numeric cache records are different cases: attributes describing shared items
need not be present when there are no shared items.

Microsoft's [SharedItems reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.shareditems)
documents blank flags, long-text declarations, and the rule against mixing date
and numeric bounds. The Office-specific boolean/text flags also follow the
desktop Excel evidence recorded with PR #1.

openpyxl 3.1.5 deduplicates caches by serialized definition equality but does not
consistently remap the referencing cache IDs. Generated independent caches now
carry per-pivot builder provenance in the standard `refreshedBy` attribute, which
keeps them distinct after loading and saving. This protects newly generated
caches; it is not a repair utility for pre-existing damaged workbooks.
