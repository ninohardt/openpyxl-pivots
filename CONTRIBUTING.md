# Contributing

Contributions are welcome, especially small changes accompanied by a workbook
fixture or a precise description of the Excel behavior being implemented.

## Development setup

```console
python -m venv .venv
python -m pip install -e ".[dev]"
pytest
```

Before submitting a change:

1. Add tests for both the generated OOXML and the openpyxl save/load round trip.
2. Open the resulting workbook in desktop Excel and report the Excel version.
3. Confirm that Excel displays no repair warning or recovery log.
4. Do not commit workbooks containing confidential or personal data.

Changes that expand the supported PivotTable model should update the feature
limits and validation status in the README.
