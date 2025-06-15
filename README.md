# measurement-handler

A small system for automatic detection of parts classified as good/bad, according to measurement tables.

This repository includes a simple script `ods_generator.py` which converts a CSV file into an ODS spreadsheet. Cells are colored depending on their value:

- `Y` &rarr; green background
- `NM` &rarr; black background with white text
- numbers &rarr; red background

When running the script, it prints the detected row and column count and allows you to choose how many rows and columns to show as a preview.

## Usage

```
python ods_generator.py input.csv output.ods
```

The script is self-contained and does not require external dependencies.
