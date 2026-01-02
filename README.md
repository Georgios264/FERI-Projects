# Exposure ingestion helper

This repository contains a small Python helper that automates copying exposure tables from
raw fund spreadsheets into your existing Excel template. Configure each fund once, and
re-run monthly without manually selecting and pasting ranges.

## How it works
- YAML profiles (`config/fund_profiles.yaml`) describe how to locate the exposure table in a
  raw workbook and where to paste it into the template. Each profile must specify **either**
  a fixed `range` **or** a `start_cell` with bounds, not both.
- The script reads only cell values (no formatting) and can clear the target block before
  writing new values.
- Supports fixed ranges (e.g., `B14:H29`) or dynamic reading from a start cell with a
  configurable maximum size. Trailing blank rows/columns are trimmed automatically.

## Quickstart
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Adjust `config/fund_profiles.yaml` to match your funds. Example for Berry Street:
   ```yaml
   funds:
     berry_street:
       source:
         sheet: "Exposure"      # sheet in the raw file
         range: "B14:H29"       # rectangle that holds the exposure table
       target:
         sheet: "Exposure_Input" # sheet in your template
         start_cell: "B6"        # top-left cell of the yellow input block
         clear_rows: 25
         clear_cols: 7
   ```
3. Run the copier:
   ```bash
   python src/auto_copy.py \
       --raw /path/to/raw.xlsx \
       --template /path/to/template.xlsx \
       --fund berry_street \
       --output /path/to/output.xlsx
   ```
   - Add `--dry-run` to preview the detected table (first 5 rows, row/column counts) without
     writing a file. Useful for testing a new fund layout.
   - Use `--list-funds` to show all configured profiles and exit.
4. Open the generated `output.xlsx` and run your existing VBA macro or Selenium uploader.

## Config tips
- Use `range` for layouts that never change; use `start_cell`+`stop_at_blank_rows` for tables
  that vary in length.
- Set `clear_rows`/`clear_cols` to wipe previous data in the template block before writing.
- Add more funds under the `funds:` key; each name becomes the `--fund` argument.

## Notes
- The script raises clear errors when sheets or profiles are missing.
- `data_only=True` is used when reading the raw workbook so formulas are resolved to values.
