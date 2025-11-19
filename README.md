# pche-crossregistration

Automate filling PCHE cross-registration PDFs from a single Excel sheet.

## Prerequisites
- Python 3.10+ recommended
- Install dependencies: `pip install -r requirements.txt`

## Input locations
- Excel data (default): `excel_input/cross_registration_data.xlsx`
- PDF templates: `pche_template/`
- Output directory (default): `Filled_PDFs/`

Paths are resolved relative to the script, so you can run it from any working directory.

## Usage
Fill PDFs with defaults:
- `python code/fill_in_pdf.py`

Provide custom paths:
- `python code/fill_in_pdf.py --excel-path /path/to/data.xlsx --output-dir /tmp/pche_outputs`

Each Excel row generates a filled PDF named `<last>-<preferred>-<home_school>.pdf`
with filename components normalized for safety. If a file with the same name exists,
the script appends `-1`, `-2`, etc. to avoid overwriting existing PDFs in the output
directory. Any row-level errors are summarized at the end without stopping the whole
run.
