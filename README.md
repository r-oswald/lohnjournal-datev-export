# DATEV Lohnjournal PDF Importer

Extract employee payroll data from DATEV Lohnjournal PDFs into SQLite and Excel.

## Features

- Parse encrypted DATEV Lohnjournal PDFs using coordinate-based extraction
- Store data in SQLite (one table per month)
- Export to Excel with monthly sheets and summary

## Extracted Fields

| Category | Fields |
|----------|--------|
| Employee Info | Pers.-Nr., Name, Steuerklasse, Faktor, Ki.Freibetrag |
| Days | St.Tg., SV.Tg. |
| Gross | Gesamtbrutto, Steuerbrutto, KV/RV/AV/PV-Brutto |
| Taxes | Lohnsteuer, Kirchensteuer, SolZ |
| Pausch. Taxes | Pausch.verst.Bezüge, Pausch.Lohnsteuer, Pausch.KiSt, Pausch.SolZ |
| Employee (AN) | KV/RV/AV/PV-Beitrag AN |
| Employer (AG) | KV/RV/AV/PV-Beitrag AG |
| Other | Umlage 1/2/U3, Nettobezüge, Auszahlungsbetrag |

## Installation

```bash
git clone https://github.com/r-oswald/lohnjournal-datev-export.git
cd lohnjournal-datev-export
python -m venv .venv
source .venv/bin/activate  # Linux/Mac (.venv\Scripts\activate on Windows)
pip install -r requirements.txt
```

## Usage

```bash
python import_all_lohnjournal.py -p /path/to/pdfs -n output_name --password PDF_PASSWORD
```

Creates:
- `output_name.db` - SQLite database
- `output_name_export.xlsx` - Excel with summary + monthly sheets

### Options

| Option | Description |
|--------|-------------|
| `-p, --pdf-folder` | Folder with PDFs (default: `./pdfs`) |
| `-d, --db` | Database output path |
| `-e, --excel` | Excel output path |
| `-n, --name` | Base name for outputs (default: `lohnjournal_complete`) |
| `-P, --password` | PDF password |

### Examples

```bash
# Basic usage
python import_all_lohnjournal.py -p ./my_pdfs -n company_2025 --password 12345

# Custom output paths
python import_all_lohnjournal.py -p ./pdfs --db ./out/data.db --excel ./out/report.xlsx -P 12345

# From ZIP
unzip "Lohnjournal_2025.zip" -d pdfs/ && python import_all_lohnjournal.py -p pdfs -n company
```

## Output

### SQLite
Tables named `lohnjournal_MONTH_YEAR` (e.g., `lohnjournal_Januar_2025`).

### Excel
- **Zusammenfassung**: All employees with totals across months
- **Monthly sheets**: Individual month data

## Technical Notes

**Coordinate-based extraction**: Uses `pdfplumber` to map values by X-position, handling empty values and variable-width columns reliably.

**DATEV number format**: `2.43000` = €2,430.00 (dot = thousands separator, last 2 digits = cents, trailing `-` = negative).

## Requirements

- Python 3.10+
- pdfplumber, pandas, openpyxl

## Compatibility

Tested with DATEV Lohnjournal exports 2024/2025. Different versions may need X-coordinate adjustments in `lohnjournal_parser.py`.

## License

MIT
