# Torah Reading (Leyning) Calendar Generator

A Python script that generates detailed Torah reading schedules using HebCal's Leyning API and exports them to a local Excel (`.xlsx`) workbook.

## Features

- Fetches Torah reading data for specified date ranges
- Creates a formatted `.xlsx` workbook with:
  - Weekly parsha details
  - Aliyot verse ranges (with Virtual Tikkun hyperlinks)
  - Hebrew dates
  - Special Shabbatot
  - Page numbers in Etz Hayim (optional)
  - Override HebCal Haftarah verses (optional)
  - Weekday readings tab for daily minyan
- Handles special readings (Rosh Chodesh, Fast Days, Chol Ha-moed)
- Supports custom page number mapping and Haftarah verses via CSV
- Generated entirely locally — no Google account, credentials, or network
  beyond the single HebCal data fetch (which can also be supplied offline)
- Editable `.xlsx` template for the Header/Footer layout, with a built-in
  fallback layout if no template is present

## Prerequisites

- Python 3.6+
- Required Python packages:
  ```
  requests
  openpyxl
  pandas
  tenacity
  ```

## Installation

1. Clone the repository or download `leyning.py`
2. Install required packages:
   ```bash
   pip install requests openpyxl pandas tenacity
   ```

## Usage

Basic command:
```bash
python leyning.py START_DATE END_DATE -s OUTPUT.xlsx
```

Example:
```bash
python leyning.py 2025-04-01 2026-03-31 -s leyning_5786.xlsx --pages page_numbers_and_haftarot.csv
```

### Arguments

- `START_DATE`: Start date in YYYY-MM-DD format
- `END_DATE`: End date in YYYY-MM-DD format
- `-s, --sheet`: Output `.xlsx` path (the `.xlsx` extension is added if omitted)
- `-v, --verbose`: Enable verbose output
- `-t, --test`: Test mode - process only first parsha
- `--pages`: CSV file with page numbers
- `--scroll`: Name of the Torah scroll (default: `Gunther`)
- `--template`: Path to a local `.xlsx` template (default: `template.xlsx`
  beside the script). Falls back to the built-in layout if missing/invalid.
- `--json`: Read HebCal leyning JSON from a local file instead of calling the
  API (useful offline or for repeatable runs)
- `--make-template PATH`: Write a fresh starter template to `PATH` and exit

If `-s/--sheet` is omitted, the fetched HebCal JSON is printed to stdout.

### Template

The Header and Footer layout lives in `template.xlsx` (sheets named `Header`
and `Footer`). It is a true format template — both cell values and styling
(fills, font sizes) are copied into every parsha sheet. Edit it in any
spreadsheet program to change the layout without touching the code.

Dynamic positioning is driven by marker cells, so rows can be added or moved:

- **Header**: `Torah(s) Scroll` in column A marks the scroll row;
  `Reader` / `Aliyah` / `Hebrew Name(s)` / `Notes` in columns C–F mark the
  last header row (aliyot begin on the next row).
- **Footer**: `Etz Hayyim` in column D marks the honors header; the Torah and
  Haftarah page numbers are placed on the two rows below it.

Regenerate the default template at any time:
```bash
python leyning.py --make-template template.xlsx
```

If the template file is missing or its markers can't be found, the script
prints a warning and uses an equivalent built-in hardcoded layout.

### Page Numbers CSV Format

Create a CSV with columns:
- `Parsha`: Parsha name
- `Torah Page`: Torah reading page number
- `Haftarah Page`: Haftarah page number
- `Haftarah verses`: Haftarah verse reference

Example:
```csv
Parsha,Torah Page,Haftarah Page,Haftarah verses
Bereishit,3,36,Isaiah 42:5-43:10
```

## Output Format

The script creates an `.xlsx` workbook with:
- A "Minyan" tab for weekday readings
- Individual tabs for each parsha containing:
  - Service information
  - Aliyah assignments
  - Verse ranges
  - Page numbers
  - Honor assignments

Hyperlink and conditional formulas (Virtual Tikkun links, the Musaf-leader
default) are written as real Excel formulas and are evaluated when the file is
opened in Excel or LibreOffice.

## Error Handling

- Retries the HebCal API call with exponential backoff
- Falls back to the built-in layout if the template is unusable
- Validates date formats
- Reports errors verbosely with `-v` flag

## Contributing

Submit issues and pull requests on GitHub. Please include:
- Clear description of changes/issues
- Test cases for new features
- Updated documentation as needed

## License

This project uses the HebCal API which has its own terms of service. Please review [HebCal's terms](https://www.hebcal.com/home/terms) before use.

## Acknowledgments

- HebCal for providing the Leyning API
- openpyxl for `.xlsx` generation
– Claude.ai
