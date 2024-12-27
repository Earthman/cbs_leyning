# Torah Reading (Leyning) Calendar Generator

A Python script that generates detailed Torah reading schedules using HebCal's Leyning API and exports them to Google Sheets.

## Features

- Fetches Torah reading data for specified date ranges
- Creates formatted Google Sheets with:
  - Weekly parsha details
  - Aliyot verse ranges
  - Hebrew dates
  - Special Shabbatot
  - Page numbers in Etz Hayim (optional)
  - Override HebCal Haftarah verses (optional)
  - Weekday readings tab for daily minyan
- Handles special readings (Rosh Chodesh, Fast Days, Chol Ha-moed)
- Supports custom page number mapping and Haftarah verses via CSV

## Prerequisites

- Python 3.6+
- Google Sheets API credentials (`credentials.json`)
- Required Python packages:
  ```
  requests
  google-oauth2-client
  gspread
  pandas
  tenacity
  tqdm
  ```

## Installation

1. Clone the repository or download `leyning.py`
2. Install required packages:
   ```bash
   pip install requests google-auth-oauthlib gspread pandas tenacity tqdm
   ```
3. Place your Google Sheets API credentials file as `credentials.json` in the script directory

## Usage

Basic command:
```bash
python leyning.py START_DATE END_DATE -s SHEET_NAME -e EMAIL
```

Example:
```bash
python leyning.py 2024-01-01 2024-12-31 -s "Torah Readings 2024" -e user@example.com
```

### Arguments

- `START_DATE`: Start date in YYYY-MM-DD format
- `END_DATE`: End date in YYYY-MM-DD format
- `-s, --sheet`: Google Sheet name
- `-e, --email`: Email address to share the sheet with
- `-v, --verbose`: Enable verbose output
- `-t, --test`: Test mode - process only first parsha
- `--pages`: CSV file with page numbers

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

The script creates a Google Sheet with:
- A "Minyan" tab for weekday readings
- Individual tabs for each parsha containing:
  - Service information
  - Aliyah assignments
  - Verse ranges
  - Page numbers
  - Honor assignments

## Error Handling

- Retries API calls with exponential backoff
- Handles Google Sheets API rate limits
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
- Google Sheets API for spreadsheet functionality
– Claude.ai