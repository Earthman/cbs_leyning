import requests
import json
from datetime import datetime, timedelta
import sys
import argparse
import gspread
from google.oauth2.service_account import Credentials
from collections import defaultdict
import time
from tenacity import retry, stop_after_attempt, wait_exponential
from tqdm import tqdm
import pandas as pd

def set_column_widths(worksheet, verbose=False):
    """Set the width of columns to match the template sheet."""
    if verbose:
        print("Setting column widths...")
    
    # Define column widths (in pixels) matched to template sheet
    column_widths = [
        ('A', 113),  # First column - for labels
        ('B', 233),  # Second column - for parsha names and service parts
        ('C', 184),  # Third column - for assignee names
        ('D', 184),  # Fourth column - for dates and page numbers
        ('E', 184),  # Fifth column - for Hebrew names
        ('F', 442),  # Sixth column - for notes
    ]
    
    # Prepare the batch update request
    requests = []
    for col, width in column_widths:
        col_index = ord(col) - ord('A')  # Convert column letter to 0-based index
        requests.append({
            'updateDimensionProperties': {
                'range': {
                    'sheetId': worksheet.id,
                    'dimension': 'COLUMNS',
                    'startIndex': col_index,
                    'endIndex': col_index + 1
                },
                'properties': {
                    'pixelSize': width
                },
                'fields': 'pixelSize'
            }
        })
    
    # Execute the batch update
    worksheet.spreadsheet.batch_update({'requests': requests})
    time.sleep(1)  # Small delay to respect rate limits

def int_to_roman(num):
    """Convert integer to Roman numeral."""
    roman_symbols = [
        ('M', 1000), ('CM', 900), ('D', 500), ('CD', 400),
        ('C', 100), ('XC', 90), ('L', 50), ('XL', 40),
        ('X', 10), ('IX', 9), ('V', 5), ('IV', 4), ('I', 1)
    ]
    result = ''
    for symbol, value in roman_symbols:
        while num >= value:
            result += symbol
            num -= value
    return result

def format_verse_range(aliyah):
    """Format verse range with verse count."""
    try:
        book = aliyah['k']
        
        # Parse beginning verse reference
        start_parts = aliyah['b'].split(':')
        start_chapter = start_parts[0]
        start_verse = start_parts[1]
        
        # Parse ending verse reference
        end_parts = aliyah['e'].split(':')
        end_chapter = end_parts[0]
        end_verse = end_parts[1]
        
        # Format the range based on whether chapters are the same
        if start_chapter == end_chapter:
            verse_range = f"{start_chapter}:{start_verse}-{end_verse}"
        else:
            verse_range = f"{start_chapter}:{start_verse}-{end_chapter}:{end_verse}"
        
        # Add verse count if available
        if 'v' in aliyah:
            return f"{book} {verse_range} ({aliyah['v']})"
        else:
            return f"{book} {verse_range}"
            
    except (KeyError, IndexError, AttributeError) as e:
        print(f"Error formatting verse range: {e}")
        print(f"Aliyah data: {aliyah}")
        return "Error formatting verse range"

def get_reading_type(name):
    """Determine the type of reading based on the name."""
    name_lower = name.lower()
    if 'fast' in name_lower or 'taanit' in name_lower:
        return 'fast_day'
    elif 'rosh chodesh' in name_lower:
        return 'rosh_chodesh'
    elif 'chol ha-moed' in name_lower or 'chol hamoed' in name_lower:
        return 'chol_hamoed'
    return 'regular'

def is_special_day(name):
    """Check if this is a special day that should be included in minyan readings."""
    name_lower = name.lower()
    return any(term in name_lower for term in [
        'rosh chodesh',
        'chol ha-moed',
        'chol hamoed',
        'fast',
        'taanit'
    ])

def load_page_numbers(csv_path):
    """Load page numbers from CSV file."""
    import pandas as pd
    
    df = pd.read_csv(csv_path)
    # Rename column if old spelling exists
    if 'Haftara verses' in df.columns:
        df = df.rename(columns={'Haftara verses': 'Haftarah verses'})
    return df.set_index('Parsha').to_dict('index')


def write_header(worksheet, parsha_data, verbose=False):
    """Write header section (rows 1-14) with support for special Shabbats."""
    # Parse and format dates
    full_date = datetime.strptime(parsha_data['date'], '%Y-%m-%d')
    gregorian_date = full_date.strftime('%B %-d')  # e.g., "January 4"
    previous_date = (full_date - timedelta(days=1)).strftime('%B %-d')  # For Kabbalat Shabbat
    
    # Parse Hebrew date to get just month and day
    hebrew_date_parts = parsha_data['hdate'].split()  # e.g., "26 Tevet 5784"
    hebrew_date = f"{hebrew_date_parts[1]} {hebrew_date_parts[0]}"  # e.g., "Tevet 26"
    
    # Check for special Shabbat
    special_shabbat = None
    # Check top-level reason.haftara
    if isinstance(parsha_data.get('reason'), dict):
        special_shabbat = parsha_data['reason'].get('haftara')
    # Check haft.reason if no top-level reason found
    if not special_shabbat and 'haft' in parsha_data:
        haft = parsha_data['haft']
        if isinstance(haft, dict):
            special_shabbat = haft.get('reason')
        elif isinstance(haft, list):
            # If it's a list, check each haftarah entry for a reason
            for h in haft:
                if isinstance(h, dict) and 'reason' in h:
                    special_shabbat = h['reason']
                    break
    
    # Calculate verse counts for Row 14
    total_verses = 0
    parsha_verses = 0
    if 'fullkriyah' in parsha_data:
        for key, aliyah in parsha_data['fullkriyah'].items():
            if key != 'M':  # Don't include Maftir in parsha verses
                verses = aliyah.get('v', 0)
                total_verses += verses
                parsha_verses += verses
            elif key == 'M':  # Add Maftir to total but not parsha verses
                total_verses += aliyah.get('v', 0)
    
    # Prepare header data with special Shabbat in D2 if present
    header_data = [
        ["", parsha_data['name']['en'], "", gregorian_date, hebrew_date],  # Row 1
        ["Rabbi Amanda Russell", "", "", special_shabbat if special_shabbat else "", ""],  # Row 2 - Add special Shabbat
        ["Service leaders", f"Kabbalat Shabbat {previous_date}", "", "", ""],  # Row 3
        ["", "P'sukei D'zimrah", "", "", ""],  # Row 4
        ["", "Shacharit", "", "", ""],  # Row 5
        ["", "Musaf", "", "", ""],  # Row 6 (formula will be added separately)
        ["", "Torah Service", "", "", ""],  # Row 7
        ["", "Gabbai", "Sam (default)", "", ""],  # Row 8
        ["", "Distribute honors", "Todd (default)", "", ""],  # Row 9
        ["", "Read announcements", "Jerilyn (default)", "", ""],  # Row 10
        ["Board hosts", "", "", "", ""],  # Row 11
        ["", "", "", "", ""],  # Row 12
        ["Torah(s) Scroll", "Neuhas", "", "", ""],  # Row 13
        ["", f"Full kriyah - {total_verses} verses (parsha={parsha_verses})", "Reader", "Aliyah", "Hebrew Name(s)", "Notes"]  # Row 14
    ]
    
    # Batch update all data
    worksheet.batch_update([{
        'range': 'A1:F14',
        'values': header_data
    }])
    time.sleep(5)  # Respect rate limits
    
    # Update the formula cell separately using update_acell
    formula = '=if(ISNUMBER(SEARCH("Richman",$A$2)), "RDR default", "RAR default")'
    worksheet.update_acell('C6', formula)
    time.sleep(5)  # Respect rate limits
    
    # Apply formatting
    formats = [
        # Row 1 - 24pt font
        {
            'range': 'A1:F1',
            'format': {'textFormat': {'fontSize': 24}}
        },
        # Row 2 - 14pt font (including special Shabbat in D2)
        {
            'range': 'A2:F2',
            'format': {'textFormat': {'fontSize': 14}}
        },
        # Row 3 "Service leaders" - gray background
        {
            'range': 'A3',
            'format': {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}
        },
        # Row 11 "Board hosts" - gray background
        {
            'range': 'A11',
            'format': {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}
        },
        # Row 13 "Torah(s) Scroll" and "Neuhas" - orange background
        {
            'range': 'A13:B13',
            'format': {'backgroundColor': {'red': 1.0, 'green': 0.8, 'blue': 0.6}}
        },
        # Row 14 - gray background for all cells
        {
            'range': 'A14:F14',
            'format': {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}
        }
    ]
    
    # Apply each format
    for format_spec in formats:
        worksheet.format(format_spec['range'], format_spec['format'])
        time.sleep(5)  # Respect rate limits
                      
def write_aliyot(worksheet, fullkriyah, parsha_data, page_numbers=None):
   """Write aliyot section (rows 15-23)."""
   if not fullkriyah:
       return
   
   total_verses = 0
   parsha_verses = 0
   if fullkriyah:
       for key, aliyah in fullkriyah.items():
           if key != 'M':
               verses = aliyah.get('v', 0)
               total_verses += verses
               parsha_verses += verses
           elif key == 'M':
               total_verses += aliyah.get('v', 0)
   
   worksheet.update(
       values=[[
           "",
           f"Full kriyah - {total_verses} verses (parsha={parsha_verses})", 
           "Reader",
           "Aliyah",
           "Hebrew Name(s)",
           "Notes"
       ]],
       range_name='A14:F14'
   )
   time.sleep(5)
   
   colors = [
       {'red': 1.0, 'green': 1.0, 'blue': 0.8},
       {'red': 1.0, 'green': 0.8, 'blue': 1.0},
       {'red': 0.8, 'green': 1.0, 'blue': 1.0},
   ]
   
   row = 15
   color_index = 0
   
   for key in sorted(fullkriyah.keys()):
       if key == 'M':
           continue
           
       aliyah = fullkriyah[key]
       aliyah_num = int(key) if key.isdigit() else key
       display_num = int_to_roman(int(aliyah_num)) if isinstance(aliyah_num, int) else aliyah_num
       verse_info = format_verse_range(aliyah)
       
       worksheet.update(
           values=[[
               display_num,
               verse_info,
               "",
               "",
               "",
               ""
           ]],
           range_name=f'A{row}:F{row}'
       )
       
       worksheet.format(f'A{row}:C{row}', {
           'backgroundColor': colors[color_index]
       })
       worksheet.format(f'A{row}', {
           'horizontalAlignment': 'CENTER'
       })
       
       time.sleep(5)
       color_index = (color_index + 1) % 3
       row += 1
   
   if 'M' in fullkriyah:
       maftir = fullkriyah['M']
       verse_info = format_verse_range(maftir)
       
       worksheet.update(
           values=[[
               "Maf",
               verse_info,
               "",
               "",
               "",
               ""
           ]],
           range_name=f'A{row}:F{row}'
       )
       
       worksheet.format(f'A{row}:C{row}', {
           'backgroundColor': colors[color_index]
       })
       worksheet.format(f'A{row}', {
           'horizontalAlignment': 'CENTER'
       })
       
       time.sleep(5)
       color_index = (color_index + 1) % 3
       row += 1
   
   if parsha_data:
       if page_numbers and pd.notna(page_numbers.get('Haftarah verses')):
           verse_info = page_numbers['Haftarah verses']
       elif 'haft' in parsha_data:
           haftarah_parts = parsha_data['haft']
           if isinstance(haftarah_parts, list):
               verse_parts = []
               total_verses = 0
               for part in haftarah_parts:
                   verse_parts.append(f"{part['b']}-{part['e']}")
                   total_verses += part['v']
               book = haftarah_parts[0]['k']
               verse_info = f"{book} {', '.join(verse_parts)} ({total_verses})"
           else:
               part = haftarah_parts
               verse_info = f"{part['k']} {part['b']}-{part['e']} ({part['v']})"
       
       worksheet.update(
           values=[[
               "Haf",
               verse_info,
               "",
               "",
               "",
               ""
           ]],
           range_name=f'A{row}:F{row}'
       )
       
       worksheet.format(f'A{row}:C{row}', {
           'backgroundColor': colors[color_index]
       })
       worksheet.format(f'A{row}', {
           'horizontalAlignment': 'CENTER'
       })
       
       time.sleep(5)
       
def write_footer(worksheet, page_numbers=None, verbose=False):
    """Write footer section (rows 24-34)."""
    if page_numbers:
        torah_page = f"Torah page {str(int(page_numbers['Torah Page']))}"\
            if pd.notna(page_numbers.get('Torah Page')) else "Torah page"
        haftarah_page = f"Haftarah page {str(int(page_numbers['Haftarah Page']))}"\
            if pd.notna(page_numbers.get('Haftarah Page')) else "Haftarah page"
    else:
        torah_page = "Torah page"
        haftarah_page = "Haftarah page"

    footer_data = [
        ["", "", "", "", "", ""],  # Row 24 (blank)
        ["", "Honors", "", "Etz Hayyim", "", ""],  # Row 25
        ["P'ticha 1", "", "", torah_page, "", ""],  # Row 26
        ["P'ticha 2", "", "", haftarah_page, "", ""],  # Row 27
        ["Hagbah", "", "", "", "", ""],  # Row 28
        ["G'lilah", "", "", "", "", ""],  # Row 29
        ["Prayer for Country", "", "", "", "", ""],  # Row 30
        ["Prayer for Israel", "", "", "", "", ""],  # Row 31
        ["Prayer for Peace", "", "", "", "", ""],  # Row 32
        ["Anim Zmerot", "", "", "", "", ""],  # Row 33
        ["Adon Olam", "", "", "", "", ""]  # Row 34
    ]

    range_name = 'A24:F34'
    worksheet.batch_update([{
        'range': range_name,
        'values': footer_data
    }])
    time.sleep(5)

    gray_format = {
        'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}
    }
    
    for cell in ['A25', 'B25', 'D25']:
        worksheet.format(cell, gray_format)
        time.sleep(5)
        
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def get_leyning(start_date, end_date, verbose=False):
    """
    Fetch leyning data from HebCal API with retry logic
    """
    url = f"https://www.hebcal.com/leyning?cfg=json&start={start_date}&end={end_date}"
    
    if verbose:
        print(f"Fetching data from {url}")
    
    response = requests.get(url)
    response.raise_for_status()
    
    return response.json()

def set_global_format(worksheet, verbose=False):
    """Set global formatting rules for the worksheet."""
    if verbose:
        print("Applying global formatting...")
    
    # Set default font and size for the entire sheet
    worksheet.format(
        'A1:F1000',  # Apply to a large range to cover all potential cells
        {
            'textFormat': {
                'fontFamily': 'Arial',
                'fontSize': 11
            },
            'wrapStrategy': 'OVERFLOW_CELL'  # Text will overflow into adjacent cells
        }
    )
    time.sleep(5)  # Respect rate limits


def write_minyan(worksheet, parsha_data, verbose=False):
    """Update worksheet with weekday Torah readings and special days."""
    if verbose:
        print("Updating Minyan readings tab...")

    # Clear existing content and set formatting
    worksheet.clear()
    time.sleep(5)
    set_global_format(worksheet, verbose)
    set_column_widths(worksheet, verbose)

    # Define background colors
    GRAY_BG = {'red': 0.9, 'green': 0.9, 'blue': 0.9}  # Regular headers
    RED_BG = {'red': 1.0, 'green': 0.8, 'blue': 0.8}   # Fast days
    GREEN_BG = {'red': 0.8, 'green': 1.0, 'blue': 0.8} # Rosh Chodesh and Chol Ha-moed

    # Collect all relevant readings in chronological order
    readings = []
    for item in parsha_data['items']:
        # Skip regular parsha readings that aren't weekday readings
        if not ('weekday' in item or 'fullkriyah' in item and is_special_day(item['name']['en'])):
            continue
            
        reading_type = get_reading_type(item['name']['en'])
        readings.append({
            'readings': item.get('weekday', item.get('fullkriyah', {})),
            'parsha_name': item['name']['en'],
            'date': item['date'],
            'hdate': item['hdate'],
            'type': reading_type
        })
    
    # Sort by date
    readings.sort(key=lambda x: x['date'])

    if not readings:
        if verbose:
            print("No readings found")
        return

    # Prepare all rows for batch update
    all_rows = []
    header_rows = []  # Keep track of which rows are headers
    
    # Write each reading
    for reading_info in readings:
        date_obj = datetime.strptime(reading_info['date'], '%Y-%m-%d')
        
        # Format dates
        secular_date = date_obj.strftime('%b %d')
        hebrew_date = ' '.join(reading_info['hdate'].split()[:-1])  # Remove year
        
        # Record this as a header row
        header_rows.append(len(all_rows))
        
        # Add header row
        all_rows.append([
            secular_date,
            hebrew_date, 
            reading_info['parsha_name'],
            date_obj.strftime('%A')
        ])
        
        # Add aliyah readings
        for aliyah_num, reading in reading_info['readings'].items():
            if aliyah_num != 'M':  # Skip Maftir for weekday readings
                roman_num = int_to_roman(int(aliyah_num)) if aliyah_num.isdigit() else aliyah_num
                verse_info = format_verse_range(reading)
                all_rows.append([roman_num, verse_info, '', ''])
        
        # Add blank row between sections
        all_rows.append(['', '', '', ''])

    # Write all data at once
    range_name = f'A1:D{len(all_rows)}'
    worksheet.batch_update([{
        'range': range_name,
        'values': all_rows
    }])
    time.sleep(5)

    # Apply formatting
    for i, (reading_info, header_row) in enumerate(zip(readings, header_rows)):
        # Determine background color based on reading type
        if reading_info['type'] == 'fast_day':
            bg_color = RED_BG
        elif reading_info['type'] in ['rosh_chodesh', 'chol_hamoed']:
            bg_color = GREEN_BG
        else:
            bg_color = GRAY_BG

        # Format header - add 1 because sheet rows are 1-based
        worksheet.format(f'A{header_row + 1}:D{header_row + 1}', {
            'backgroundColor': bg_color,
            'textFormat': {'bold': True},
            'horizontalAlignment': 'CENTER'
        })
        time.sleep(1)
        
        # Center align aliyah numbers for this section
        start_row = header_row + 2  # First aliyah row
        
        # Find end of current section by looking for the next blank row
        end_row = start_row
        while end_row < len(all_rows) and (end_row == start_row or any(all_rows[end_row-1])):
            if all_rows[end_row-1][0]:  # If there's content in column A
                worksheet.format(f'A{end_row}', {
                    'horizontalAlignment': 'CENTER'
                })
                time.sleep(1)
            end_row += 1

    if verbose:
        print("Minyan readings tab updated successfully")

def write_to_sheets(data, sheet_name, user_email, test_mode=False, page_numbers=None, verbose=False):
   """
   Write leyning data to Google Sheets, with separate tabs for each parsha
   """
   try:
       scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
       credentials = Credentials.from_service_account_file('credentials.json', scopes=scopes)
       gc = gspread.authorize(credentials)

       if verbose:
           print(f"Connecting to Google Sheets: {sheet_name}")

       try:
           spreadsheet = gc.open(sheet_name)
           if verbose:
               print("Found existing spreadsheet")
       except gspread.SpreadsheetNotFound:
           spreadsheet = gc.create(sheet_name)
           if verbose:
               print("Created new spreadsheet")
       
       spreadsheet.share(None, perm_type='anyone', role='writer', with_link=True)
       if user_email:
           spreadsheet.share(user_email, perm_type='user', role='writer')
           if verbose:
               print(f"Shared spreadsheet with {user_email}")
       time.sleep(5)

       parsha_data = defaultdict(list)
       for item in data['items']:
           parsha_name = item['name']['en']
           if not is_special_day(parsha_name):
               parsha_data[parsha_name].append(item)

       if test_mode and parsha_data:
           first_parsha = next(iter(parsha_data))
           parsha_data = {first_parsha: parsha_data[first_parsha]}
           if verbose:
               print(f"Test mode: Processing only parsha {first_parsha}")

       worksheets = spreadsheet.worksheets()
       
       first_sheet = worksheets[0]
       if first_sheet.title != "Minyan":
           first_sheet.update_title("Minyan")
           time.sleep(5)
       first_sheet.clear()
       time.sleep(5)
       
       if len(worksheets) > 1:
           if verbose:
               print("Removing old worksheets...")
           for worksheet in tqdm(worksheets[1:], disable=not verbose):
               try:
                   spreadsheet.del_worksheet(worksheet)
                   time.sleep(5)
               except Exception as e:
                   if verbose:
                       print(f"Error deleting worksheet: {e}")
                   continue
       
       write_minyan(first_sheet, data, verbose)
       
       if verbose:
           print("Processing parshas...")
       
       for parsha_name, items in tqdm(parsha_data.items(), disable=not verbose):
           if verbose:
               print(f"\nProcessing {parsha_name}")

           parsha_instance = next(
               (item for item in items if 'fullkriyah' in item),
               items[0]
           )

           try:
               worksheet = spreadsheet.add_worksheet(parsha_name, 1000, 26)
           except gspread.exceptions.APIError as e:
               if "already exists" in str(e):
                   if verbose:
                       print(f"Sheet {parsha_name} already exists, trying to delete it first")
                   try:
                       old_sheet = spreadsheet.worksheet(parsha_name)
                       spreadsheet.del_worksheet(old_sheet)
                       time.sleep(5)
                       worksheet = spreadsheet.add_worksheet(parsha_name, 1000, 26)
                   except Exception as inner_e:
                       print(f"Error handling duplicate sheet: {inner_e}")
                       continue
               else:
                   raise e

           time.sleep(5)

           parsha_pages = page_numbers.get(parsha_name) if page_numbers else None
           
           set_global_format(worksheet, verbose)
           set_column_widths(worksheet, verbose)
           write_header(worksheet, parsha_instance)
           write_aliyot(worksheet, parsha_instance.get('fullkriyah', {}), parsha_instance, page_numbers=parsha_pages)
           write_footer(worksheet, page_numbers=parsha_pages)

       worksheets = spreadsheet.worksheets()
       if worksheets[0].title != "Minyan":
           spreadsheet.reorder_worksheets([first_sheet] + [ws for ws in worksheets if ws.title != "Minyan"])
           time.sleep(5)

       if verbose:
           print(f"\nSuccessfully wrote data to {sheet_name}")
           print(f"Spreadsheet URL: {spreadsheet.url}")

       return spreadsheet.url

   except Exception as e:
       print(f"Error writing to Google Sheets: {e}", file=sys.stderr)
       raise

def main():
   parser = argparse.ArgumentParser(description='Fetch Torah reading information from HebCal API')
   parser.add_argument('start_date', help='Start date in YYYY-MM-DD format')
   parser.add_argument('end_date', help='End date in YYYY-MM-DD format')
   parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose output')
   parser.add_argument('-s', '--sheet', help='Google Sheet name (if not provided, will only print JSON)')
   parser.add_argument('-e', '--email', help='Email address to share the sheet with')
   parser.add_argument('-t', '--test', action='store_true', help='Test mode - only process first parsha')
   parser.add_argument('--pages', help='CSV file with page numbers')
   
   args = parser.parse_args()
   
   try:
       datetime.strptime(args.start_date, '%Y-%m-%d')
       datetime.strptime(args.end_date, '%Y-%m-%d')
   except ValueError:
       print("Error: Dates must be in YYYY-MM-DD format", file=sys.stderr)
       sys.exit(1)
   
   # Get the leyning data
   data = get_leyning(args.start_date, args.end_date, verbose=args.verbose)
   
   # Load page numbers if CSV provided
   page_numbers = None
   if args.pages:
       page_numbers = load_page_numbers(args.pages)
       
   # Only print JSON output if verbose is on
   if args.verbose:
       print(json.dumps(data, indent=2, ensure_ascii=False))

   # Write to Google Sheets if requested
   if args.sheet:
       if not args.email:
           print("Error: --email is required when using --sheet", file=sys.stderr)
           sys.exit(1)
       sheet_url = write_to_sheets(data, args.sheet, args.email, 
                                 test_mode=args.test, 
                                 page_numbers=page_numbers,
                                 verbose=args.verbose)
       print(f"\nData written to Google Sheet: {sheet_url}")

if __name__ == "__main__":
   main()

