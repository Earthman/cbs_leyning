import requests
import json
from datetime import datetime, timedelta
import sys
import argparse
import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as OAuth2Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from collections import defaultdict
import time
from tenacity import retry, stop_after_attempt, wait_exponential
from tqdm import tqdm
import pandas as pd
import os
import pickle
import warnings

# Suppress gspread deprecation warnings for worksheet.update()
# We're using the correct API for v6+, but the library still shows warnings
warnings.filterwarnings('ignore', message='.*Method signature.*range_name.*values.*', category=DeprecationWarning)

sleepy_time = 2  # Default for service account
oauth_sleepy_time = 3  # Longer delay for OAuth to avoid rate limits

# Optimized sleep times for template mode
template_sleep_time = 1.5  # For header/footer with templates (fewer API calls)
aliyot_sleep_time = 3  # For aliyot section (many API calls, increased to avoid rate limit)

def smart_sleep(context='default', use_template=False):
    """
    Sleep for appropriate duration based on context and template usage.

    Args:
        context: 'header', 'footer', 'aliyot', or 'default'
        use_template: Whether template mode is active
    """
    if use_template and context in ['header', 'footer']:
        time.sleep(template_sleep_time)
    elif context == 'aliyot':
        time.sleep(aliyot_sleep_time)
    else:
        time.sleep(sleepy_time)

def find_marker_in_column(data, column_index, marker_text):
    """
    Search for marker text in a specific column.

    Args:
        data: List of rows (each row is a list of cell values)
        column_index: 0-based column index to search
        marker_text: Text to search for

    Returns:
        Row index (0-based) or None if not found
    """
    for row_idx, row in enumerate(data):
        if column_index < len(row) and row[column_index] == marker_text:
            return row_idx
    return None

def get_oauth2_credentials():
    """Get OAuth2 credentials from saved token or by running auth flow."""
    print("Getting OAuth2 credentials...", file=sys.stderr)
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.file'
    ]
    
    creds = None
    token_file = 'token.pickle'
    
    # Load existing token if it exists
    if os.path.exists(token_file):
        print(f"Loading token from {token_file}...", file=sys.stderr)
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
        print(f"Token loaded. Valid: {creds.valid}, Expired: {creds.expired}", file=sys.stderr)
    
    # If there are no (valid) credentials available, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired token...", file=sys.stderr)
            creds.refresh(Request())
            print("Token refreshed.", file=sys.stderr)
        else:
            if not os.path.exists('oauth_credentials.json'):
                print("Error: oauth_credentials.json not found!", file=sys.stderr)
                print("Please run 'python setup_oauth.py' and 'python authenticate_oauth.py' first.", file=sys.stderr)
                sys.exit(1)
            
            print("Starting OAuth flow...", file=sys.stderr)
            flow = InstalledAppFlow.from_client_secrets_file(
                'oauth_credentials.json', SCOPES)
            creds = flow.run_local_server(port=8080, open_browser=True)
        
        # Save the credentials for next run
        print("Saving refreshed token...", file=sys.stderr)
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)
    
    print("OAuth2 credentials ready.", file=sys.stderr)
    return creds

def load_template(template_name, gc, verbose=False):
    """
    Load template spreadsheet and detect dynamic dimensions.

    Args:
        template_name: Name of the template spreadsheet
        gc: Authorized gspread client
        verbose: Print detailed information

    Returns:
        Dictionary containing:
            'header_data': List of rows from Header tab
            'header_length': Number of rows in header
            'scroll_row': Row index (0-based) containing "Torah(s) Scroll"
            'column_header_row': Row index (0-based) with "Reader", "Aliyah", etc.
            'footer_data': List of rows from Footer tab
            'footer_length': Number of rows in footer

    Raises:
        Exception if template not found or markers missing
    """
    try:
        # Open template spreadsheet by name or ID
        if verbose:
            print(f"Loading template: {template_name}", file=sys.stderr)

        # Check if template_name looks like a spreadsheet ID (long alphanumeric string)
        # IDs are typically 44 characters and contain letters, numbers, hyphens, underscores
        if len(template_name) > 30 and all(c.isalnum() or c in '-_' for c in template_name):
            # Treat as ID
            spreadsheet = gc.open_by_key(template_name)
            if verbose:
                print(f"Opened template by ID", file=sys.stderr)
        else:
            # Treat as name
            spreadsheet = gc.open(template_name)
            if verbose:
                print(f"Opened template by name", file=sys.stderr)

        # Load Header tab
        header_sheet = spreadsheet.worksheet('Header')
        header_data = header_sheet.get_all_values()

        # Find Torah Scroll row (search column A for "Torah(s) Scroll")
        scroll_row = find_marker_in_column(header_data, 0, "Torah(s) Scroll")
        if scroll_row is None:
            raise ValueError("Template Header must contain 'Torah(s) Scroll' in column A")

        # Find column header row (search column C for "Reader")
        column_header_row = find_marker_in_column(header_data, 2, "Reader")
        if column_header_row is None:
            raise ValueError("Template Header must contain 'Reader' in column C")

        # Validate that column_header_row also has other required markers
        if column_header_row < len(header_data):
            row = header_data[column_header_row]
            required_markers = {3: "Aliyah", 4: "Hebrew Name(s)", 5: "Notes"}
            for col_idx, marker in required_markers.items():
                if col_idx >= len(row) or row[col_idx] != marker:
                    raise ValueError(f"Template Header row {column_header_row+1} must contain '{marker}' in column {chr(65+col_idx)}")

        header_length = column_header_row + 1  # +1 because we want row count, not index

        # Load Footer tab
        footer_sheet = spreadsheet.worksheet('Footer')
        footer_data = footer_sheet.get_all_values()

        # Count non-empty rows in footer
        footer_length = sum(1 for row in footer_data if any(cell.strip() for cell in row))

        if verbose:
            print(f"Template loaded: {template_name}", file=sys.stderr)
            print(f"  Header length: {header_length} rows", file=sys.stderr)
            print(f"  Torah Scroll row: {scroll_row + 1} (spreadsheet row)", file=sys.stderr)
            print(f"  Column header row: {column_header_row + 1} (spreadsheet row)", file=sys.stderr)
            print(f"  Footer length: {footer_length} rows", file=sys.stderr)

        return {
            'header_data': header_data,
            'header_length': header_length,
            'scroll_row': scroll_row,
            'column_header_row': column_header_row,
            'footer_data': footer_data,
            'footer_length': footer_length
        }

    except gspread.SpreadsheetNotFound:
        raise Exception(f"Template spreadsheet '{template_name}' not found")
    except gspread.WorksheetNotFound as e:
        raise Exception(f"Template is missing required tab: {e}")

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

def write_header(worksheet, parsha_data, scroll_name="Gunther", template_data=None, verbose=False):
    """
    Write header section with dynamic field population.

    Args:
        worksheet: gspread worksheet object
        parsha_data: Dictionary containing parsha information
        scroll_name: Name of Torah scroll
        template_data: Optional template data dictionary
        verbose: Print detailed information

    Returns:
        Next available row number for aliyot (1-based for spreadsheet)
    """
    use_template = template_data is not None

    # Parse and format dates
    full_date = datetime.strptime(parsha_data['date'], '%Y-%m-%d')
    gregorian_date = full_date.strftime('%B %-d')
    previous_date = (full_date - timedelta(days=1)).strftime('%B %-d')

    # Parse Hebrew date to get just month and day
    hebrew_date_parts = parsha_data['hdate'].split()
    hebrew_date = f"{hebrew_date_parts[1]} {hebrew_date_parts[0]}"

    # Calculate verse counts
    total_verses = 0
    parsha_verses = 0
    if 'fullkriyah' in parsha_data:
        for key, aliyah in parsha_data['fullkriyah'].items():
            verses = aliyah.get('v', 0)
            total_verses += verses
            if key != 'M':
                parsha_verses += verses

    # Check for special Shabbat
    special_shabbat = None
    if isinstance(parsha_data.get('reason'), dict):
        special_shabbat = parsha_data['reason'].get('haftara')
    if not special_shabbat and 'haft' in parsha_data:
        haft = parsha_data['haft']
        if isinstance(haft, dict):
            special_shabbat = haft.get('reason')
        elif isinstance(haft, list):
            for h in haft:
                if isinstance(h, dict) and 'reason' in h:
                    special_shabbat = h['reason']
                    break

    # Ensure scroll_name is a string
    scroll_name_str = str(scroll_name) if scroll_name is not None else "Gunther"

    if template_data:
        # TEMPLATE MODE: Copy template and update dynamic fields
        header_data = template_data['header_data']
        header_length = template_data['header_length']
        scroll_row = template_data['scroll_row']
        column_header_row = template_data['column_header_row']

        # Copy entire header in one batch operation
        worksheet.batch_update([{
            'range': f'A1:F{header_length}',
            'values': header_data[:header_length]
        }])
        smart_sleep('header', use_template=True)

        # Batch update dynamic data cells
        updates = [
            {'range': 'B1', 'values': [[parsha_data['name']['en']]]},
            {'range': 'D1', 'values': [[gregorian_date]]},
            {'range': 'E1', 'values': [[hebrew_date]]},
            {'range': 'B3', 'values': [[f"Kabbalat Shabbat {previous_date}"]]},
            {'range': f'B{scroll_row + 1}', 'values': [[scroll_name_str]]},
            {'range': f'B{column_header_row + 1}',
             'values': [[f"Full kriyah - {total_verses} verses (parsha={parsha_verses})"]]},
        ]

        if special_shabbat:
            updates.append({'range': 'D2', 'values': [[special_shabbat]]})

        worksheet.batch_update(updates)
        smart_sleep('header', use_template=True)

        # Update formulas
        worksheet.update('C6',
                        [['=if(ISNUMBER(SEARCH("Richman",$A$2)), "RDR default", "RAR default")']],
                        value_input_option='USER_ENTERED')
        smart_sleep('header', use_template=True)

        # Virtual Tikkun link
        vt_link = f'=hyperlink("https://myvirtualtikkun.com/?shul=cbssf&view=both&scroll={scroll_name_str}&parsha={parsha_data["name"]["en"]}", "Virtual Tikkun")'
        worksheet.update(f'C{scroll_row + 1}', [[vt_link]], value_input_option='USER_ENTERED')
        smart_sleep('header', use_template=True)

        # Apply formatting (optional - may already be in template)
        worksheet.format('A1:F1', {'textFormat': {'fontSize': 24}})
        smart_sleep('header', use_template=True)

        worksheet.format('A2:F2', {'textFormat': {'fontSize': 14}})
        smart_sleep('header', use_template=True)

        return header_length + 1  # Next row after header

    else:
        # FALLBACK MODE: Use existing logic
        virtual_tikkun_link = f'=hyperlink("https://myvirtualtikkun.com/?shul=cbssf&view=both&scroll={scroll_name_str}&parsha={parsha_data["name"]["en"]}", "Virtual Tikkun")'

        header_data = [
            ["", parsha_data['name']['en'], "", gregorian_date, hebrew_date],
            ["Rabbi Amanda Russell", "", "", special_shabbat if special_shabbat else "", ""],
            ["Service leaders", f"Kabbalat Shabbat {previous_date}", "", "", ""],
            ["", "P'sukei D'zimrah", "", "", ""],
            ["", "Shacharit", "", "", ""],
            ["", "Musaf", "", "", ""],
            ["", "Torah Service", "", "", ""],
            ["", "Gabbai", "Sam (default)", "", ""],
            ["", "Distribute honors", "Todd (default)", "", ""],
            ["", "Read announcements", "Jerilyn (default)", "", ""],
            ["Board hosts", "", "", "", ""],
            ["", "", "", "", ""],
            ["Torah(s) Scroll", scroll_name_str, "", "", ""],
            ["", f"Full kriyah - {total_verses} verses (parsha={parsha_verses})", "Reader", "Aliyah", "Hebrew Name(s)", "Notes"]
        ]

        worksheet.batch_update([{
            'range': 'A1:F14',
            'values': header_data
        }])
        time.sleep(sleepy_time)

        formula = '=if(ISNUMBER(SEARCH("Richman",$A$2)), "RDR default", "RAR default")'
        worksheet.update_acell('C6', formula)
        time.sleep(sleepy_time)

        if virtual_tikkun_link.startswith('=hyperlink'):
            worksheet.update('C13', [[virtual_tikkun_link]], value_input_option='USER_ENTERED')
            time.sleep(sleepy_time)

        formats = [
            {'range': 'A1:F1', 'format': {'textFormat': {'fontSize': 24}}},
            {'range': 'A2:F2', 'format': {'textFormat': {'fontSize': 14}}},
            {'range': 'A3', 'format': {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}},
            {'range': 'A11', 'format': {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}},
            {'range': 'A13:B13', 'format': {'backgroundColor': {'red': 1.0, 'green': 0.8, 'blue': 0.6}}},
            {'range': 'A14:F14', 'format': {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}}
        ]

        for format_spec in formats:
            worksheet.format(format_spec['range'], format_spec['format'])
            time.sleep(sleepy_time)

        return 15  # Hardcoded for fallback mode

def write_aliyot(worksheet, fullkriyah, parsha_data, start_row, page_numbers=None, scroll_name="Gunther"):
   """
   Write aliyot section starting at specified row.

   Args:
       worksheet: gspread worksheet object
       fullkriyah: Dictionary of aliyot data
       parsha_data: Dictionary containing parsha information
       start_row: Row number to begin writing (1-based)
       page_numbers: Optional page number data
       scroll_name: Name of Torah scroll

   Returns:
       Next available row number after aliyot (1-based)
   """
   if not fullkriyah:
       return start_row

   colors = [
       {'red': 1.0, 'green': 1.0, 'blue': 0.8},
       {'red': 1.0, 'green': 0.8, 'blue': 1.0},
       {'red': 0.8, 'green': 1.0, 'blue': 1.0},
   ]

   row = start_row
   color_index = 0
   
   # Get parsha name for Virtual Tikkun links
   parsha_name = parsha_data['name']['en'] if parsha_data else ""
   # Ensure scroll_name is a string
   scroll_name_str = str(scroll_name) if scroll_name is not None else "Gunther"
   
   for key in sorted(fullkriyah.keys()):
       if key == 'M':
           continue
           
       aliyah = fullkriyah[key]
       aliyah_num = int(key) if key.isdigit() else key
       display_num = int_to_roman(int(aliyah_num)) if isinstance(aliyah_num, int) else aliyah_num
       verse_info = format_verse_range(aliyah)
       
       # Create Virtual Tikkun link for this aliyah
       vt_aliyah = f"A{aliyah_num}"
       vt_link = f'=hyperlink("https://myvirtualtikkun.com/?shul=cbssf&scroll={scroll_name_str}&parsha={parsha_name}&aliyah={vt_aliyah}", "{display_num}")'
       
       # First write the plain data (without the hyperlink in column A)
       worksheet.update(
           f'A{row}:F{row}',
           [[
               display_num,  # This will be replaced by the hyperlink formula
               verse_info,
               "",
               "",
               "",
               ""
           ]]
       )
       
       # Then update column A with the hyperlink formula
       worksheet.update(f'A{row}', [[vt_link]], value_input_option='USER_ENTERED')
       
       worksheet.format(f'A{row}:C{row}', {
           'backgroundColor': colors[color_index]
       })
       worksheet.format(f'A{row}', {
           'horizontalAlignment': 'CENTER'
       })

       smart_sleep('aliyot')
       color_index = (color_index + 1) % 3
       row += 1

   if 'M' in fullkriyah:
       maftir = fullkriyah['M']
       verse_info = format_verse_range(maftir)

       # Create Virtual Tikkun link for Maftir
       vt_link = f'=hyperlink("https://myvirtualtikkun.com/?shul=cbssf&scroll={scroll_name_str}&parsha={parsha_name}&aliyah=M", "Maf")'

       # First write the plain data
       worksheet.update(
           f'A{row}:F{row}',
           [[
               "Maf",  # This will be replaced by the hyperlink formula
               verse_info,
               "",
               "",
               "",
               ""
           ]]
       )

       # Then update column A with the hyperlink formula
       worksheet.update(f'A{row}', [[vt_link]], value_input_option='USER_ENTERED')

       worksheet.format(f'A{row}:C{row}', {
           'backgroundColor': colors[color_index]
       })
       worksheet.format(f'A{row}', {
           'horizontalAlignment': 'CENTER'
       })

       smart_sleep('aliyot')
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
           f'A{row}:F{row}',
           [[
               "Haf",
               verse_info,
               "",
               "",
               "",
               ""
           ]]
       )

       worksheet.format(f'A{row}:C{row}', {
           'backgroundColor': colors[color_index]
       })
       worksheet.format(f'A{row}', {
           'horizontalAlignment': 'CENTER'
       })

       smart_sleep('aliyot')
       row += 1

   return row  # Return next available row

def write_footer(worksheet, start_row, page_numbers=None, template_data=None, verbose=False):
    """
    Write footer section starting at specified row.

    Args:
        worksheet: gspread worksheet object
        start_row: Row number to begin writing footer (1-based)
        page_numbers: Optional page number data
        template_data: Optional template data dictionary
        verbose: Print detailed information
    """
    use_template = template_data is not None

    # Prepare page number text
    if page_numbers:
        torah_page = f"Torah page {str(int(page_numbers['Torah Page']))}"\
            if pd.notna(page_numbers.get('Torah Page')) else "Torah page"
        haftarah_page = f"Haftarah page {str(int(page_numbers['Haftarah Page']))}"\
            if pd.notna(page_numbers.get('Haftarah Page')) else "Haftarah page"
    else:
        torah_page = "Torah page"
        haftarah_page = "Haftarah page"

    if template_data:
        # TEMPLATE MODE: Copy template and update dynamic fields
        footer_data = template_data['footer_data']
        footer_length = template_data['footer_length']

        # Copy entire footer
        end_row = start_row + footer_length - 1
        worksheet.batch_update([{
            'range': f'A{start_row}:F{end_row}',
            'values': footer_data[:footer_length]
        }])
        smart_sleep('footer', use_template=True)

        # Update dynamic page numbers (rows 3 and 4 of footer = start_row+2, start_row+3)
        updates = [
            {'range': f'D{start_row + 2}', 'values': [[torah_page]]},
            {'range': f'D{start_row + 3}', 'values': [[haftarah_page]]},
        ]
        worksheet.batch_update(updates)
        smart_sleep('footer', use_template=True)

        # Apply formatting if needed (may already be in template)
        gray_format = {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}
        for cell in [f'A{start_row + 1}', f'B{start_row + 1}', f'D{start_row + 1}']:
            worksheet.format(cell, gray_format)
            smart_sleep('footer', use_template=True)

    else:
        # FALLBACK MODE: Use existing logic with dynamic positioning
        footer_data = [
            ["", "", "", "", "", ""],  # Blank row
            ["", "Honors", "", "Etz Hayyim", "", ""],
            ["P'ticha 1", "", "", torah_page, "", ""],
            ["P'ticha 2", "", "", haftarah_page, "", ""],
            ["Hagbah", "", "", "", "", ""],
            ["G'lilah", "", "", "", "", ""],
            ["Prayer for Country", "", "", "", "", ""],
            ["Prayer for Israel", "", "", "", "", ""],
            ["Prayer for Peace", "", "", "", "", ""],
            ["Anim Zmerot", "", "", "", "", ""],
            ["Adon Olam", "", "", "", "", ""]
        ]

        end_row = start_row + 10  # 11 rows (0-10)
        range_name = f'A{start_row}:F{end_row}'
        worksheet.batch_update([{
            'range': range_name,
            'values': footer_data
        }])
        time.sleep(sleepy_time)

        gray_format = {'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}}
        for cell in [f'A{start_row + 1}', f'B{start_row + 1}', f'D{start_row + 1}']:
            worksheet.format(cell, gray_format)
            time.sleep(sleepy_time)

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
    time.sleep(sleepy_time)  # Respect rate limits

def write_minyan(worksheet, parsha_data, verbose=False):
    """Update worksheet with weekday Torah readings and special days."""
    if verbose:
        print("Updating Minyan readings tab...")

    # Clear existing content and set formatting
    worksheet.clear()
    time.sleep(sleepy_time)
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
    time.sleep(sleepy_time)

    # Collect all format requests to batch them
    format_requests = []
    
    for i, (reading_info, header_row) in enumerate(zip(readings, header_rows)):
        # Determine background color based on reading type
        if reading_info['type'] == 'fast_day':
            bg_color = RED_BG
        elif reading_info['type'] in ['rosh_chodesh', 'chol_hamoed']:
            bg_color = GREEN_BG
        else:
            bg_color = GRAY_BG

        # Add header format request
        format_requests.append({
            'range': f'A{header_row + 1}:D{header_row + 1}',
            'format': {
                'backgroundColor': bg_color,
                'textFormat': {'bold': True},
                'horizontalAlignment': 'CENTER'
            }
        })
        
        # Center align aliyah numbers for this section
        start_row = header_row + 2  # First aliyah row
        
        # Find end of current section by looking for the next blank row
        end_row = start_row
        while end_row < len(all_rows) and (end_row == start_row or any(all_rows[end_row-1])):
            if all_rows[end_row-1][0]:  # If there's content in column A
                format_requests.append({
                    'range': f'A{end_row}',
                    'format': {'horizontalAlignment': 'CENTER'}
                })
            end_row += 1
    
    # Apply all formatting in a single batch
    if format_requests:
        worksheet.batch_format(format_requests)
        time.sleep(sleepy_time)

    if verbose:
        print("Minyan readings tab updated successfully")

def write_to_sheets(data, sheet_name, test_mode=False, page_numbers=None, scroll_name="Gunther", verbose=False, use_oauth=False, template_name=None):
   """
   Write leyning data to Google Sheets with template support.

   Args:
       data: Leyning data from HebCal API
       sheet_name: Name of output spreadsheet
       test_mode: Only process first parsha
       page_numbers: Optional page number data
       scroll_name: Name of Torah scroll
       verbose: Print detailed information
       use_oauth: Use OAuth2 instead of service account
       template_name: Optional name of template spreadsheet
   """

   try:
       # Set sleep time based on authentication method
       global sleepy_time
       if use_oauth:
           # Use OAuth2 authentication (personal Google account)
           if verbose:
               print("Using OAuth2 authentication (personal Google account)")
           credentials = get_oauth2_credentials()
           print("Authorizing with gspread...", file=sys.stderr)
           gc = gspread.authorize(credentials)
           print("Authorization complete.", file=sys.stderr)
           sleepy_time = oauth_sleepy_time  # Use shorter delays for OAuth

           if verbose:
               print("Authenticated with your personal Google account")
               print(f"Using reduced sleep time: {sleepy_time}s")
       else:
           # Use service account authentication
           scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
           credentials = Credentials.from_service_account_file('credentials.json', scopes=scopes)
           gc = gspread.authorize(credentials)

           # Test service account access
           if verbose:
               print(f"Service account email: {credentials.service_account_email}")
               print(f"Scopes: {credentials.scopes}")

       # Load template once at beginning
       template_data = None
       if template_name:
           try:
               template_data = load_template(template_name, gc, verbose)
               if verbose:
                   print(f"Template mode enabled - using optimized sleep times")
                   print(f"  Header/Footer: {template_sleep_time}s")
                   print(f"  Aliyot: {aliyot_sleep_time}s")
           except Exception as e:
               print(f"Warning: Could not load template '{template_name}': {e}", file=sys.stderr)
               print("Falling back to dynamic generation with standard sleep times", file=sys.stderr)
               template_data = None

       if verbose:
           print(f"Connecting to Google Sheets: {sheet_name}")
       
       print(f"Looking for spreadsheet: {sheet_name}...", file=sys.stderr)
       try:
           spreadsheet = gc.open(sheet_name)
           print("Found existing spreadsheet", file=sys.stderr)
           if verbose:
               print("Found existing spreadsheet")
       except gspread.SpreadsheetNotFound:
           print("Spreadsheet not found, creating new one...", file=sys.stderr)
           spreadsheet = gc.create(sheet_name)
           print("Created new spreadsheet", file=sys.stderr)
           if verbose:
               print("Created new spreadsheet")
       
       # Wait a moment for the spreadsheet to be fully accessible
       time.sleep(sleepy_time)
       
       # Make it viewable by anyone with the link
       try:
           spreadsheet.share(None, perm_type='anyone', role='reader', with_link=True)
           if verbose:
               print("Made spreadsheet viewable by anyone with link")
       except Exception as e:
           print(f"Note: Could not make spreadsheet publicly viewable: {e}")
       
       print(f"Spreadsheet URL: {spreadsheet.url}")
       print(f"Spreadsheet ID: {spreadsheet.id}")
       
       time.sleep(sleepy_time)

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
           time.sleep(sleepy_time)
       first_sheet.clear()
       time.sleep(sleepy_time)
       
       if len(worksheets) > 1:
           if verbose:
               print("Removing old worksheets...")
           for worksheet in worksheets[1:]:
               try:
                   spreadsheet.del_worksheet(worksheet)
                   time.sleep(sleepy_time)
               except Exception as e:
                   if verbose:
                       print(f"Error deleting worksheet: {e}")
                   continue
       
       write_minyan(first_sheet, data, verbose)
       
       print(f"Minyan tab complete. Processing {len(parsha_data)} parshas...", file=sys.stderr)
       if verbose:
           print("Processing parshas...")
       
       # Process each parsha
       for parsha_name, items in parsha_data.items():
           if verbose:
               print(f"\nProcessing {parsha_name}")

           parsha_instance = next(
               (item for item in items if 'fullkriyah' in item),
               items[0]
           )

           print(f"Creating worksheet for {parsha_name}...", file=sys.stderr)
           try:
               worksheet = spreadsheet.add_worksheet(parsha_name, 1000, 26)
               print(f"Worksheet {parsha_name} created.", file=sys.stderr)
           except gspread.exceptions.APIError as e:
               if "already exists" in str(e):
                   if verbose:
                       print(f"Sheet {parsha_name} already exists, trying to delete it first")
                   try:
                       old_sheet = spreadsheet.worksheet(parsha_name)
                       spreadsheet.del_worksheet(old_sheet)
                       time.sleep(sleepy_time)
                       worksheet = spreadsheet.add_worksheet(parsha_name, 1000, 26)
                   except Exception as inner_e:
                       print(f"Error handling duplicate sheet: {inner_e}")
                       continue
               else:
                   raise e

           time.sleep(sleepy_time)

           parsha_pages = page_numbers.get(parsha_name) if page_numbers else None

           set_global_format(worksheet, verbose)
           set_column_widths(worksheet, verbose)

           # Chain calls with row tracking
           next_row = write_header(worksheet, parsha_instance,
                                  scroll_name=scroll_name,
                                  template_data=template_data,
                                  verbose=verbose)

           next_row = write_aliyot(worksheet,
                                  parsha_instance.get('fullkriyah', {}),
                                  parsha_instance,
                                  start_row=next_row,
                                  page_numbers=parsha_pages,
                                  scroll_name=scroll_name)

           # Add blank row between aliyot and footer
           write_footer(worksheet,
                       start_row=next_row + 1,
                       page_numbers=parsha_pages,
                       template_data=template_data,
                       verbose=verbose)

       worksheets = spreadsheet.worksheets()
       if worksheets[0].title != "Minyan":
           spreadsheet.reorder_worksheets([first_sheet] + [ws for ws in worksheets if ws.title != "Minyan"])
           time.sleep(sleepy_time)

       if verbose:
           print(f"\nSuccessfully wrote data to {sheet_name}")
           print(f"Spreadsheet URL: {spreadsheet.url}")

       print(f"\n{'='*60}")
       print(f"SPREADSHEET CREATED: {sheet_name}")
       print(f"URL: {spreadsheet.url}")
       print(f"ID: {spreadsheet.id}")
       print(f"{'='*60}")
       print("The spreadsheet is viewable by anyone with the link.")
       print("To edit, make a copy of the spreadsheet.")
       print(f"{'='*60}")

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
   parser.add_argument('-t', '--test', action='store_true', help='Test mode - only process first parsha')
   parser.add_argument('--pages', help='CSV file with page numbers')
   parser.add_argument('--scroll', help='Name of scroll (default is Gunther )')
   parser.add_argument('--oauth', action='store_true', help='Use OAuth2 authentication (personal Google account) instead of service account')
   parser.add_argument('--template', default='14sr1RhmqfOOj9OiKhEbcLiIvgx7vDMd9RxECesUi6e8', help='Template spreadsheet name or ID (default: Leyning Sheets Template)')


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

   scroll_name = args.scroll if args.scroll else "Gunther"

   # Write to Google Sheets if requested
   if args.sheet:
       sheet_url = write_to_sheets(data, args.sheet,
                                 test_mode=args.test,
                                 page_numbers=page_numbers,
                                 scroll_name=args.scroll,
                                 verbose=args.verbose,
                                 use_oauth=args.oauth,
                                 template_name=args.template)
       print(f"\nData written to Google Sheet: {sheet_url}")

if __name__ == "__main__":
   main() 