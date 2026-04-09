import gspread
import configparser
import pandas as pd

from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import get_effective_format

from pathlib import Path

from types_payroll import DEFAULT_PAYROLL_METRICS_2026, PayrollMetrics
import util

spreadsheet_configs = {
    "PaycorPayroll": "1sbywsk3A3xdyO3280-GTd34tyHwwgzMCcWVLuJBhQoE"
    # "CCPaycorPayroll": "1u9xaf1AGFItt5ErTTw0zvseuwJUOoUwDBx9FXuqoLQM"
}


summary_spreadsheet_id="14LNQnrTL6P5jBvnb7cvBVYEu3KNweIyYpAGUnreN0E4"

# Set up gspread
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
scopes_string = ' '.join(scope)
creds = ServiceAccountCredentials.from_json_keyfile_name('config/lla-payroll-c3b730c6f614.json', scopes_string)
client = gspread.authorize(creds)
# main_spreadsheet = client.open_by_key(main_spreadsheet_id)
# second_spreadsheet = client.open_by_key(second_spreadsheet_id)
# weekly_spreadsheet = client.open_by_key(weekly_spreadsheet_id)
summary_spreadsheet = client.open_by_key(summary_spreadsheet_id)


config = configparser.ConfigParser()
config.read('config/config.ini')
# x_api_key = config['API']['X-API-Key']
payroll_period = util.get_period()


def column_letter_to_number(column_letter):
    """
    Convert Excel-style column letter to corresponding column number.
    
    Parameters:
    - column_letter: Excel-style column letter (e.g., 'A', 'Z', 'AA', 'AZ', 'BA', ...)
    
    Returns:
    - Column number as integer.
    """
    number = 0
    for i, char in enumerate(reversed(column_letter)):
        number += (ord(char) - 64) * (26 ** i)
    return number


def get_or_create_sheet(spreadsheet, sheet_name):
    
    # Check if worksheet with the name 'sheet_name' exists
    sheet_names = [sheet.title for sheet in spreadsheet.worksheets()]
    if sheet_name not in sheet_names:
        # Duplicate the "Master" worksheet and name it as 'sheet_name'
        master_sheet = spreadsheet.worksheet("Master")
        master_sheet.duplicate(new_sheet_name=sheet_name)
        
        # Re-fetch the new sheet to ensure we have the full worksheet object
        new_sheet = spreadsheet.worksheet(sheet_name)
        
        # Prepare the list of Worksheet objects in the desired order
        other_sheets = [sheet for sheet in spreadsheet.worksheets() if sheet.title != sheet_name]
        all_sheets_in_order = [new_sheet] + other_sheets
        
        # Reorder the sheets
        spreadsheet.reorder_worksheets(all_sheets_in_order)
        
        return new_sheet
    else:
        return spreadsheet.worksheet(sheet_name)

def clean_currency_value(value):
    """Clean and convert a currency-formatted string to a float."""
    return float(value.replace('$', '').replace(',', ''))


def fill_summary(from_spreadsheet, summary_worksheet):
    print("fill_summary")
    summary_shops = summary_worksheet.get_all_values()[1:]
    payroll_sheet = from_spreadsheet.worksheet(payroll_period.execute_on_date)
    metric: PayrollMetrics = DEFAULT_PAYROLL_METRICS_2026
    payroll_values = payroll_sheet.get_all_values(value_render_option='UNFORMATTED_VALUE')
    cells_to_update = []

    for i, row in ((index, shop) for index, shop in enumerate(summary_shops) if shop[0] != ''):
        shop_id = row[0]
        found_shop_id = False
        total_overtime = total_payroll = sales = 0

        for idx, payroll_row in enumerate(payroll_values):
            if found_shop_id and payroll_row[metric.overtime-2] == "Personnel over":
                total_overtime = payroll_row[metric.overtime-1]
                total_payroll = payroll_row[metric.total_payroll-1]
                break
            if payroll_row[0] == shop_id:  # when shop_id is found
                found_shop_id = True
                sales = payroll_values[idx + metric.sales][5]

        if found_shop_id:
            # Convert values to float (they come as raw numbers from UNFORMATTED_VALUE)
            sales = float(sales) if sales else 0
            total_payroll = float(total_payroll) if total_payroll else 0
            total_overtime = float(total_overtime) if total_overtime else 0
            # Creating the Cell objects and adding them to the cells_to_update list
            cells_to_update.append(gspread.Cell(row=i+2, col=3, value=str(sales)))
            cells_to_update.append(gspread.Cell(row=i+2, col=4, value=str(total_payroll)))
            cells_to_update.append(gspread.Cell(row=i+2, col=5, value=str(total_overtime)))

    # Batch update the cells
    summary_worksheet.update_cells(cells_to_update)

def order_summary(summary_worksheet):
    print("order_summary")
    
    all_values = summary_worksheet.get_all_values()
    
    # Separate out the header, footer, and the data in between
    headers = all_values[0]
    footer = all_values[-1]
    data = all_values[1:-1]

    # Exclude the last column from sorting
    data_without_last_col = [row[:-1] for row in data]

    # Sort data based on the ratio of values in columns 4 and 3 (indices 3 and 2)
    sorted_data = sorted(data_without_last_col, key=lambda x: clean_currency_value(x[3]) / clean_currency_value(x[2]) if x[2] and clean_currency_value(x[2]) != 0 else 0)

    # Initialize the final_data list with the headers
    final_data = [headers[:-1]]

    # Add each row of sorted_data to final_data, cleaning only the numeric columns (3rd and onward)
    for row in sorted_data:
        formatted_row = row[:2]  # Start with the first and second columns as text
        for cell in row[2:]:  # Clean only the numeric columns
            if cell != '':
                formatted_row.append(clean_currency_value(cell))
            else:
                formatted_row.append('')
        final_data.append(formatted_row)

    # Determine the last column based on the number of columns in the data (after excluding the last column)
    last_col_letter = chr(64 + len(final_data[0]))  # Convert number to column letter (e.g., 1 -> A, 2 -> B, ...)
    range_str = f"A1:{last_col_letter}{len(final_data)}"
    
    # Update the worksheet with the combined data
    summary_worksheet.update(range_str, final_data)


summary_worksheet = get_or_create_sheet(summary_spreadsheet, payroll_period.execute_on_date)
for name, current_spreadsheet_id in spreadsheet_configs.items():
    current_spreadsheet = client.open_by_key(current_spreadsheet_id)
    fill_summary(current_spreadsheet, summary_worksheet)

order_summary(summary_worksheet)

summary_web_app_url = config.get('API', 'apps-script-web-app-url')
util.trigger_apps_script(summary_web_app_url, "setTabColor", payroll_period.execute_on_date)
util.trigger_apps_script(summary_web_app_url, "processActiveSheet", payroll_period.execute_on_date)

