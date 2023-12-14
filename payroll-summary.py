import gspread
import configparser
import pandas as pd

from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import get_effective_format

from pathlib import Path

main_spreadsheet_id="1sNzFxpxb3XxLRP2xAENd1LC0k6V113ZUTNH-vs5OMpA"
second_spreadsheet_id="1DTEvWiKttZY0G7jL1SJiNIP1SWkm3w73XhdmTw-ph6U"
summary_spreadsheet_id="14LNQnrTL6P5jBvnb7cvBVYEu3KNweIyYpAGUnreN0E4"

# Set up gspread
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
scopes_string = ' '.join(scope)
creds = ServiceAccountCredentials.from_json_keyfile_name('config/lla-payroll-c3b730c6f614.json', scopes_string)
client = gspread.authorize(creds)
main_spreadsheet = client.open_by_key(main_spreadsheet_id)
second_spreadsheet = client.open_by_key(second_spreadsheet_id)
summary_spreadsheet = client.open_by_key(summary_spreadsheet_id)

execute_on_date = "2023-12-10"

# config = configparser.ConfigParser()
# config.read('config/config.ini')
# x_api_key = config['API']['X-API-Key']


def get_payroll_period(spreadsheet, execute_on_date):
    # Access the "Payroll" sheet
    payroll_sheet = spreadsheet.worksheet("Payroll")
    
    # Fetch all values from the sheet
    all_values = payroll_sheet.get_all_values()
    
    # Find the headers
    headers = all_values[0]
    

    
    # Search for the row with the matching "Execute On" date
    for row in all_values[1:]:
        if row[0] == execute_on_date:
            from_date = row[headers.index("FromDate")]
            to_date = row[headers.index("ToDate")]
            week_no = row[headers.index("WeekNo")]
            sheet_name = row[headers.index("SheetName")]
            
            return from_date, to_date, week_no, sheet_name
            
    # If the "Execute On" date is not found
    return None

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

def get_payroll_settings(spreadsheet, payroll):

    from_date, to_date, week_no, sheet_name = payroll
     # Read the "Setting" worksheet to get the column mappings
    print("Read settings")
    setting_worksheet = spreadsheet.worksheet("Settings")
    metrics = setting_worksheet.col_values(1)  # Metrics are in the first column
    if week_no == "Week1":
        columns = setting_worksheet.col_values(2)  # Week1 columns are in the second column
    elif week_no == "Week2":
        columns = setting_worksheet.col_values(3)  # Week2 columns are in the third column
        
    # Extend for more weeks as needed
    else:
        return f"Week {week_no} settings not found."
    
    # Create a dictionary mapping metrics to columns
    metric_to_column = dict(zip(metrics, columns))
    # Convert column letters to numbers
    metric_to_column = {metric: column_letter_to_number(column) for metric, column in metric_to_column.items()}
    return metric_to_column

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
    payroll = get_payroll_period(from_spreadsheet, execute_on_date)
    metric = get_payroll_settings(from_spreadsheet, payroll)
    from_date, to_date, week_no, sheet_name = payroll
    payroll_sheet = from_spreadsheet.worksheet(sheet_name)
    
    payroll_values = payroll_sheet.get_all_values()
    cells_to_update = []

    for i, row in ((index, shop) for index, shop in enumerate(summary_shops) if shop[0] != ''):
        shop_id = row[0]
        found_shop_id = False
        total_overtime = total_payroll = sales = 0

        for payroll_row in payroll_values:
            if found_shop_id and payroll_row[metric["Overtime"]-2] == "Personnel over":  
                total_overtime = payroll_row[metric["Overtime"]-1]
                total_payroll = payroll_row[metric["TotalPayroll"]-1]  
                
                break
            if payroll_row[0] == shop_id:  # when shop_id is found
                found_shop_id = True
                sales = payroll_row[metric["Sales"]-1]

        if found_shop_id:
            # Clean and convert the values
            sales = clean_currency_value(sales)
            total_payroll = clean_currency_value(total_payroll)
            total_overtime = clean_currency_value(total_overtime)
            # Creating the Cell objects and adding them to the cells_to_update list
            cells_to_update.append(gspread.Cell(row=i+2, col=3, value=sales))
            cells_to_update.append(gspread.Cell(row=i+2, col=4, value=total_payroll))
            cells_to_update.append(gspread.Cell(row=i+2, col=5, value=total_overtime))

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

    # Combine the header (without the last column), sorted data, and footer (without the last column)
    final_data = [headers[:-1]] + sorted_data + [footer[:-1]]

    # Determine the last column based on the number of columns in the data (after excluding the last column)
    last_col_letter = chr(64 + len(final_data[0]))  # Convert number to column letter (e.g., 1 -> A, 2 -> B, ...)
    range_str = f"A1:{last_col_letter}{len(final_data)}"
    
    # Update the worksheet with the combined data
    summary_worksheet.update(range_str, final_data)



        
            





summary_worksheet = get_or_create_sheet(summary_spreadsheet, execute_on_date)
fill_summary(main_spreadsheet,summary_worksheet)
fill_summary(second_spreadsheet,summary_worksheet)
order_summary(summary_worksheet)

