import gspread
import requests
import configparser
import pandas as pd

from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import get_effective_format

from pathlib import Path


execute_on_date = "2023-11-19"

# main
# spreadsheet_id="1sNzFxpxb3XxLRP2xAENd1LC0k6V113ZUTNH-vs5OMpA"

# second
spreadsheet_id="1DTEvWiKttZY0G7jL1SJiNIP1SWkm3w73XhdmTw-ph6U"

# test
# spreadsheet_id="1Bq46L6Bj0xAeqJ_mQm8wQ6SINzgMItgxqhcvjrKmUwg"




scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('config/lla-payroll-c3b730c6f614.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(spreadsheet_id)

config = configparser.ConfigParser()
config.read('config/config.ini')
x_api_key = config['API']['X-API-Key']


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

def update_shop_data_from_api(main_worksheet, payroll, metric):
    cells_to_update = []
    from_date, to_date, week_no, sheet_name = payroll
    print(f"FromDate: {from_date}, ToDate: {to_date}, WeekNo: {week_no}, SheetName: {sheet_name}")
    print("Call LLA Api")
    # Define API URL, headers, and request body
    api_url = "https://api.jarvis-lla.com/api/v1.0/imports/payroll-shops"
    headers = {
        "X-API-Key": x_api_key,
        "Content-Type": "application/json"
    }
    payload = {
        "fromDate": from_date,
        "toDate": to_date
    }
    
    # Make the API call
    response = requests.post(api_url, headers=headers, json=payload)
    response_data = response.json()
    
   
    print("Fill shops info")
    if "shops" in response_data:
        cells_to_update.append(gspread.Cell(2, metric["Date"], from_date))
        cells_to_update.append(gspread.Cell(3, metric["Date"], to_date))
        for shop in response_data["shops"]:
            print(shop["name"])
            location_id = shop["locationId"]
            cell_found = main_worksheet.find(location_id, in_column=1)
            
            # If locationId is found in the worksheet
            if cell_found:
                row = cell_found.row
                # Prepare the cells to update
                cells_to_update.append(gspread.Cell(row, metric["Sales"], shop["sales"]))
                cells_to_update.append(gspread.Cell(row, metric["Car Count"], shop["carCount"]))
                cells_to_update.append(gspread.Cell(row, metric["Alignments"], shop["alignments"]))
                cells_to_update.append(gspread.Cell(row, metric["Tire Units"], shop["tires"]))
                cells_to_update.append(gspread.Cell(row, metric["Fluids"], shop["fluids"]))
                cells_to_update.append(gspread.Cell(row, metric["Brake Sales"], shop["brakes"]))
    
    # Update all the prepared cells in one go
    if cells_to_update:
        main_worksheet.update_cells(cells_to_update)       
    
    return "Shop data updated successfully"

def update_technicians_from_api(main_worksheet, payroll, metric):
    cells_to_update = []
    from_date, to_date, week_no, sheet_name = payroll
    print(f"FromDate: {from_date}, ToDate: {to_date}, WeekNo: {week_no}, SheetName: {sheet_name}")
    print("Call Technicians Summary Api")
    # Define API URL, headers, and request body
    api_url = "https://api.jarvis-lla.com/api/v1.0/imports/payroll-technicians-summary"
    headers = {
        "X-API-Key": x_api_key,
        "Content-Type": "application/json"
    }
    payload = {
        "fromDate": from_date,
        "toDate": to_date
    }
    
    # Make the API call
    response = requests.post(api_url, headers=headers, json=payload)
    response_data = response.json()
    
    # Get all employee IDs from the worksheet and store them in memory
    employee_ids = main_worksheet.col_values(1)
   
    print("Fill technicians info")
    if "employees" in response_data:
        for employee in response_data["employees"]:
            print(employee["employeeId"])
            employee_id = employee["employeeId"]
            
            # If employee_id is found in the in-memory list
            if employee_id in employee_ids:
                row = employee_ids.index(employee_id) + 1  # Adding 1 because list indices start from 0 while worksheet rows start from 1
                # Prepare the cells to update
                cells_to_update.append(gspread.Cell(row, metric["LaborH"], employee["laborHours"]))
                cells_to_update.append(gspread.Cell(row, metric["TechH"], employee["technicianHours"]))
                cells_to_update.append(gspread.Cell(row, metric["InvoicedL"], employee["invoicedL"]))
    
    # Update all the prepared cells in one go
    if cells_to_update:
        main_worksheet.update_cells(cells_to_update)       
    
    return "Technicians data updated successfully"


def update_attendance_from_api(main_worksheet, payroll, metric):
    cells_to_update = []
    from_date, to_date, week_no, sheet_name = payroll
    print(f"FromDate: {from_date}, ToDate: {to_date}, WeekNo: {week_no}, SheetName: {sheet_name}")
    print("Call Attendance Api")
    # Define API URL, headers, and request body
    api_url = "https://api.jarvis-lla.com/api/v1.0/imports/payroll-attendance-summary"
    headers = {
        "X-API-Key": x_api_key,
        "Content-Type": "application/json"
    }
    payload = {
        "fromDate": from_date,
        "toDate": to_date
    }
    
    # Make the API call
    response = requests.post(api_url, headers=headers, json=payload)
    response_data = response.json()
    
    # Get all employee IDs from the worksheet and store them in memory
    employee_ids = main_worksheet.col_values(1)
   
    print("Fill attendance info")
    if "employees" in response_data:
        for employee in response_data["employees"]:
            print(employee["employeeId"])
            employee_id = employee["employeeId"]

            # If employee_id is found in the in-memory list
            if employee_id in employee_ids:
                row = employee_ids.index(employee_id) + 1  # Adding 1 because list indices start from 0 while worksheet rows start from 1
                
                # Prepare the cells to update
                hours =  employee['workedHours']
                if (hours > 40):
                    cells_to_update.append(gspread.Cell(row, metric["Hours"],40))
                    cells_to_update.append(gspread.Cell(row, metric["Overtime"],hours-40))
                else:
                    cells_to_update.append(gspread.Cell(row, metric["Hours"],hours))
                    cells_to_update.append(gspread.Cell(row, metric["Overtime"],''))
                

    if cells_to_update:
        main_worksheet.update_cells(cells_to_update)       
    
    return "Attendance data updated successfully"


def update_attendance_from_file(main_worksheet, payroll, metric):
    from_date, to_date, week_no, sheet_name = payroll
    cells_to_update = []
    base_dir = Path('attendance') / execute_on_date
    file_path = base_dir / f"{execute_on_date}.xlsx"

     # Get all employee IDs from the worksheet and store them in memory
    employee_ids = main_worksheet.col_values(1)

    # Reading the sheet 'Grouped' from the Excel file
    df = pd.read_excel(file_path, sheet_name='Grouped')

    # Iterate over each row and print its content
    for index, employee in df.iterrows():
        employee_id = employee['EmployeeID']
        # If employee_id is found in the in-memory list
        if employee_id in employee_ids:
            row = employee_ids.index(employee_id) + 1  # Adding 1 because list indices start from 0 while worksheet rows start from 1
            
            # Prepare the cells to update
            hours =  employee['Total Hours Sum']
            if (hours > 40):
                cells_to_update.append(gspread.Cell(row, metric["Hours"],40))
                cells_to_update.append(gspread.Cell(row, metric["Overtime"],hours-40))
            else:
                cells_to_update.append(gspread.Cell(row, metric["Hours"],hours))
                cells_to_update.append(gspread.Cell(row, metric["Overtime"],''))
               

    if cells_to_update:
        main_worksheet.update_cells(cells_to_update)       
    
    return "Attendance data updated successfully"

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



payroll = get_payroll_period(spreadsheet, execute_on_date)
metric = get_payroll_settings(spreadsheet, payroll)

from_date, to_date, week_no, sheet_name = payroll
# Write the API data to the main worksheet
print("Get ot create payroll sheet")
main_worksheet = get_or_create_sheet(spreadsheet, sheet_name)

if payroll:
    result = update_shop_data_from_api(main_worksheet, payroll, metric)
    print(result)

    result = update_technicians_from_api(main_worksheet, payroll, metric)
    print(result)

    # result = update_attendance_from_api(main_worksheet, payroll, metric)
    result = update_attendance_from_file(main_worksheet, payroll, metric)
    print(result)
    
else:
    print(f"No data found for Execute On date: {execute_on_date}")

