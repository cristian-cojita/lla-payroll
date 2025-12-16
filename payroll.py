import gspread
import gspread.utils
import requests
import configparser
import pandas as pd
import datetime
import time

from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from pathlib import Path

import util
from types_payroll import DEFAULT_PAYROLL_METRICS, PayrollMetrics
from gspread_formatting import CellFormat, Color, format_cell_ranges



spreadsheet_configs = {
    # "PaycorPayroll2026": "1sbywsk3A3xdyO3280-GTd34tyHwwgzMCcWVLuJBhQoE",
    "PaycorPayroll": "1_I9CIGk3CcTIJP5u8xgDXiUPQpN_zASGI3T4RhgH4Ro",
    "PaycorOfficePayroll": "15a4QNJ36WCu0Rt5gU23skj4gEa-Yz1SRM5Jrts8pQJE"
}



scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
]
scopes_string = " ".join(scope)
creds = ServiceAccountCredentials.from_json_keyfile_name(
    "config/lla-payroll-c3b730c6f614.json", scopes_string
)
client = gspread.authorize(creds)
drive_service = build('drive', 'v3', credentials=creds)

config = configparser.ConfigParser()
config.read("config/config.ini")
x_api_key = config["API"]["X-API-Key"]

payroll_info = util.get_period()
cell_format_flat_rate = CellFormat(backgroundColor=Color(0, 1, 1))  # Cyan color


def get_or_create_sheet(spreadsheet):
    sheet_name = payroll_info.execute_on_date
    # Check if worksheet with the name 'sheet_name' exists
    sheet_names = [sheet.title for sheet in spreadsheet.worksheets()]
    if sheet_name not in sheet_names:
        # Duplicate the "Master" worksheet and name it as 'sheet_name'
        master_sheet = spreadsheet.worksheet("Paycor_Master")
        master_sheet.duplicate(new_sheet_name=sheet_name)

        # Re-fetch the new sheet to ensure we have the full worksheet object
        new_sheet = spreadsheet.worksheet(sheet_name)

        # Prepare the list of Worksheet objects in the desired order
        other_sheets = [
            sheet for sheet in spreadsheet.worksheets() if sheet.title != sheet_name
        ]
        all_sheets_in_order = [new_sheet] + other_sheets

        # Reorder the sheets
        spreadsheet.reorder_worksheets(all_sheets_in_order)

        return new_sheet
    else:
        return spreadsheet.worksheet(sheet_name)


def update_shop_data_from_api(main_worksheet, metric):
    cells_to_update = []
    
    print(
        f"FromDate: {payroll_info.payroll_from}, ToDate: {payroll_info.payroll_to}"
    )
    print("Call LLA Api")
    # Define API URL, headers, and request body
    api_url = "https://jarvis-lla.com/api/v1.0/imports/payroll-locations"
    headers = {"X-API-Key": x_api_key, "Content-Type": "application/json"}
    payload = {"fromDate": payroll_info.payroll_from, "toDate": payroll_info.payroll_to}

    # Make the API call
    response = requests.post(api_url, headers=headers, json=payload)
    response_data = response.json()

    print("Fill locations info")
    if "locations" in response_data:
        # Write the payroll date range at the top of the sheet
        cells_to_update.append(gspread.Cell(2, metric.date, payroll_info.payroll_from))
        cells_to_update.append(gspread.Cell(3, metric.date, payroll_info.payroll_to))
        # Fetch all location IDs from the worksheet once
        worksheet_location_ids = main_worksheet.col_values(1)  # 1-based index

        for location in response_data["locations"]:
            print(location["name"])
            location_id = location["locationId"]
            if location_id in worksheet_location_ids:
                row = worksheet_location_ids.index(location_id) + 1  # rows are 1-based
                # Prepare the cells to update
                cells_to_update.append(
                    gspread.Cell(row, metric.sales, location["sales"])
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.car_bonus, location["carBonus"])
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.alignments, location["alignments"])
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.tire_units, location["tires"])
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.fluids, location["fluids"])
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.brake_sales, location["brakes"])
                )

    # Update all the prepared cells in one go
    if cells_to_update:
        main_worksheet.update_cells(cells_to_update)

    return "Shop data updated successfully"


def update_technicians_from_api(main_worksheet, metric):
    cells_to_update = []
    print(
        f"FromDate: {payroll_info.payroll_from}, ToDate: {payroll_info.payroll_to}"
    )
    print("Call Technicians Summary Api")
    # Define API URL, headers, and request body
    api_url = "https://jarvis-lla.com/api/v1.0/imports/payroll-technicians-summary"
    headers = {"X-API-Key": x_api_key, "Content-Type": "application/json"}
    payload = {"fromDate": payroll_info.payroll_from, "toDate": payroll_info.payroll_to}

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
                row = (
                    employee_ids.index(employee_id) + 1
                )  # Adding 1 because list indices start from 0 while worksheet rows start from 1
                # Prepare the cells to update
                cells_to_update.append(
                    gspread.Cell(row, metric.paycor_id, employee.get("paycorId", ""))
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.labor_h, employee.get("laborHours", ""))
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.tech_h, employee.get("technicianHours", ""))
                )
                cells_to_update.append(
                    gspread.Cell(row, metric.invoiced_l, employee.get("invoicedL", ""))
                )

    # Update all the prepared cells in one go
    if cells_to_update:
        main_worksheet.update_cells(cells_to_update)

    return "Technicians data updated successfully"


def update_attendance_from_api(main_worksheet, metric):
    cells_to_update = []
    cells_flat_rate = []
    print(
        f"FromDate: {payroll_info.payroll_from}, ToDate: {payroll_info.payroll_to}"
    )
    print("Call Attendance Api")
    # Define API URL, headers, and request body
    api_url = "https://jarvis-lla.com/api/v1.0/imports/payroll-attendance-summary"
    headers = {"X-API-Key": x_api_key, "Content-Type": "application/json"}
    payload = {"fromDate": payroll_info.payroll_from, "toDate": payroll_info.payroll_to}

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
                row = employee_ids.index(employee_id) + 1 # Adding 1 because list indices start from 0 while worksheet rows start from 1
                if "paycorId" in employee:
                    cells_to_update.append(gspread.Cell(row, metric.paycor_id, employee["paycorId"]))
                cells_to_update.append(gspread.Cell(row, metric.clocked_hours, employee["workedHours"]))
                
                if(employee["paymentType"] == "FlatRate"):
                    tech_h_cell = gspread.utils.rowcol_to_a1(row, metric.tech_h)
                    hours_cell = gspread.utils.rowcol_to_a1(row, metric.hours)
                    cells_to_update.append(gspread.Cell(row, metric.hours, '=' + tech_h_cell))
                    cells_flat_rate.append((hours_cell, cell_format_flat_rate))
                else:
                    cells_to_update.append(gspread.Cell(row, metric.hours, str(40 if employee["workedHours"] > 40 else employee["workedHours"])))

                if "overtimeHours" in employee:
                    cells_to_update.append(gspread.Cell(row, metric.overtime, employee["overtimeHours"]))

    if cells_to_update:
        main_worksheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')

    # Apply cyan background to flat rate cells
    if cells_flat_rate:
        format_cell_ranges(main_worksheet, cells_flat_rate)

    return "Attendance data updated successfully"


def create_backup(spreadsheet, spreadsheet_id):
    # Backup configuration
    backup_folder_id = "1iSsFLYs7v3P3Q8BmJBMJcMBN05Mzx53h"  # ID of the "Backup" folder
    today = datetime.date.today()
    today_str = today.strftime('%Y-%m-%d')  # Format yyyy-mm-dd

    # Check if a subfolder for today already exists
    query = f"mimeType='application/vnd.google-apps.folder' and '{backup_folder_id}' in parents and name='{today_str}'"
    response = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = response.get('files', [])

    if not files:
        # Create a new folder
        file_metadata = {
            'name': today_str,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [backup_folder_id]
        }
        folder = drive_service.files().create(body=file_metadata, fields='id').execute()
        folder_id = folder.get('id')
    else:
        # Folder already exists
        folder_id = files[0].get('id')

    # Copy the spreadsheet to the new folder
    copy_title = f"{spreadsheet.title} ({today_str})"
    copy_metadata = {
        'name': copy_title,
        'parents': [folder_id]
    }
    drive_service.files().copy(fileId=spreadsheet_id, body=copy_metadata).execute()
    print("Backup completed successfully.")


# Main execution - process all spreadsheets
for name, current_spreadsheet_id in spreadsheet_configs.items():
    print(f"\n=== Processing {name} spreadsheet ===")
        
    spreadsheet = client.open_by_key(current_spreadsheet_id)

    # Create a backup file in the "Backup" folder
    create_backup(spreadsheet, current_spreadsheet_id)

    metric: PayrollMetrics = DEFAULT_PAYROLL_METRICS

  
    # Write the API data to the main worksheet
    print("Get or create payroll sheet")
    main_worksheet = get_or_create_sheet(spreadsheet)

    result = update_shop_data_from_api(main_worksheet, metric)
    print(result)

    result = update_technicians_from_api(main_worksheet, metric)
    print(result)

    result = update_attendance_from_api(main_worksheet, metric)
    print(f"✅ {name} processed successfully")
   
