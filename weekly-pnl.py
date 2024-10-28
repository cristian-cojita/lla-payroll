import gspread
import requests
import configparser
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta



summary_spreadsheet_id="14LNQnrTL6P5jBvnb7cvBVYEu3KNweIyYpAGUnreN0E4"
pnl_spreadsheet_id="1T12eCSi1rifYbP2oNoJIrEVNhFlNfTrfEGBTg_7YTLI"

# Set up gspread
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
scopes_string = ' '.join(scope)
creds = ServiceAccountCredentials.from_json_keyfile_name('config/lla-payroll-c3b730c6f614.json', scopes_string)
client = gspread.authorize(creds)

# Set up Google Sheets API client
service = build('sheets', 'v4', credentials=creds)

config = configparser.ConfigParser()
config.read("config/config.ini")
x_api_key = config["API"]["X-API-Key"]

summary_spreadsheet = client.open_by_key(summary_spreadsheet_id)
pnl_spreadsheet = client.open_by_key(pnl_spreadsheet_id)


def recreate_sheets_from_master(service, spreadsheet_id, source_sheet_id, target_sheet_names):
    """
    Deletes the first four sheets and recreates them as copies of the master sheet.
    
    Args:
        service: The Google Sheets API service instance.
        spreadsheet_id (str): The ID of the spreadsheet.
        source_sheet_id (int): The ID of the source (master) sheet to duplicate.
        target_sheet_names (list of str): Names of the new sheets to create.
    """
    # Get the spreadsheet details to identify sheet IDs
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_ids_to_delete = []

    # Collect the sheet IDs of the first four sheets to delete
    for sheet in spreadsheet['sheets'][:4]:  # Assuming the first four sheets are the ones to replace
        sheet_ids_to_delete.append(sheet['properties']['sheetId'])

    # Prepare delete requests
    delete_requests = [{"deleteSheet": {"sheetId": sheet_id}} for sheet_id in sheet_ids_to_delete]

    # Prepare duplication requests
    duplicate_requests = [
        {
            "duplicateSheet": {
                "sourceSheetId": source_sheet_id,
                "insertSheetIndex": i,  # Insert at the original index position
                "newSheetName": target_sheet_names[i]
            }
        }
        for i in range(len(target_sheet_names))
    ]

    # Combine delete and duplicate requests
    requests = delete_requests + duplicate_requests

    # Execute the batchUpdate to delete and recreate sheets
    body = {"requests": requests}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    print("Deleted and recreated sheets as copies of 'weeksMaster'.")

def get_pnl_data(from_date, to_date, payment_sheet, pnl_sheet):
    # Fetch all data from the payment_sheet at once to reduce API calls
    all_payment_data = payment_sheet.get_all_values()
    payment_records = all_payment_data[1:]  # Assuming the first row is the header

    pnl_data = pnl_sheet.get_all_values()  # Includes headers

    # Initialize an empty list to collect cells to update
    cells_to_update = []

    # Append cells for from_date and to_date
    cells_to_update.append(gspread.Cell(1, 1, f"{from_date.strftime('%Y-%m-%d')}  -  {to_date.strftime('%Y-%m-%d')}"))

    # Skip headers
    for i, master_row in enumerate(pnl_data[2:], start=3):  # Adjust index to start from the row 3
        location_id = master_row[0]  # Assuming LocationId is in the first column
        if location_id:
            print(f"Fetching data for location {master_row[1]}")

            # Call LLA API
            response = requests.post
            matching_record = next((row for row in payment_records if row[0] == location_id), None)

            if matching_record:
                # Assuming sales and payroll information is in specific columns in the payment_records
                sales = matching_record[2]
                payroll = matching_record[3]

                # Get values from LLA api
                # API endpoint
                url = "https://api.jarvis-lla.com/api/v1.0/tracker/period"
                headers = {"X-API-Key": x_api_key, "Content-Type": "application/json"}

                payload = {
                    "locationId": location_id,
                    "fromDate": from_date.strftime('%Y-%m-%d'),
                    "toDate": to_date.strftime('%Y-%m-%d')
                }
                
                response = requests.post(url, json=payload, headers=headers)

                if response.status_code == 200:
                    data = response.json()
                    
                    items = data.get('list', [])[0].get('items', []) 
                    totalCost = next((item for item in items if item["label"] == "Total Cost"), {}).get("value", 0)
                    grossProfit = next((item for item in items if item["label"] == "Gross Profit"), {}).get("value", 0)
                    
                else:
                    print(f"Failed to retrieve data, status code: {response.status_code}")
                    exit()
                    
                # Append data to cells_to_update list
                cells_to_update.append(gspread.Cell(i, 5, sales))  
                cells_to_update.append(gspread.Cell(i, 6, totalCost))  
                cells_to_update.append(gspread.Cell(i, 7, grossProfit))  
                cells_to_update.append(gspread.Cell(i, 8, payroll)) 

    pnl_sheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')
    




weeks_master_sheet = pnl_spreadsheet.worksheet("weeksMaster")
target_sheet_names = ["week-1", "week-2", "week-3", "week-4"]
source_sheet_id = weeks_master_sheet.id

# Recreate the first four sheets from 'weeksMaster'
recreate_sheets_from_master(service, pnl_spreadsheet_id, source_sheet_id, target_sheet_names)
for i in range(0, 4):
    try:
        payment_sheet = summary_spreadsheet.get_worksheet(i)
        pnl_sheet = pnl_spreadsheet.get_worksheet(i)
        
        # Set up date range for fetching data
        to_date = datetime.strptime(payment_sheet.title, "%Y-%m-%d") - timedelta(days=1)
        from_date = to_date - timedelta(days=6)
        print(f"{from_date.strftime('%Y-%m-%d')} - {to_date.strftime('%Y-%m-%d')}")
        
        # Fetch and update PnL data for the location
        get_pnl_data(from_date, to_date, payment_sheet, pnl_sheet)
    except ValueError:
        # This handles cases where the sheet title does not match the expected date format
        print(f"Skipping sheet '{payment_sheet.title}' - does not match expected date format.")

# pnl_sheets = pnl_spreadsheet.worksheets()
# sorted_sheets = sorted(pnl_sheets, key=lambda sheet: sheet.title)
# pnl_spreadsheet.reorder_worksheets(sorted_sheets)
