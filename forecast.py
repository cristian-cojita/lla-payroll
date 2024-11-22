import gspread
import gspread.utils
import requests
import util
import configparser
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta

# Define the list of periods
periods = [
    '2024-01',
    '2024-02',
    '2024-03',
    '2024-04',
    '2024-05',
    '2024-06',
    '2024-07',
    '2024-08',
    '2024-09',
    '2024-10',
    '2024-11',
    'month forecast',
    'year forecast',
]


summary_spreadsheet_id="14LNQnrTL6P5jBvnb7cvBVYEu3KNweIyYpAGUnreN0E4"
forecast_spreadsheet_id="1qua94fnEVv1vT45H3bTBTjRhjfBVP3_EcTOjXjCu6pw"

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
forecast_spreadsheet = client.open_by_key(forecast_spreadsheet_id)


# Read the master sheet
master_sheet = forecast_spreadsheet.worksheet("Master")
master_data = master_sheet.get_all_values()


def create_worksheet(spreadsheet, shop, rows, cols):
    sheet_name = shop["short_name"]
    try:
        # Try to open the worksheet by name
        worksheet = spreadsheet.worksheet(sheet_name)
        # If it exists, delete it
        print(f"Worksheet '{sheet_name}' exists. Deleting it.")
        spreadsheet.del_worksheet(worksheet)
    except gspread.exceptions.WorksheetNotFound:
        # If the worksheet doesn't exist, proceed to create it
        print(f"Worksheet '{sheet_name}' does not exist. Creating a new one.")

    # Create a new worksheet
    print(f"Creating worksheet '{sheet_name}'.")
    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=rows, cols=cols)

    sheet_id = worksheet.id
    
    r, g, b = int(shop['color'][0:2], 16), int(shop['color'][2:4], 16), int(shop['color'][4:6], 16)
    tab_color = {
        "red": r / 255,
        "green": g / 255,
        "blue": b / 255,
    }
    # Update the tab color
    service.spreadsheets().batchUpdate(
        spreadsheetId=forecast_spreadsheet_id,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "tabColor": tab_color,
                        },
                        "fields": "tabColor",
                    }
                }
            ]
        }
    ).execute()


    start_row = 0  # Row 1 in 0-based index
    merge_requests = []

    # Collect cells for updating titles
    # worksheet = forecast_spreadsheet.get_worksheet_by_id(sheet_id)  # Retrieve the worksheet using its ID
    cells_to_update = []

    # Loop through periods to prepare merge and value updates
    for i, period in enumerate(periods):
        start_col = 1 + i * 2  # Columns start at 'B' 
        end_col = start_col + 2  # Merge two columns per period

        # Merge cells request
        merge_requests.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": start_row + 1,  # Merge only the first row
                    "startColumnIndex": start_col,
                    "endColumnIndex": end_col,
                },
                "mergeType": "MERGE_ALL"
            }
        })

        # Add period name into the first cell of the merged range
        cell = gspread.Cell(1, start_col + 1, period)  # gspread uses 1-based indexing for rows and columns
        cells_to_update.append(cell)
        cell = gspread.Cell(2, start_col + 1, 'Actual')  
        cells_to_update.append(cell)
        cell = gspread.Cell(2, start_col + 2, 'Goal')  
        cells_to_update.append(cell)

    # Execute merge requests
    batch_update_body = {
        "requests": merge_requests
    }
    service.spreadsheets().batchUpdate(
        spreadsheetId=forecast_spreadsheet_id,
        body=batch_update_body
    ).execute()
    print("Cells merged successfully.")

    objectives = [
        "Sales", "",
        "Tire Sales", "",
        "Tire Sales %", "", 
        "Mech Sales",  "",
        "Mech Sales %", "",
        "Tire Cogs",  "",
        "Tire Cogs %",  "",
        "Mech Cogs",  "",
        "Mech Cogs %", "",
        "Total Cogs",  "",
        "Total Cogs %",  "",
        "Gross Profit", 
        "Gross Profit %",
        "Wages", 
        "Wage %", 
        "Expenses", 
        "Expenses %", 
        "Financial Effic",
        "Financial Effic %", 
        "Tire Units", "",
        "Car Count" "",
    ]

    row = 3  # Start from row 3
    for objective in objectives:
        cells_to_update.append(gspread.Cell(row, 1, objective))  
        row += 1 

    worksheet.update_cells(cells_to_update)
    print("Titles updated successfully.")
    return sheet_id

def find_goal(location_id, objective):
    """
    Finds the cell in the 'Master' sheet corresponding to the given location_id and objective,
    and returns a formula referencing that cell.
    
    Args:
        location_id (str): The location ID to search for.
        objective (str): The objective column to search for.
    
    Returns:
        str: A formula referencing the cell in the 'Master' sheet (e.g., '=Master!A1').
    """
    # Get all data from the 'Master' sheet
    
    header_row = master_data[0]  # The first row contains column headers

    # Find the column index for the objective
    try:
        objective_col_idx = header_row.index(objective) + 1  # Convert to 1-based index
    except ValueError:
        print(f"Objective '{objective}' not found in 'Master' sheet.")
        return None

    # Find the row for the given location_id
    for row_idx, row in enumerate(master_data[1:], start=2):  # Skip the header row
        if row[0] == str(location_id):  # Match the location_id in the first column
            # Construct the formula referencing the cell
            return f"=Master!{gspread.utils.rowcol_to_a1(row_idx, objective_col_idx)}"

    print(f"Location ID '{location_id}' not found in 'Master' sheet.")
    return None


def find_cell_by_label(ws, label, column_index):
    """
    Finds the row index by matching a label in the specified column of the worksheet.
    
    Args:
        ws (Worksheet): The gspread worksheet object.
        label (str): The label to search for.
        column_index (int): The column index to search in (1-based index).
    
    Returns:
        int: The row index (1-based) if the label is found, otherwise None.
    """
    # Get all values from the specified column
    column_values = ws.col_values(column_index)

    # Find the row index for the label
    for idx, value in enumerate(column_values, start=1):  # 1-based index
        if value.strip().lower() == label.strip().lower():  # Case-insensitive match
            return idx

    print(f"Label '{label}' not found in column {column_index}.")
    return None

def fill_goals(shop_sheet_id, shop):
    # Open the shop's sheet using its ID
    shop_sheet = forecast_spreadsheet.get_worksheet_by_id(shop_sheet_id)
    shop_sheet_header = shop_sheet.row_values(2)  # Assuming the header is in the second row

    # Identify all columns in the shop sheet labeled as "Goal"
    goal_columns = {
        i: col_name for i, col_name in enumerate(shop_sheet_header, start=1) if col_name.strip().lower() == "goal"
    }

    cells_to_update = []
    objective = 'Sales'

    for goal_col_idx in goal_columns:
        # Find the corresponding formula
        formula = find_goal(shop["location_id"], objective)
        if formula:
            # Locate the correct row for the given objective
            current_row = find_cell_by_label(shop_sheet, objective, 1)
            if current_row:
                # Append the formula as a value to the cell
                cell = gspread.Cell(current_row, goal_col_idx, formula)
                shop_sheet.format(f"A{current_row}:A{current_row}", {
                    "textFormat": {
                        "fontSize": 8
                    }
                })

                cells_to_update.append(cell)


    # Update cells in the shop's sheet
    if cells_to_update:
        shop_sheet.update_cells(cells_to_update, value_input_option="USER_ENTERED")
        print(f"Goals filled for shop '{shop['short_name']}'.")
    else:
        print(f"No goals found to update for shop '{shop['short_name']}'.")



def fill_shop(shop):
    print(shop["location_id"], ', ', shop["short_name"], ', ', shop['region_name'])
    rows = 100
    cols = (len(periods) + 4) * 2
    shop_sheet_id = create_worksheet(forecast_spreadsheet, shop, rows, cols)
    fill_goals(shop_sheet_id, shop)

 
def main():
    conn = util.create_conn()
    shops = util.get_shops(conn)
    for index, shop in shops.head(3).iterrows():
        fill_shop(shop)
    
    


if __name__ == "__main__":
    main()
