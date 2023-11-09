import gspread
import requests
import configparser
import pandas as pd

from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import get_effective_format

from pathlib import Path
import datetime


# main
spreadsheet_id="1RXtJB5s1cGnHoKgnQ1R8RQq5Ng9x0OGVRbyFL-_mGEE"


# test
# spreadsheet_id="1LeaBFUg-VMpp1pxqCD0cx2GVmfSG90KdlMolkglfgxc"




scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('config/lla-payroll-c3b730c6f614.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(spreadsheet_id)

config = configparser.ConfigParser()
config.read('config/config.ini')
x_api_key = config['API']['X-API-Key']


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

def get_value_case_insensitive(dictionary, key):
    lower_key = key.lower()
    for k, v in dictionary.items():
        if k.lower() == lower_key:
            return v
    return None


def update_shop_data_from_api(current_worksheet, day):
    cells_to_update = []
    properties = ["Sales", "Car Count", "Brake Sales", "Tire Units", "Fluids","Alignments"]
    days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

    all_data = current_worksheet.get_all_values()

    print("Call LLA Api")
    # Define API URL, headers, and request body
    api_url = "https://lla-api.cojita.com/api/v1/imports/week-targets"
    headers = {
        "X-API-Key": x_api_key,
        "Content-Type": "application/json"
    }
    payload = {
        
    }
    
    # Make the API call
    response = requests.post(api_url, headers=headers, json=payload)
    response_data = response.json()
    
   
    print("Fill shops info")
    if "shops" in response_data:
        for shop in response_data["shops"]:
            print(shop["name"])
            location_id = shop["locationId"]
            
            shop_row = next((i for i, row in enumerate(all_data) if row[0] == location_id), None)


            if shop_row is not None:
                for prop in properties:
                    prop_row = None
                    
                    for row in range(shop_row+1, min(shop_row+5, len(all_data))):  
                        if all_data[row][1] == prop:
                            prop_row = row
                            break

                    if prop_row:
                        jsonProp = (
                                    prop.lower()
                                    .replace("brake sales", "brakes")
                                    .replace(" ", "")
                                    .replace("unit", "")
                                )
                        print(f"{prop} - {jsonProp}")

                        day_col = next((i for i, cell in enumerate(all_data[shop_row]) if cell == day), None)

                        if day_col is not None:
                            value = get_value_case_insensitive(shop, jsonProp)
                            if value is not None:
                                value = value[day.lower()]

                            cell = gspread.Cell(row=prop_row + 1, col=day_col + 1, value=value)  # +1 because gspread uses 1-based indexing
                            cells_to_update.append(cell)

    # Update all the prepared cells in one go
    if cells_to_update:
        current_worksheet.update_cells(cells_to_update)

    return "Shop data updated successfully"


 
current_date = datetime.datetime.now() - datetime.timedelta(days=1)
day_name = current_date.strftime('%A')

clark = spreadsheet.worksheet("Clark")
result = update_shop_data_from_api(clark, day_name)

modica = spreadsheet.worksheet("Modica")
result = update_shop_data_from_api(modica, day_name)
print(result)
