import gspread
import configparser

from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build


spreadsheet_configs = {
    "PaycorPayroll2026": "1sbywsk3A3xdyO3280-GTd34tyHwwgzMCcWVLuJBhQoE",
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


def insert_rows_before_lla(spreadsheet):
    """
    Read the Paycor_Master sheet, find rows where second column has 'LLA',
    insert 8 empty rows before each of those rows without copying formatting,
    and move columns A,B,C,D from the LLA row up to the first inserted row.
    """
    master_sheet = spreadsheet.worksheet("Paycor_Master")
    sheet_id = master_sheet.id
    
    # Get all values from the sheet
    all_values = master_sheet.get_all_values()
    
    # Find all row indices where second column (index 1) has 'LLA'
    lla_rows = []
    for idx, row in enumerate(all_values):
        if len(row) > 1 and row[1] == "LLA":
            lla_rows.append(idx)  # 0-indexed
    
    print(f"Found {len(lla_rows)} rows with 'LLA' in second column: {[r + 1 for r in lla_rows]}")
    
    # Sort in reverse order to insert from bottom to top (prevents index shifting)
    lla_rows.sort(reverse=True)
    
    # Build batch update requests to insert rows without formatting
    requests = []
    for row_idx in lla_rows:
        print(f"Preparing to insert 8 rows before row {row_idx + 1}")
        requests.append({
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_idx,
                    "endIndex": row_idx + 8
                },
                "inheritFromBefore": False
            }
        })
        # Clear formatting for the inserted rows (set white background)
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 8
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        })
    
    # Execute all inserts in a single batch request
    if requests:
        print("Inserting rows...")
        spreadsheet.batch_update({"requests": requests})
    
    # Now move columns A,B,C,D from LLA rows up 7 rows
    # After insertion, LLA rows have shifted down by 7 * (number of LLA rows processed before them)
    # Since we processed bottom to top, we need to recalculate positions
    
    # Re-read the sheet to get updated values
    all_values = master_sheet.get_all_values()
    
    # Find LLA rows again (they've shifted)
    lla_rows_new = []
    for idx, row in enumerate(all_values):
        if len(row) > 1 and row[1] == "LLA":
            lla_rows_new.append(idx)  # 0-indexed
    
    print(f"LLA rows after insertion: {[r + 1 for r in lla_rows_new]}")
    
    # Prepare cells to update: move A,B,C,D up 7 rows and clear original
    cells_to_update = []
    for row_idx in lla_rows_new:
        target_row = row_idx - 8 + 1  # 1-indexed, 8 rows up
        source_row = row_idx + 1  # 1-indexed
        
        # Get values from columns A,B,C,D (indices 0-3)
        row_data = all_values[row_idx]
        for col in range(4):  # A, B, C, D
            value = row_data[col] if col < len(row_data) else ''
            # Add cell for target position (7 rows up)
            cells_to_update.append(gspread.Cell(target_row, col + 1, value))
            # Add cell to clear original position
            cells_to_update.append(gspread.Cell(source_row, col + 1, ''))
    
    if cells_to_update:
        print("Moving columns A,B,C,D up 8 rows...")
        master_sheet.update_cells(cells_to_update, value_input_option='RAW')
    
    return f"Inserted 8 rows before {len(lla_rows)} LLA rows and moved A,B,C,D up"


def copy_rows_after_lla(spreadsheet):
    """
    Find rows where second column has 'LLA', and copy rows 15-25 
    (with all formatting) to the row immediately after each LLA row.
    """
    master_sheet = spreadsheet.worksheet("Paycor_Master")
    sheet_id = master_sheet.id
    
    # Get all values from the sheet
    all_values = master_sheet.get_all_values()
    
    # Source rows are 15-25 (0-indexed: 14-24)
    source_start_row = 14  # 0-indexed (row 15)
    source_end_row = 25    # 0-indexed exclusive (row 25 inclusive)
    num_source_rows = source_end_row - source_start_row  # 11 rows
    
    print(f"Source rows 15-25 ({num_source_rows} rows)")
    
    # Find all row indices where second column (index 1) has 'LLA'
    lla_rows = []
    for idx, row in enumerate(all_values):
        if len(row) > 1 and row[1] == "LLA":
            lla_rows.append(idx)  # 0-indexed
    
    print(f"Found {len(lla_rows)} rows with 'LLA' in second column: {[r + 1 for r in lla_rows]}")
    
    # Build batch update requests to copy rows with formatting
    requests = []
    for lla_idx in lla_rows:
        # Target starts at row after LLA (lla_idx + 1 in 0-indexed)
        target_start_row = lla_idx + 1  # 0-indexed, row after LLA
        
        requests.append({
            "copyPaste": {
                "source": {
                    "sheetId": sheet_id,
                    "startRowIndex": source_start_row,
                    "endRowIndex": source_end_row,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(all_values[0]) if all_values else 26
                },
                "destination": {
                    "sheetId": sheet_id,
                    "startRowIndex": target_start_row,
                    "endRowIndex": target_start_row + num_source_rows,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(all_values[0]) if all_values else 26
                },
                "pasteType": "PASTE_NORMAL",
                "pasteOrientation": "NORMAL"
            }
        })
    
    if requests:
        print(f"Copying rows 15-25 (with formatting) after {len(lla_rows)} LLA rows...")
        spreadsheet.batch_update({"requests": requests})
    
    return f"Copied rows 15-25 (with formatting) after {len(lla_rows)} LLA rows"


def set_commission(spreadsheet):
    """
    For each location (marked by column B='LLA'):
    - Find the row where column P='Commission' within that location
    - Employees start at Commission row + 1
    - Employees end at the row where column H='Personnel over'
    - Set column P = $10 for all employees and clear background color
    """
    master_sheet = spreadsheet.worksheet("Paycor_Master")
    sheet_id = master_sheet.id
    
    # Get all values from the sheet
    all_values = master_sheet.get_all_values()
    
    # Column indices (0-indexed): B=1, H=7, P=15
    COL_B = 1   # Location marker (LLA)
    COL_H = 7   # Personnel over marker
    COL_P = 15  # Commission column
    
    # Find all location boundaries (rows with LLA in column B)
    lla_rows = []
    for idx, row in enumerate(all_values):
        if len(row) > COL_B and row[COL_B] == "LLA":
            lla_rows.append(idx)  # 0-indexed
    
    # Add end of sheet as final boundary
    lla_rows.append(len(all_values))
    
    print(f"Found {len(lla_rows) - 1} locations")
    
    # Collect all employee rows that need commission update
    employee_ranges = []
    
    for i in range(len(lla_rows) - 1):
        location_start = lla_rows[i]
        location_end = lla_rows[i + 1]
        
        # Find "Commission" row within this location (column P)
        commission_row = None
        for idx in range(location_start, location_end):
            row = all_values[idx]
            if len(row) > COL_P and row[COL_P] == "Commission":
                commission_row = idx
                break
        
        if commission_row is None:
            print(f"  Location at row {location_start + 1}: No 'Commission' found in column P")
            continue
        
        # Employees start at commission_row + 1
        employees_start = commission_row + 1
        
        # Find "Personnel over" row (column H) to mark end of employees
        employees_end = None
        for idx in range(employees_start, location_end):
            row = all_values[idx]
            if len(row) > COL_H and row[COL_H] == "Personnel over":
                employees_end = idx
                break
        
        if employees_end is None:
            print(f"  Location at row {location_start + 1}: No 'Personnel over' found in column H")
            continue
        
        if employees_start < employees_end:
            employee_ranges.append((employees_start, employees_end))
            print(f"  Location at row {location_start + 1}: Employees rows {employees_start + 1} to {employees_end}")
    
    # Build batch requests to update commission values and clear background
    requests = []
    cells_to_update = []
    
    for start_row, end_row in employee_ranges:
        # Set column P values to $10
        for row_idx in range(start_row, end_row):
            cells_to_update.append(gspread.Cell(row_idx + 1, COL_P + 1, ""))  # 1-indexed
        
        # Set formatting: Calibri 11, align right, top/bottom black borders, white background
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row,
                    "endRowIndex": end_row,
                    "startColumnIndex": COL_P,
                    "endColumnIndex": COL_P + 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 1.0,
                            "blue": 1.0
                        },
                        "textFormat": {
                            "fontFamily": "Calibri",
                            "fontSize": 11
                        },
                        "horizontalAlignment": "RIGHT",
                        "borders": {
                            "top": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                            "bottom": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                            "left": {"style": "NONE"},
                            "right": {"style": "NONE"}
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.borders"
            }
        })
        
        # Add conditional formatting rule for "R" in formula -> blue background (d9e2f3)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": sheet_id,
                        "startRowIndex": start_row,
                        "endRowIndex": end_row,
                        "startColumnIndex": COL_P,
                        "endColumnIndex": COL_P + 1
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": f'=ISNUMBER(FIND("R", FORMULATEXT(P{start_row + 1})))'}]
                        },
                        "format": {
                            "backgroundColor": {
                                "red": 0.851,    # d9 = 217/255
                                "green": 0.886,  # e2 = 226/255
                                "blue": 0.953    # f3 = 243/255
                            }
                        }
                    }
                },
                "index": 0
            }
        })
        
        # Add conditional formatting rule for "S" in formula -> orange background (fce4d6)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": sheet_id,
                        "startRowIndex": start_row,
                        "endRowIndex": end_row,
                        "startColumnIndex": COL_P,
                        "endColumnIndex": COL_P + 1
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": f'=ISNUMBER(FIND("S", FORMULATEXT(P{start_row + 1})))'}]
                        },
                        "format": {
                            "backgroundColor": {
                                "red": 0.988,    # fc = 252/255
                                "green": 0.894,  # e4 = 228/255
                                "blue": 0.839    # d6 = 214/255
                            }
                        }
                    }
                },
                "index": 1
            }
        })
    
    # Update values
    if cells_to_update:
        print(f"Setting commission to $10 for {len(cells_to_update)} employees...")
        master_sheet.update_cells(cells_to_update, value_input_option='RAW')
    
    # Apply formatting and conditional formatting
    if requests:
        print("Applying formatting and conditional formatting rules...")
        spreadsheet.batch_update({"requests": requests})
    
    return f"Set commission for {len(employee_ranges)} locations"


def set_total_payroll(spreadsheet):
    """
    For each location (marked by column B='LLA'):
    - Find the row where column T='Total Pay' within that location
    - Employees start at Total Pay row + 1
    - Employees end at the row where column S='Total Payroll'
    - Set column T of row Total Payroll as sum of employees column T
    """
    master_sheet = spreadsheet.worksheet("Paycor_Master")
    
    # Get all values from the sheet
    all_values = master_sheet.get_all_values()
    
    # Column indices (0-indexed): B=1, S=18, T=19
    COL_B = 1   # Location marker (LLA)
    COL_S = 18  # Total Payroll marker
    COL_T = 19  # Total Pay column
    
    # Find all location boundaries (rows with LLA in column B)
    lla_rows = []
    for idx, row in enumerate(all_values):
        if len(row) > COL_B and row[COL_B] == "LLA":
            lla_rows.append(idx)  # 0-indexed
    
    # Add end of sheet as final boundary
    lla_rows.append(len(all_values))
    
    print(f"Found {len(lla_rows) - 1} locations")
    
    # Collect all formulas to set
    cells_to_update = []
    
    for i in range(len(lla_rows) - 1):
        location_start = lla_rows[i]
        location_end = lla_rows[i + 1]
        
        # Find "Total Pay" row within this location (column T)
        total_pay_row = None
        for idx in range(location_start, location_end):
            row = all_values[idx]
            if len(row) > COL_T and row[COL_T] == "Total Pay":
                total_pay_row = idx
                break
        
        if total_pay_row is None:
            print(f"  Location at row {location_start + 1}: No 'Total Pay' found in column T")
            continue
        
        # Employees start at total_pay_row + 1
        employees_start = total_pay_row + 1
        
        # Find "Total Payroll" row (column S) to mark end of employees
        total_payroll_row = None
        for idx in range(employees_start, location_end):
            row = all_values[idx]
            if len(row) > COL_S and row[COL_S] == "Total Payroll":
                total_payroll_row = idx
                break
        
        if total_payroll_row is None:
            print(f"  Location at row {location_start + 1}: No 'Total Payroll' found in column S")
            continue
        
        employees_end = total_payroll_row
        
        if employees_start < employees_end:
            # Create SUM formula for column T (1-indexed rows for formula)
            formula = f"=SUM(T{employees_start + 1}:T{employees_end})"
            cells_to_update.append(gspread.Cell(total_payroll_row + 1, COL_T + 1, formula))  # 1-indexed
            print(f"  Location at row {location_start + 1}: Setting Total Payroll at row {total_payroll_row + 1} = SUM(T{employees_start + 1}:T{employees_end})")
    
    # Update values
    if cells_to_update:
        print(f"Setting Total Payroll formulas for {len(cells_to_update)} locations...")
        master_sheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')
    
    return f"Set Total Payroll for {len(cells_to_update)} locations"


def set_total_labor_percent(spreadsheet):
    """
    For each location (marked by column B='LLA'):
    - Find the row where column E='Sales' to get Sales value in column F
    - Find the row where column S='Total Payroll' to get Total Payroll value in column T
    - Direct Labor row is Total Payroll row + 1, update its formula to use Sales
    - Find the row where column S='Total Labor %'
    - Set column T of Total Labor % row as Total Payroll / Sales
    """
    import re
    
    master_sheet = spreadsheet.worksheet("Paycor_Master")
    
    # Get all values AND formulas from the sheet in one call
    all_values = master_sheet.get_all_values()
    
    # Get all formulas for column T in a single batch read
    # Using get() with value_render_option='FORMULA' to get formulas
    col_t_range = f"T1:T{len(all_values)}"
    all_formulas_t = master_sheet.get(col_t_range, value_render_option='FORMULA')
    
    # Column indices (0-indexed): B=1, E=4, F=5, S=18, T=19
    COL_B = 1   # Location marker (LLA)
    COL_E = 4   # Sales label column
    COL_S = 18  # Total Payroll / Total Labor % marker
    COL_T = 19  # Values column
    
    # Find all location boundaries (rows with LLA in column B)
    lla_rows = []
    for idx, row in enumerate(all_values):
        if len(row) > COL_B and row[COL_B] == "LLA":
            lla_rows.append(idx)  # 0-indexed
    
    # Add end of sheet as final boundary
    lla_rows.append(len(all_values))
    
    print(f"Found {len(lla_rows) - 1} locations")
    
    # Collect all cells to update (both Total Labor % and Direct Labor)
    cells_to_update = []
    
    for i in range(len(lla_rows) - 1):
        location_start = lla_rows[i]
        location_end = lla_rows[i + 1]
        
        # Find "Sales" row within this location (column E)
        sales_row = None
        for idx in range(location_start, location_end):
            row = all_values[idx]
            if len(row) > COL_E and row[COL_E] == "Sales":
                sales_row = idx
                break
        
        if sales_row is None:
            print(f"  Location at row {location_start + 1}: No 'Sales' found in column E")
            continue
        
        # Find "Total Payroll" row within this location (column S)
        total_payroll_row = None
        for idx in range(location_start, location_end):
            row = all_values[idx]
            if len(row) > COL_S and row[COL_S] == "Total Payroll":
                total_payroll_row = idx
                break
        
        if total_payroll_row is None:
            print(f"  Location at row {location_start + 1}: No 'Total Payroll' found in column S")
            continue
        
        # Direct Labor row is Total Payroll row + 1
        direct_labor_row = total_payroll_row + 1
        sales_cell = f"F{sales_row + 1}"
        
        # Get the current formula from the pre-fetched data
        if direct_labor_row < len(all_formulas_t) and len(all_formulas_t[direct_labor_row]) > 0:
            current_formula = all_formulas_t[direct_labor_row][0]
            
            if current_formula and current_formula.startswith('='):
                # Replace the divisor (cell reference after /) with the Sales cell reference
                new_formula = re.sub(r'/[A-Z]+\d+\)?$', f'/{sales_cell})', current_formula)
                if new_formula == current_formula:
                    # Try without the closing parenthesis
                    new_formula = re.sub(r'/[A-Z]+\d+$', f'/{sales_cell}', current_formula)
                
                # Remove surrounding parentheses: =(formula) => =formula
                if new_formula.startswith('=(') and new_formula.endswith(')'):
                    new_formula = '=' + new_formula[2:-1]
                
                if new_formula != current_formula:
                    cells_to_update.append(gspread.Cell(direct_labor_row + 1, COL_T + 1, new_formula))
                    print(f"    Direct Labor row {direct_labor_row + 1}: {current_formula} => {new_formula}")
                else:
                    print(f"    Direct Labor row {direct_labor_row + 1}: Formula unchanged: {current_formula}")
            else:
                print(f"    Direct Labor row {direct_labor_row + 1}: No formula found, value: {current_formula}")
        
        # Find "Total Labor %" row within this location (column S)
        total_labor_row = None
        for idx in range(location_start, location_end):
            row = all_values[idx]
            if len(row) > COL_S and row[COL_S] == "Total Labor %":
                total_labor_row = idx
                break
        
        if total_labor_row is None:
            print(f"  Location at row {location_start + 1}: No 'Total Labor %' found in column S")
            continue
        
        # Create formula: Total Payroll (T) / Sales (F)
        formula = f"=T{total_payroll_row + 1}/F{sales_row + 1}"
        cells_to_update.append(gspread.Cell(total_labor_row + 1, COL_T + 1, formula))
        print(f"  Location at row {location_start + 1}: Setting Total Labor % at row {total_labor_row + 1} = T{total_payroll_row + 1}/F{sales_row + 1}")
    
    # Update all cells in a single write
    if cells_to_update:
        print(f"Updating {len(cells_to_update)} cells in a single write...")
        master_sheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')
    
    return f"Set Total Labor % and Direct Labor for {len(cells_to_update)} cells"


# Main execution - process all spreadsheets
for name, current_spreadsheet_id in spreadsheet_configs.items():
    print(f"\n=== Processing {name} spreadsheet ===")
    
    spreadsheet = client.open_by_key(current_spreadsheet_id)
    
    # result = insert_rows_before_lla(spreadsheet)
    # print(result)
    
    # result = copy_rows_after_lla(spreadsheet)
    # print(result)
    
    # result = set_commission(spreadsheet)
    # print(result)
    
    # result = set_total_payroll(spreadsheet)
    # print(result)
    
    # result = set_total_labor_percent(spreadsheet)
    # print(result)
    
    print(f"✅ {name} processed successfully")
