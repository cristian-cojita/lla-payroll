import datetime
import configparser

import gspread
import pandas as pd
import requests
from oauth2client.service_account import ServiceAccountCredentials

import util


regionals_spreadsheet_id = "1HjrLdSBuajLcm26Ci17DYrusTq13eTr4Z3LsEjWE6Ew"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
scopes_string = ' '.join(scope)
creds = ServiceAccountCredentials.from_json_keyfile_name('config/lla-payroll-c3b730c6f614.json', scopes_string)
client = gspread.authorize(creds)
regionals_spreadsheet = client.open_by_key(regionals_spreadsheet_id)

config = configparser.ConfigParser()
config.read('config/config.ini')
x_api_key = config['API']['X-API-Key']


def get_regionals_period() -> str:
    override = config.get('Settings', 'regionals_period', fallback='').strip()
    if override:
        return override
    today = datetime.date.today()
    last_day_prev_month = today.replace(day=1) - datetime.timedelta(days=1)
    return last_day_prev_month.strftime('%Y-%m')


def get_or_create_sheet(spreadsheet, sheet_name):
    sheet_names = [sheet.title for sheet in spreadsheet.worksheets()]
    if sheet_name not in sheet_names:
        master_sheet = spreadsheet.worksheet("Master")
        master_sheet.duplicate(new_sheet_name=sheet_name)

        new_sheet = spreadsheet.worksheet(sheet_name)

        other_sheets = [sheet for sheet in spreadsheet.worksheets() if sheet.title != sheet_name]
        all_sheets_in_order = [new_sheet] + other_sheets
        spreadsheet.reorder_worksheets(all_sheets_in_order)

        return new_sheet
    return spreadsheet.worksheet(sheet_name)


def find_column_indices(header_row, required_headers):
    """Return 1-based column indices for each required header (case-insensitive, trimmed)."""
    normalized = [h.strip().lower() for h in header_row]
    indices = {}
    for header in required_headers:
        target = header.strip().lower()
        if target not in normalized:
            raise ValueError(f"Header '{header}' not found in sheet. Got: {header_row}")
        indices[header] = normalized.index(target) + 1
    return indices


def fetch_monthly_sales_goal_by_location(period):
    """Call Jarvis monthly-targets API and return {location_id: tier1 sales goal}."""
    month_number = int(period.replace('-', ''))
    url = "https://jarvis-lla.com/api/v1.0/imports/monthly-targets"
    headers = {"X-API-Key": x_api_key, "Content-Type": "application/json"}
    response = requests.post(url, json={"MonthNumber": month_number}, headers=headers)
    response.raise_for_status()
    data = response.json()
    return {
        loc["locationId"]: loc.get("sales", {}).get("tier1")
        for loc in data.get("locations", [])
        if loc.get("locationId")
    }


def fetch_pnl_by_location(period):
    engine = util.create_conn()
    query = """
        SELECT s.location_id, qp.net_income, qp.total_sales
        FROM qb_pnl qp
        JOIN shops s ON qp.shop_id = s.id
        WHERE qp.period = %(period)s
    """
    df = pd.read_sql_query(query, engine, params={"period": period})
    return {str(row["location_id"]): (row["total_sales"], row["net_income"]) for _, row in df.iterrows()}


def fill_regionals(worksheet, pnl_by_location, monthly_goal_by_location):
    print("fill_regionals")
    all_values = worksheet.get_all_values()
    header_row = all_values[0]
    cols = find_column_indices(header_row, ["LocationId", "Sales", "Contributions", "Monthly Sales Goal"])
    locationid_col = cols["LocationId"]
    sales_col = cols["Sales"]
    contributions_col = cols["Contributions"]
    monthly_goal_col = cols["Monthly Sales Goal"]

    cells_to_update = []
    rows_updated = 0
    for i, row in enumerate(all_values[1:], start=2):
        if len(row) < locationid_col:
            continue
        location_id = row[locationid_col - 1].strip()
        if not location_id:
            continue

        pnl_match = pnl_by_location.get(location_id)
        if pnl_match is not None:
            total_sales, net_income = pnl_match
            cells_to_update.append(gspread.Cell(row=i, col=sales_col, value=str(float(total_sales)) if total_sales is not None else ''))
            cells_to_update.append(gspread.Cell(row=i, col=contributions_col, value=str(float(net_income)) if net_income is not None else ''))

        monthly_goal = monthly_goal_by_location.get(location_id)
        if monthly_goal is not None:
            cells_to_update.append(gspread.Cell(row=i, col=monthly_goal_col, value=str(float(monthly_goal))))

        if pnl_match is not None or monthly_goal is not None:
            rows_updated += 1

    if cells_to_update:
        worksheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')
    print(f"Updated {rows_updated} rows ({len(cells_to_update)} cells)")


period = get_regionals_period()
print(f"Regionals period: {period}")

worksheet = get_or_create_sheet(regionals_spreadsheet, period)
pnl_by_location = fetch_pnl_by_location(period)
print(f"Fetched {len(pnl_by_location)} location rows from qb_pnl for period {period}")

monthly_goal_by_location = fetch_monthly_sales_goal_by_location(period)
print(f"Fetched {len(monthly_goal_by_location)} monthly sales goals for period {period}")

fill_regionals(worksheet, pnl_by_location, monthly_goal_by_location)
