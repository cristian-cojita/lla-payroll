import configparser
import requests
from datetime import datetime, timedelta


config = configparser.ConfigParser()
config.read('config/config.ini')
x_api_key = config['API']['X-API-Key']


def update_shop_data_from_api(day):
    print("Call LLA Api")
    api_url = f"https://api.jarvis-lla.com/api/v1.0/targets/setdayresults?date={day}"
    # api_url = f"https://lla-api.cojita.com/api/v1.0/targets/setdayresults?date={day}"
    headers = {
        "X-API-Key": x_api_key,
        "Content-Type": "application/json"
    }
    payload = {
        
    }
    
    # Make the API call
    response = requests.post(api_url, headers=headers, json=payload)
    


start_date = datetime(2023, 8, 27)
end_date = datetime(2023, 9, 30)

current_date = start_date
while current_date <= end_date:
    day=current_date.strftime('%Y-%m-%d')
    print(day)
    update_shop_data_from_api(day)
    current_date += timedelta(days=1)