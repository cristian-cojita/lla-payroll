import datetime
from sqlalchemy import create_engine
import configparser
import psycopg2
import pandas as pd
from dataclasses import dataclass

@dataclass
class PayrollInfo:
    """A data class to hold payroll date information."""
    execute_on_date: str
    payroll_from: str
    payroll_to: str


def config(filename='config/config.ini', section='postgresql'):
    parser = configparser.ConfigParser()
    parser.read(filename)
    db = {}
    print(parser.sections())
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception(f'Section {section} not found in the {filename} file')
    return db

def create_conn():
    params = config()
    db_url = f"postgresql://{params['user']}:{params['password']}@{params['host']}:{params['port']}/{params['dbname']}"
    engine = create_engine(db_url)
    return engine

def get_shops(engine):
    query = """SELECT s.*, r.name  as region_name, r.id as region_id, r.color, r.order_no as region_order_no 
            FROM shops s 
            JOIN regions r on s.regionid = r.id  
            ORDER BY r.order_no, s.order_no"""
    df = pd.read_sql_query(query, engine)
    return df

def get_regions(engine):
    query = "SELECT r.id, r.name, r.color, r.order_no as region_order_no FROM regions r ORDER BY r.order_no desc"
    df = pd.read_sql_query(query, engine)
    return df
   
   
def execute_on_date() -> PayrollInfo:
    parser = configparser.ConfigParser()
    parser.read('config/config.ini')
    
    execute_on_date_from_config: str = parser.get('Settings', 'execute_on_date', fallback='').strip()

    base_date: datetime.date
    if execute_on_date_from_config:
        # If a date is provided in the config, use it
        base_date = datetime.datetime.strptime(execute_on_date_from_config, '%Y-%m-%d').date()
    else:
        # Otherwise, calculate the most recent Sunday
        today = datetime.date.today()
        # today.weekday() is Monday 0, Sunday 6.
        days_to_subtract = (today.weekday() + 1) % 7
        base_date = today - datetime.timedelta(days=days_to_subtract)

    # Calculate payroll_from and payroll_to based on the base_date
    payroll_from_date = base_date - datetime.timedelta(days=7)
    payroll_to_date = base_date - datetime.timedelta(days=1)

    # Create and return a PayrollInfo object
    return PayrollInfo(
        execute_on_date=base_date.strftime('%Y-%m-%d'),
        payroll_from=payroll_from_date.strftime('%Y-%m-%d'),
        payroll_to=payroll_to_date.strftime('%Y-%m-%d')
    )