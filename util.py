from sqlalchemy import create_engine
import configparser
import psycopg2
import pandas as pd


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
    