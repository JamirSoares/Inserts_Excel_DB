from pyodbc import connect, Error
from dotenv import load_dotenv
from os import getenv

load_dotenv()

def get_connection():
    try:
        conn = connect(getenv("SQL_CONNECTION_STRING"))
        return conn
    except Error as e:
        print("Erro ao conectar no banco:")
        print(e)
        return None
