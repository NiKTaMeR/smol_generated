import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import logging
from openpyxl import load_workbook

def ingest_data():
    root = tk.Tk()
    root.withdraw()

    file_types = [("Excel files", "*.xlsx"), ("Access Database", "*.accdb")]
    file_path = filedialog.askopenfilename(title="Select the raw data file", filetypes=file_types)

    if file_path.endswith(".xlsx"):
        data = pd.read_excel(file_path)
    elif file_path.endswith(".accdb"):
        import pyodbc
        conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file_path + ';'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        table_name = cursor.tables().fetchone().table_name
        data = pd.read_sql(f"SELECT * FROM {table_name}", conn)
        conn.close()
    else:
        logging.error("Invalid file type selected.")
        raise ValueError("Invalid file type selected.")

    return data