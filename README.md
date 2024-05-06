# Excel-to-SQL-Import

import openpyxl
import csv
import os
import pyodbc
 
# Folder path containing Excel and CSV files
folder_path = r'D:\xlxfiles'
 
# SQL Server connection parameters
server = '20.204.123.250'
database = 'hotel booking'
username = 'harsh'
password = 'laminateabdH@12'
 
# Establish connection to SQL Server
conn_str = f"DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password}"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
 
try:
    # Iterate over all files in the folder
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            # Extract file name without extension
            table_name = os.path.splitext(file_name)[0]
 
            # Open the Excel file
            workbook = openpyxl.load_workbook(file_path)
 
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
 
                # Extract column names from the first row
                columns = [str(cell.value) for cell in sheet[1]]
 
                # Construct SQL query to create table
                create_table_query = f"CREATE TABLE [{table_name}] ({', '.join([f'[{column}] NVARCHAR(MAX)' for column in columns])})"
 
                # Drop the table if it already exists
                try:
                    cursor.execute(f"DROP TABLE IF EXISTS [{table_name}]")
                    conn.commit()
                    print(f"Table '{table_name}' dropped successfully.")
                except pyodbc.Error as e:
                    print(f"Error dropping table: {e}")
 
                # Execute SQL query to create table
                cursor.execute(create_table_query)
                conn.commit()
 
                # Insert data into the table
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    # Convert all row values to strings
                    row = [str(value) for value in row]
                    insert_query = f"INSERT INTO [{table_name}] ({', '.join(['[' + col + ']' for col in columns])}) VALUES ({', '.join(['?'] * len(row))})"            
                    cursor.execute(insert_query, row)
                    conn.commit()
 
                print(f"Data from '{file_path}' sheet '{sheet_name}' has been exported to table '{table_name}' in SSMS.")
 
            workbook.close()  # Close the workbook after processing
 
        elif file_name.endswith('.csv'):
            # Extract file name without extension
            table_name = os.path.splitext(file_name)[0]
 
            # Open the CSV file
            with open(file_path, 'r', newline='') as csvfile:
                csv_reader = csv.reader(csvfile)
                columns = next(csv_reader)  # Extract column names from the first row
 
                # Construct SQL query to create table
                create_table_query = f"CREATE TABLE [{table_name}] ({', '.join([f'[{column}] NVARCHAR(MAX)' for column in columns])})"
 
                # Drop the table if it already exists
                try:
                    cursor.execute(f"DROP TABLE IF EXISTS [{table_name}]")
                    conn.commit()
                    print(f"Table '{table_name}' dropped successfully.")
                except pyodbc.Error as e:
                    print(f"Error dropping table: {e}")
 
                # Execute SQL query to create table
                cursor.execute(create_table_query)
                conn.commit()
 
                # Insert data into the table
                for row in csv_reader:
                    # Convert all row values to strings
                    row = [str(value) for value in row]
                    insert_query = f"INSERT INTO [{table_name}] ({', '.join(['[' + col + ']' for col in columns])}) VALUES ({', '.join(['?'] * len(row))})"            
                    cursor.execute(insert_query, row)
                    conn.commit()
 
                print(f"Data from '{file_path}' has been exported to table '{table_name}' in SSMS.")
 
finally:
    # Close connections
    cursor.close()
    conn.close()
