server_name='EXLBRD-V-SQL1'
db_name='FabManager'

'''
Database Id: EXLBRD-V-SQL1.FabManager

Server Name: EXLBRD-V-SQL1
DataBase: FabManager

Tables:

> FabManager.Tables

Spools: dbo.tbllISOSpools (Same data as 'SpoolData') Seems to also have Statuses
Also has dbo.tblSpools with very similar informaation

Welds:

'''

import pyodbc
print(pyodbc.drivers())

import pyodbc

conn_str = (
   f"DRIVER={{ODBC Driver 17 for SQL Server}};"
   f"SERVER={server_name};"
   f"Trusted_Connection=yes;"
   f"DATABASE={db_name};"
   f"Encrypt=yes;"
   f"TrustServerCertificate=yes"
)
conn = pyodbc.connect(conn_str)

# Test the connection
try:
    cursor = conn.cursor()
    cursor.execute("SELECT @@version")
    version = cursor.fetchone()
    print("Connection successful!")
    print(f"SQL Server version: {version[0]}")
except pyodbc.Error as e:
    print(f"Connection failed: {str(e)}")
finally:
    if 'conn' in locals():
        conn.close()