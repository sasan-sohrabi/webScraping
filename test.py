import pandas as pd
from sqlalchemy import create_engine
import pyodbc

# Define the connection string
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=.\sasanpy;"  # Local server instance
    "DATABASE=PlaceTemp;"  # Replace with the name of your database
    "Trusted_Connection=yes;"  # Use Windows Authentication
)

# Establish connection
try:

    # Query data
    query = "SELECT * FROM dbo.Place"
    df = pd.read_sql(query, conn)

    # Display the data
    row_count = len(df)
    print(f"Number of rows: {row_count}")

    # Close the connection
    conn.close()
except Exception as e:
    print(f"Error: {e}")
