import pyodbc

# Define the connection string
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=.\SQLSERVERPY;"  # Local server instance
    "DATABASE=IME_WebScraping;"  # Replace with the name of your database
    "Trusted_Connection=yes;"  # Use Windows Authentication
)
print(pyodbc.drivers())
