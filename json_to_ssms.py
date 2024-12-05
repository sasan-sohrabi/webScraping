import json
import os
import pandas as pd
import pyodbc

# Directory containing your JSON files
directory = 'G:\\Education\\DataSource\\Data\\IME'

# Establish a connection to SQL Server
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=.\SASANENERGY;"  
    "DATABASE=IME_WebScraping;"  
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()

# Loop through all files in the directory and aggregate data
all_data = []

for filename in os.listdir(directory):
    if filename.endswith('.json'):
        file_path = os.path.join(directory, filename)

        # Read and load the JSON file
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)

            # Add data to the list
            if len(data['data']) > 0:
                all_data.extend(data['data'])
            else:
                print("There is no data!", filename, len(data['data']))

# Convert the aggregated data to a pandas DataFrame
df = pd.DataFrame(all_data)

# Clean up any empty strings in the data (Replace "" with None)
df = df.applymap(lambda x: None if x == '' else x)

# Rename columns in the DataFrame to match SQL table columns
df.rename(columns={
    'نام کالا': 'commodity_name',
    'نماد': 'symbol',
    'تالار': 'hall',
    'تولید کننده': 'producer',
    'نوع قرارداد': 'contract_type',
    'قیمت پایانی میانگین   موزون': 'avg_closing_price',
    'ارزش معامله (هزارریال)': 'transaction_value',
    'قیمت پایانی': 'closing_price',
    'پایین ترین': 'lowest_price',
    'بالاترین': 'highest_price',
    'قیمت پایه عرضه': 'base_price',
    'حجم عرضه': 'supply_volume',
    'تقاضا': 'demand_volume',
    'حجم قرارداد': 'contract_volume',
    'واحد': 'unit',
    'تاریخ معامله': 'transaction_date',
    'تاریخ تحویل': 'delivery_date',
    'مکان تحویل': 'delivery_location',
    'عرضه کننده': 'seller',
    'تاریخ قیمت تسویه': 'settlement_price_date',
    'کارگزار': 'broker',
    'نحوه عرضه': 'offer_method',
    'روش خرید': 'purchase_method',
    'نوع بسته بندی': 'packaging_type',
    'نوع تسویه': 'settlement_type',
    'نوع ارز': 'currency_type',
    'کد عرضه': 'offer_code'
}, inplace=True)

# Create the table with NVARCHAR(MAX) for all columns
table_name = "IME_Transactions"
create_table_query = f"CREATE TABLE {table_name} (ID INT IDENTITY(1,1) PRIMARY KEY, "

# Add columns dynamically, all as NVARCHAR(MAX)
for column in df.columns:
    create_table_query += f"[{column}] NVARCHAR(MAX), "
create_table_query = create_table_query.rstrip(", ") + ");"

# Execute table creation
try:
    cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
    cursor.execute(create_table_query)
    conn.commit()
    print(f"Table '{table_name}' created successfully.")
except Exception as e:
    print("Error creating table:", e)

# Adjust the SQL Insert statement to match the number of columns (27 columns)
insert_query = f"""
    INSERT INTO {table_name} (
        commodity_name, symbol, hall, producer, contract_type, avg_closing_price, 
        transaction_value, closing_price, lowest_price, highest_price, base_price, 
        supply_volume, demand_volume, contract_volume, unit, transaction_date, delivery_date, 
        delivery_location, seller, settlement_price_date, broker, offer_method, purchase_method, 
        packaging_type, settlement_type, currency_type, offer_code
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

# Insert each row of data, ensuring all values are converted to strings
for index, row in df.iterrows():
    # Convert all values to strings, ensuring None remains as NULL in SQL
    row = tuple(str(x) if x is not None else None for x in row)
    cursor.execute(insert_query, row)

# Commit the transaction
conn.commit()

# Close the connection
cursor.close()
conn.close()

print("Data inserted successfully!")