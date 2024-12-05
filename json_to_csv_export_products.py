import json
import os
import pandas as pd

# Directory containing your JSON files
directory = 'G:\Education\DataSource\Data\IME\export'

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
print(df.columns)

# Rename columns in the DataFrame to match SQL table columns
df.rename(columns={
    'نام کالا': 'commodity_name',
    'نماد': 'symbol',
    'تالار': 'hall',
    'تولید کننده': 'producer',
    'نوع قرارداد': 'contract_type',
    'قیمت پایانی میانگین   موزون': 'avg_closing_price',
    'ارزش معامله': 'transaction_value',
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
    'کد عرضه': 'offer_code',
    'نرخ ارز روزانه تسویه': 'currency_day',

}, inplace=True)

df.to_csv(f"{directory}/result.csv", index=False)