from traceback import print_tb

test = """commodity_name, symbol, hall, producer, contract_type, avg_closing_price,
        transaction_value, closing_price, lowest_price, highest_price, base_price,
        supply_volume, demand_volume, contract_volume, unit, transaction_date, delivery_date,
        delivery_location, seller, settlement_price_date, broker, offer_method, purchase_method,
        packaging_type, settlement_type, currency_type, offer_code"""

test_list = test.split(',')
print(len(test_list))