import pandas as pd

pricing = pd.read_excel('excel_sheets/Pricing.xlsx')
pricing.set_index('NDC', inplace=True)
pricing_json_data = pricing.to_json(orient='index')
with open('../json_files/pricing.json', 'w') as f:
    f.write(pricing_json_data)

thirty_six_months = pd.read_excel('excel_sheets/36Months.xlsx')
thirty_six_months.set_index('ITEM_CODE', inplace=True)
json_data = thirty_six_months.to_json(orient='index')
with open('../json_files/36Months.json', 'w') as f:
    f.write(json_data)

path = '../excel_sheets/Configurations.xlsx'
configurations = pd.read_excel(path)
cols = ['Material (10 Digit)', 'Description', 'Inner Case Qty (ea)', 'Outer Case Qty (ea)']
configurations = configurations[cols]
configurations.set_index('Material (10 Digit)', inplace=True)
configuration_json_data = configurations.to_json(orient='index')
with open('../json_files/configurations.json', 'w') as f:
    f.write(configuration_json_data)