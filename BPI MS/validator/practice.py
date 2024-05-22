import pandas as pd

input_path = 'test_validator_new_1211.xlsx'
config_condition = 'test_validator'

if config_condition in input_path:
    xls = pd.ExcelFile(input_path)
    read_pd = pd.read_excel(xls, engine='openpyxl', sheet_name='SMS')
    df = pd.DataFrame(read_pd)

    print(df)