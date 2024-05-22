import openpyxl, argparse, time, traceback, datetime, shutil
# import pandas as pd
# import numpy as np
# from openpyxl.worksheet.formula import ArrayFormula
# from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


# xls = pd.ExcelFile('Book1.xlsx')
# read_pd = pd.read_excel(xls, engine='openpyxl')
# df = pd.DataFrame(read_pd)

# df['Birthday'] = pd.to_datetime(df['Birthday'], format='%Y-%m-%d')


# filtered_df = df.loc[(df['Birthday'] < '2023-11-06')]

# # startDate = '2022-02-11'
# # endDate = '2023-11-06'
# # filtered_df = df.loc[(df['Birthday'] >= startDate)
# #                      & (df['Birthday'] <= endDate)]

# print(filtered_df)


print(shutil.which('filter_sorter'))