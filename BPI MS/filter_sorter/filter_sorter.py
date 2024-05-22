import openpyxl, argparse, time, traceback, datetime, os
import pandas as pd
import numpy as np
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

def filter_process(input_path, config_path, output_path):
    xls = pd.ExcelFile(input_path)
    read_pd = pd.read_excel(xls, engine='openpyxl')
    df = pd.DataFrame(read_pd)

    # input_file_book = openpyxl.load_workbook(input_path, read_only=False)
    # input_file_ws = input_file_book['Sheet1']

    config = pd.ExcelFile(config_path)
    config_pd = pd.read_excel(config, engine='openpyxl')
    config_df = pd.DataFrame(config_pd)

    config_sc = ['contains', 'startsWith', 'endsWith', 'exact', 'All', 'withValue', 'beforeDate', 'afterDate', 'dateRange']
    config_action = ['Keep', 'Remove', 'Replace', 'paddLeft', 'paddRight', 'addPrefix']

    df_list = df
    for i in range(0, len(config_df)):  
        config_mapper = config_df.iloc[i].astype(str)
        df_list = df_list.astype(str)
        if config_mapper['Search Condition'] == config_sc[0]:
            if config_mapper['Action'] == config_action[0]:
                df_list = df_list[df_list[config_mapper[0]].str.contains(str(config_mapper['Search Input']))]
            if config_mapper['Action'] == config_action[1]:
                df_list = df_list[df_list[config_mapper[0]].str.contains(str(config_mapper['Search Input'])) == False]
            if config_mapper['Action'] == config_action[2]:
                df_list = df_list.replace(config_mapper['Search Input'], config_mapper['Output value'])
            if config_mapper['Action'] == config_action[3]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='left', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[4]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='right', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[5]:
                df_list[config_mapper[0]][df_list[config_mapper[0]].str.contains(str(config_mapper['Search Input']))] = config_mapper['Output value'] + df_list[config_mapper[0]]

        if config_mapper['Search Condition'] == config_sc[1]:
            df_list[config_mapper[0]] = df_list[config_mapper[0]].astype(str)
            if config_mapper['Action'] == config_action[0]:
                df_list = df_list[df_list[config_mapper[0]].str.startswith(config_mapper['Search Input'])]
            if config_mapper['Action'] == config_action[1]:
                df_list = df_list[df_list[config_mapper[0]].str.startswith(config_mapper['Search Input']) == False]
            if config_mapper['Action'] == config_action[2]:
                df_list = df_list.replace(config_mapper['Search Input'], config_mapper['Output value'])
            if config_mapper['Action'] == config_action[3]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='left', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[4]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='right', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[5]:
                df_list[config_mapper[0]][df_list[config_mapper[0]].str.startswith(str(config_mapper['Search Input']))] = config_mapper['Output value'] + df_list[config_mapper[0]]

        if config_mapper['Search Condition'] == config_sc[2]:
            df_list[config_mapper[0]] = df_list[config_mapper[0]].astype(str)
            df_list[config_mapper[0]] = df_list[config_mapper[0]].str.split('.').str[0]      
            if config_mapper['Action'] == config_action[0]:
                df_list = df_list[df_list[config_mapper[0]].str.endswith(config_mapper['Search Input'])]
            if config_mapper['Action'] == config_action[1]:
                df_list = df_list[df_list[config_mapper[0]].str.endswith(config_mapper['Search Input']) == False]
            if config_mapper['Action'] == config_action[2]:
                df_list = df_list.replace(config_mapper['Search Input'], config_mapper['Output value'])
            if config_mapper['Action'] == config_action[3]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='left', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[4]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='right', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[5]:
                df_list[config_mapper[0]][df_list[config_mapper[0]].str.endswith(str(config_mapper['Search Input']))] = config_mapper['Output value'] + df_list[config_mapper[0]]

        if config_mapper['Search Condition'] == config_sc[3]:
            if config_mapper['Action'] == config_action[0]:
                temporary_list = config_mapper['Search Input'] = [config_mapper['Search Input']]
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.split('.').str[0]
                df_list = df_list[df_list[config_mapper[0]].isin(temporary_list)]
            if config_mapper['Action'] == config_action[1]:
                temporary_list = config_mapper['Search Input'] = [config_mapper['Search Input']]
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.split('.').str[0]
                df_list = df_list[df_list[config_mapper[0]].isin(temporary_list) == False]
            if config_mapper['Action'] == config_action[2]:
                if '.' in df_list[config_mapper[0]].values.any():
                    df_list[config_mapper[0]] = df_list[config_mapper[0]].str.split('.').str[0]
                    df_list = df_list.replace(config_mapper['Search Input'], config_mapper['Output value'])
                else:          
                    df_list = df_list.replace(config_mapper['Search Input'], config_mapper['Output value'])
            if config_mapper['Action'] == config_action[3]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='left', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[4]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='right', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[5]:
                temporary_list = config_mapper['Search Input'] = [config_mapper['Search Input']]
                df_list[config_mapper[0]][df_list[config_mapper[0]].isin(temporary_list)] = config_mapper['Output value'] + df_list[config_mapper[0]]
                # df_list[config_mapper[0]][df_list[config_mapper[0]].str.contains(str(config_mapper['Search Input']))] = config_mapper['Output value'] + df_list[config_mapper[0]]

        if config_mapper['Search Condition'] == config_sc[4]:
            if config_mapper['Action'] == config_action[0]:
                df_list = df_list[df_list[config_mapper[0]].str.contains(config_mapper['Search Input'])]
            if config_mapper['Action'] == config_action[1]:
                df_list = df_list[df_list[config_mapper[0]].str.contains(config_mapper['Search Input']) == False]
            if config_mapper['Action'] == config_action[2]:
                df_list = df_list.replace(config_mapper['Search Input'], config_mapper['Output value'])
            if config_mapper['Action'] == config_action[3]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].astype(str)
                config_mapper['Padd Length'] = config_mapper['Padd Length'].split('.')[0]
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.split('.').str[0]
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='left', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[4]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].astype(str)
                config_mapper['Padd Length'] = config_mapper['Padd Length'].split('.')[0]
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.split('.').str[0]
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.pad(width= int(config_mapper['Padd Length']), side='right', fillchar = config_mapper['Add Value'])
            if config_mapper['Action'] == config_action[5]:
                df_list[config_mapper[0]][df_list[config_mapper[0]].str.contains(str(config_mapper['Search Input']))] = config_mapper['Output value'] + df_list[config_mapper[0]]

        if config_mapper['Search Condition'] == config_sc[5]:
            if config_mapper['Action'] == config_action[0]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].replace('nan', np.nan, regex=True)
                df_list = df_list.dropna()
                df_list[config_mapper[0]] = df_list[config_mapper[0]].str.split('.').str[0]
            if config_mapper['Action'] == config_action[1]:
                df_list[config_mapper[0]] = df_list[config_mapper[0]].replace('nan', np.nan, regex=True)
                df_list = df_list[df_list[config_mapper[0]].isna()]

        if config_mapper['Search Condition'] == config_sc[6]:
            if config_mapper['Action'] == config_action[0]:
                df_list = df_list.loc[(df_list[config_mapper[0]] < config_mapper['Search Input'])]
                df_list[config_mapper[0]] = pd.to_datetime(df_list[config_mapper[0]])
                df_list[config_mapper[0]] = df_list[config_mapper[0]].dt.strftime('%m/%d/%Y')
            if config_mapper['Action'] == config_action[1]:
                df_list = df_list.loc[(df_list[config_mapper[0]] < config_mapper['Search Input']) == False]
                df_list[config_mapper[0]] = pd.to_datetime(df_list[config_mapper[0]])
                df_list[config_mapper[0]] = df_list[config_mapper[0]].dt.strftime('%m/%d/%Y')
            if config_mapper['Action'] == config_action[2]:
                temp_list = []
                temp_list.append(df_list[config_mapper[0]].loc[(df_list[config_mapper[0]] < config_mapper['Search Input'])].values)
                for value in temp_list:
                    df_list = df_list.replace(value, pd.to_datetime(config_mapper['Output value']).strftime('%m/%d/%Y'))

        if config_mapper['Search Condition'] == config_sc[7]:
            if config_mapper['Action'] == config_action[0]:
                df_list = df_list.loc[(df_list[config_mapper[0]] > config_mapper['Search Input'])]        
                df_list[config_mapper[0]] = pd.to_datetime(df_list[config_mapper[0]])
                df_list[config_mapper[0]] = df_list[config_mapper[0]].dt.strftime('%m/%d/%Y')
            if config_mapper['Action'] == config_action[1]:
                df_list = df_list.loc[(df_list[config_mapper[0]] > config_mapper['Search Input']) == False]
                df_list[config_mapper[0]] = pd.to_datetime(df_list[config_mapper[0]])
                df_list[config_mapper[0]] = df_list[config_mapper[0]].dt.strftime('%m/%d/%Y')
            if config_mapper['Action'] == config_action[2]:
                temp_list = []
                temp_list.append(df_list[config_mapper[0]].loc[(df_list[config_mapper[0]] > config_mapper['Search Input'])].values)
                for value in temp_list:
                    df_list = df_list.replace(value, pd.to_datetime(config_mapper['Output value']).strftime('%m/%d/%Y'))
                
        if config_mapper['Search Condition'] == config_sc[8]:
            if config_mapper['Action'] == config_action[0]:
                df_list = df_list.loc[(df_list[config_mapper[0]] >= config_mapper['Search Start Date']) & (df_list[config_mapper[0]] <= config_mapper['Search End Date'])]
                df_list[config_mapper[0]] = pd.to_datetime(df_list[config_mapper[0]])
                df_list[config_mapper[0]] = df_list[config_mapper[0]].dt.strftime('%m/%d/%Y')
            if config_mapper['Action'] == config_action[1]:
                df_list = df_list.loc[(df_list[config_mapper[0]] >= config_mapper['Search Start Date']) & (df_list[config_mapper[0]] <= config_mapper['Search End Date']) == False]
                df_list[config_mapper[0]] = pd.to_datetime(df_list[config_mapper[0]])
                df_list[config_mapper[0]] = df_list[config_mapper[0]].dt.strftime('%m/%d/%Y')
            if config_mapper['Action'] == config_action[2]:
                temp_list = []
                temp_list.append(df_list[config_mapper[0]].loc[(df_list[config_mapper[0]] >= config_mapper['Search Start Date']) & (df_list[config_mapper[0]] <= config_mapper['Search End Date'])].values)
                for value in temp_list:
                    df_list = df_list.replace(value, pd.to_datetime(config_mapper['Output value']).strftime('%m/%d/%Y'))

        df_list = df_list.replace(r'\bn\b', np.nan, regex=True)
        df_list = df_list.replace('nan', np.nan, regex=True)
        
        continue
    
    for v in df_list:
        if df_list[v].isnull().values.all():
            df_list[v] = df_list[v]
        else:
            df_list[v] = df_list[v].str.split('.').str[0]
    
    final_df = pd.DataFrame(df_list)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    final_df.to_excel(writer, index = False)
    writer.save()
    
def main(args=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("input_path")
    parser.add_argument("config_path")
    parser.add_argument("output_path")
    
    args =  parser.parse_args()

    cwd = os.getcwd()
    dt_string = datetime.datetime.now().strftime("%m-%d-%Y %H-%M-%S")
    log_file = cwd + '\\' + dt_string + ' - ' + 'log_file.log'
    print(cwd)
    try:
        filter_process(args.input_path, args.config_path, args.output_path)
    except Exception as exception:
        traceback.print_exception(
            type(exception),
            exception,
            exception.__traceback__
        )
        with open(log_file, "w") as file:
            file.write(f"Error in file: {exception}")

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))