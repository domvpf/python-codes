import openpyxl, time, pandas as pd, sys, shutil, rfunctions, warnings, xlwings
from datetime import datetime
from xlwings.constants import AutoFillType
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=UserWarning)

def dataTransformer(*params):
    now = datetime.now()
    dt_string = now.strftime("%m-%d-%Y %H-%M-%S")
    log_file = 'dataTransformer-'+dt_string+'-log.log'

    log_string = ""
    output_folder = ""

    config_file = params[0]
    
    config_file_pd = pd.read_excel(config_file, dtype=str)
    config_file_df = pd.DataFrame(config_file_pd).fillna('')
    config_file_list = config_file_df.values.tolist()
    
    formula_keys_pd = pd.read_excel(config_file, 'FORMULA KEYS',dtype=str)
    formula_keys_df = pd.DataFrame(formula_keys_pd).fillna('')
    formula_keys_list = formula_keys_df.values.tolist()
      
    for i in config_file_list:
            if i[0] == 'in_MapperWorkbookFilePath':
                in_MapperWorkbookFilePath = i[1]
            if i[0] == 'in_MapperWorksheetName':
                in_MapperWorksheetName = i[1]
            if i[0] == 'in_InputFilePath':
                in_InputFilePath = i[1]
            if i[0] == 'in_InputWorksheetName':
                in_InputWorksheetName = i[1]
            if i[0] == 'in_OutputFilePath':
                in_OutputFilePath = i[1]
            if i[0] == 'in_OutputWorksheetName':
                in_OutputWorksheetName = i[1]
            if i[0] == 'in_StandardTemplateFilePath':
                in_StandardTemplateFilePath = i[1]
            if i[0] == 'in_StandardTemplateWorksheetName':
                in_StandardTemplateWorksheetName = i[1]

    print('--------------------------CONFIG FILE-------------------------------')
    print('Mapper file: ',in_MapperWorkbookFilePath)
    print('Mapper file sheet name: ',in_MapperWorksheetName)
    print('Input file: ',in_InputFilePath)
    print('Input file sheet name: ',in_InputWorksheetName)
    print('Output file: ',in_OutputFilePath)
    print('Output file sheet name: ',in_OutputWorksheetName)
    print('Standard template file: ',in_StandardTemplateFilePath)
    print('Standard template file sheet name: ',in_StandardTemplateWorksheetName)
    print('--------------------------------------------------------------------')

    #Generate output file, but not yet manipulated
    xls = pd.ExcelFile(in_MapperWorkbookFilePath)
    read_pd = pd.read_excel(xls, engine='openpyxl')
    df = pd.DataFrame(read_pd)
    
    config_openpyxl = openpyxl.load_workbook(in_MapperWorkbookFilePath)
    config_openpyxl_ws = config_openpyxl[in_MapperWorksheetName]

    input_file_book = openpyxl.load_workbook(in_InputFilePath,read_only=False)
    input_file_ws = input_file_book[in_InputWorksheetName]

    app = xlwings.App(visible = False)
    input_xw = xlwings.Book(in_InputFilePath)
    input_xw_sheet = input_xw.sheets(in_InputWorksheetName)

    output_xw = xlwings.Book(in_OutputFilePath)
    output_xw_sheet = output_xw.sheets(in_OutputWorksheetName)

    process_row_start = 0
    print("------ Input Mapper Processing: ------")
    input_mapper_time = time.time()
    while len(df) > process_row_start:
        process_row_start_time = time.time()
        dataframe_mapper = df.iloc[process_row_start]
        processMapperRow(dataframe_mapper, input_file_ws, input_file_book, in_InputFilePath, formula_keys_list, in_OutputFilePath, in_OutputWorksheetName, input_xw, input_xw_sheet, output_xw_sheet, config_openpyxl_ws, process_row_start)
        process_row_start = process_row_start + 1
        print("row: " + str(process_row_start+1) + " " + dataframe_mapper['Retrieve Function'] + ": " + "%s seconds" % (time.time() - process_row_start_time))
    print("Input Mapper: DONE %s seconds" % (time.time() - input_mapper_time))

    print("---------- Saving Output ----------")
    input_xw.save()
    output_xw.save()
    input_xw.app.quit()
    print("Saved")


def processMapperRow(dataframe_mapper, input_file_ws, input_file_book, in_InputFilePath, formula_keys_list, in_OutputFilePath, in_OutputWorksheetName, input_xw, input_xw_sheet, output_xw_sheet, config_openpyxl_ws, process_row_start):
    rfunction_name = dataframe_mapper['Retrieve Function']
    plot_to_where = dataframe_mapper['Plot to Where']
    rfunction_number = rfunction_name.split(' ')[-1]    

    start_time = time.time()
    if rfunction_name:
        file_list = rfunctions.call_retrieve_functions(rfunction_name, dataframe_mapper, input_file_ws, input_xw_sheet, output_xw_sheet, config_openpyxl_ws, process_row_start)


def main(args=None):
    dataTransformer(*sys.argv[1:])

    
if __name__ == '__main__':
	start_time = time.time()
	main()
	print("--- %s seconds ---" % (time.time() - start_time))
