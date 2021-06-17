from time import sleep
import numpy as np
from datetime import datetime as dt
from pyModbusTCP.client import ModbusClient
import pandas as pd
import os
from openpyxl import load_workbook
from Read_exel import Satec


def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None, truncate_sheet=False, **to_excel_kwargs):
    """
    Append a DataFrame [df] to existing Excel file [filename]
    into [sheet_name] Sheet.
    If [filename] doesn't exist, then this function will create it.

    @param filename: File path or existing ExcelWriter
                     (Example: '/path/to/file.xlsx')
    @param df: DataFrame to save to workbook
    @param sheet_name: Name of sheet which will contain DataFrame.
                       (default: 'Sheet1')
    @param startrow: upper left cell row to dump data frame.
                     Per default (startrow=None) calculate the last row
                     in the existing DF and write to the next row...
    @param truncate_sheet: truncate (remove and recreate) [sheet_name]
                           before writing DataFrame to Excel file
    @param to_excel_kwargs: arguments which will be passed to `DataFrame.to_excel()`
                            [can be a dictionary]
    @return: None

    Usage examples:

    # >>> append_df_to_excel('d:/temp/test.xlsx', df)
    #
    # >>> append_df_to_excel('d:/temp/test.xlsx', df, header=None, index=False)
    #
    # >>> append_df_to_excel('d:/temp/test.xlsx', df, sheet_name='Sheet2',
    #                        index=False)
    #
    # >>> append_df_to_excel('d:/temp/test.xlsx', df, sheet_name='Sheet2',
    #                        index=False, startrow=25)

    (c) [MaxU](https://stackoverflow.com/users/5741205/maxu?tab=profile)
    """
    # Excel file doesn't exist - saving and exiting
    if not os.path.isfile(filename):
        df.to_excel(
            filename,
            sheet_name=sheet_name,
            startrow=startrow if startrow is not None else 0,
            **to_excel_kwargs)
        return

    # ignore [engine] parameter if it was passed
    if 'engine' in to_excel_kwargs:
        to_excel_kwargs.pop('engine')

    writer = pd.ExcelWriter(filename, engine='openpyxl', mode='a')

    # try to open an existing workbook
    writer.book = load_workbook(filename)

    # get the last row in the existing Excel sheet
    # if it was not specified explicitly
    if startrow is None and sheet_name in writer.book.sheetnames:
        startrow = writer.book[sheet_name].max_row

    # truncate sheet
    if truncate_sheet and sheet_name in writer.book.sheetnames:
        # index of [sheet_name] sheet
        idx = writer.book.sheetnames.index(sheet_name)
        # remove [sheet_name]
        writer.book.remove(writer.book.worksheets[idx])
        # create an empty sheet [sheet_name] using old index
        writer.book.create_sheet(sheet_name, idx)

    # copy existing sheets
    writer.sheets = {ws.title: ws for ws in writer.book.worksheets}

    if startrow is None:
        startrow = 0

    # write out the new sheet
    df.to_excel(writer, sheet_name, startrow=startrow, **to_excel_kwargs)

    # save the workbook
    writer.save()


i = 0
file_line = 0  # Строка списка
file_col = 2  # Столбец списка
# Читаем файл с параметрами
read_params = pd.read_excel('satec.xlsx')
# print('Количество строк:', read_params.shape[0])  # Считаем количество строк в файле
# print(type(read_params.shape[0]))
tot_line = read_params.shape[0]

df = pd.DataFrame({'Desc': [1], 'Name': [1], 'ip': [1], 'FW': [1], 'port': [1], 'id': [1], 'uab': [1], 'ubc': [1],
                   'uca': [1], 'angle_uab': [1], 'angle_ubc': [1], 'angle_uca': [1], 'la': [1], 'lb': [1], 'lc': [1],
                   'Angle_la': [1], 'Angle_lb': [1], 'Angle_lc': [1], 'Time': [1]})
df.to_excel('./satec_out.xlsx')


while i < tot_line:
    # Читаем данные из ячеек для соединения
    ip = read_params.iloc[file_line, 2]  # 1- первая строка от шапки 2 столбца
    port = read_params.iloc[file_line, 3]
    unit = read_params.iloc[file_line, 4]
    desc = read_params.iloc[file_line, 0]  # Название счетчика
    satec = Satec(port, unit, ip, desc)
    print('Desc ', desc)
    print(ip)
    out_file = os.getcwd() + '/satec_out.xlsx'
    # alpha = Alpha(port, unit, ip)

    i += 1
    file_line += 1
    # Тип счетчика

    df = pd.DataFrame(satec.out_to_excel, index=[1])
    append_df_to_excel(out_file, df, header=None)
