import threading
import time
from pyModbusTCP.client import ModbusClient
import openpyxl
import random
import numpy as np

t_start = time.perf_counter()



book = openpyxl.open('test.xlsx', read_only=True)
sheet = book.active
book_save = openpyxl.Workbook()
sheet_save = book_save.active
full_list = list()
single_dict = dict()
result_list = list()
temp_result_list = list()
temptemp = []
server_host = ''
server_port = 0
start_reg = 0
reg_qnty = 0
b = 0
count_check = 0
keyslist = []
final_list = []
final_output = []


for rows in range(1, sheet.max_row+1):
    single_dict["server_host"] = sheet[rows][0].value
    single_dict["server_port"] = int(sheet[rows][1].value)
    single_dict["start_reg"] = int(sheet[rows][3].value)
    single_dict["reg_qnty"] = int(sheet[rows][4].value)
    temptemp = sheet[rows][5].value
    res = {int(sub.split(":")[0]): sub.split(":")[1] for sub in temptemp[1:-1].split(", ")}
    single_dict["task_list"] = res
    # single_dict["task_list"] = sheet[rows][5].value
    full_list.append(single_dict.copy())
    # print(temptemp)
    # print(res)

f = open('xyz.txt', 'w')
f.write(str(full_list))
f.close()
print(full_list)
all_time = time.perf_counter() - t_start
print(f'Время генерации списка: {all_time}')