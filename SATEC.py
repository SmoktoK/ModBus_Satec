import threading, time
from pyModbusTCP.client import ModbusClient
import openpyxl
import random
import numpy as np
import time

t0 = time.clock
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
t1 = time.clock() -t0
print('Список сгенерирован за: ', t1, 'секунд')

print(full_list)


def modbus(host, port, addr, reg, task_list):
    # open or reconnect TCP to server
    c = ModbusClient()
    c.host(host)
    c.port(port)
    time.sleep(random.random())
    if not c.is_open():
        if not c.open():
            print("unable to connect to " + host + ":" + str(port) + str(addr))
            final_output.append([f'unable to connect {host}'])
    # if open() is ok, read register (modbus function 0x03)
    if c.is_open():
        # read 10 registers at address 0, store result in regs list
        regs = c.read_holding_registers(addr, reg)
        # if success display registers
        if regs:
            typelist = list(task_list.values())
            keyslist = list(task_list.keys())
            print("reg ad #0 to 9: " + str(regs))
            final_list = [regs[key] for key in keyslist]
            # print(str(final_list))
            # print(str(keyslist))
            # print(str(typelist))
            q = 0
            final_small_output = [host, addr, reg]
            for type in typelist:
                if type == "UINT16":
                    count = np.uint16(regs[keyslist[q]])
                    # print(count, q)
                    key_final = regs[keyslist[q]]
                    final_small_output.append(f'{keyslist[q]}: {count}')
                    q += 1
                elif type == "INT16":
                    count = np.int16(regs[keyslist[q]])
                    # print(count, q)
                    final_small_output.append(f'{keyslist[q]}: {count}')
                    q += 1
                elif type == "UINT32":
                    count = (np.uint16(regs[keyslist[q] + 1]) << 16) + np.uint16(regs[keyslist[q]])
                    # print(count, q)
                    final_small_output.append(f'{keyslist[q]}: {count}')
                    q += 1
                elif type == "INT32":
                    count = (np.int16(regs[keyslist[q] + 1]) << 16) + np.uint16(regs[keyslist[q]])
                    # print(count, q)
                    final_small_output.append(f'{keyslist[q]}: {count}')
                    q += 1
                else:
                    print("error data TYPE")
                    final_small_output.append([f'Error data TYPE {host}{keyslist[q]}: {keyslist[q]}'])
            final_output.append(final_small_output)
        else:
            final_output.append([f'unable to read register {host} {addr}'])


threads = []

for device in full_list:
        t = threading.Thread(target=modbus, args=[device['server_host'], device['server_port'], device['start_reg'], device['reg_qnty'], device['task_list']])
        t.start()
        threads.append(t)

for thread in threads:
    thread.join()


for answ in final_output:
    sheet_save.append(answ)

book_save.save('result.xlsx')
t2 = time.clock() - t0
print('Время выполнения программы:', t2)
