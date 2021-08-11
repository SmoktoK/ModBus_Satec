import pandas as pd
import time
import threading
import random
import numpy as np
from pyModbusTCP.client import ModbusClient
import openpyxl

book_save = openpyxl.Workbook()
sheet_save = book_save.active
final_list = []
final_output = []
# sheet_save = []
t_start = time.perf_counter()

excel_data_df = pd.read_excel('satec.xlsx', usecols=['name', 'server_host', 'server_port', 'start_reg', 'reg_qnty',
                                                     'task_list'])
my_dict = excel_data_df.to_dict()


len_dict = len(my_dict['server_host'])


# print((temptemp)[0])

# res = {int(sub.split(":")[0]): sub.split(":")[1] for sub in temptemp[1:-1].split(", ")}
# single_dict["task_list"] = res
# print('Excel Sheet to Dict:', excel_data_df.to_dict(orient='record'))
# f = open('xyz1.txt', 'w')
# f.write(str(excel_data_df.to_dict()))
# f.close()
# print(my_dict['server_host'][896]) #  обращение по ключу в внутреннему словарю

# print(len_dict)
# t_end = time.perf_counter() - t_start
# print(t_end)

print('Опросный лист сформирован.')
def modbus(name, host, port, addr, reg, task_list):
    # open or reconnect TCP to server
    c = ModbusClient()
    c.host(host)
    c.port(port)
    # c.unit_id(addr)
    # time.sleep(random.randint(3, 10))
    if not c.is_open():
        if not c.open():
            # print("unable to connect to " + host + ":" + str(port) + str(addr))
            final_output.append([name, host, 'unable to connect'])
    # if open() is ok, read register (modbus function 0x03)
    if c.is_open():
        # read 10 registers at address 0, store result in regs list
        regs = c.read_holding_registers(addr, reg)
        # if success display registers
        if regs:
            # print(regs)
            typelist = list(task_list.values())
            keyslist = list(task_list.keys())
            keyslist = [int(x) for x in keyslist]
            # print("reg ad #0 to 9: " + str(regs))
            # final_list = [regs[key] for key in keyslist]
            # print(str(final_list))
            q = 0
            final_small_output = [name, host, addr, reg]
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
                    # final_small_output.append([f'Error data TYPE {host}{keyslist[q]}: {keyslist[q]}'])
                    final_output.append([name, host, 'Error data TYPE'])
            final_output.append(final_small_output)
        else:
            final_output.append([name, host, 'unable to read register'])

#
threads = []
#
for device in range(len_dict):
    temptemp = (my_dict['task_list'][device])
    dicktator = dict(e.split(':') for e in temptemp.split(', '))
    sever_name = my_dict['name'][device]
    server_host = my_dict['server_host'][device]
    server_port = my_dict['server_port'][device]
    start_reg = my_dict['start_reg'][device]
    reg_gnty = my_dict['reg_qnty'][device]
    t = threading.Thread(target=modbus, args=[sever_name, server_host, server_port, start_reg, reg_gnty, dicktator])
    t.start()
    threads.append(t)
    device += 1
    
for thread in threads:
    thread.join()

for answ in final_output:
    sheet_save.append(answ)

book_save.save('result_out.xlsx')
all_time = time.perf_counter() - t_start
print(f'Время выполнения опроса: {all_time}')
print('End!')
