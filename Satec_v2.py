import openpyxl
from pyModbusTCP.client import ModbusClient
import time

book = openpyxl.open('test.xlsx', read_only=True)
sheet = book.active
book_save = openpyxl.Workbook()

def modbus(host, port, addr, reg, task_list):
    # open or reconnect TCP to server
    c = ModbusClient()
    c.host(host)
    c.port(port)
    time.sleep(random.random())
    print('connect to: ',  host)
    if not c.is_open():