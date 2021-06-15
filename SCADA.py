import numpy as np
from datetime import datetime as dt
from pyModbusTCP.client import ModbusClient
import pandas as pd

file_line = 1
file_col = 2
# Читаем файл с параметрами
read_params = pd.read_excel('satec.xlsx')
print('Количество строк:', read_params.shape[0])

# Читаем данные из ячеек для соединения
ip = read_params.iloc[file_line, file_col]  # 1- первая строка от шапки 2 столбца
port = read_params.iloc[1, 3]
unit = read_params.iloc[1, 4]
c = ModbusClient()
c.host(ip)
c.port(port)
c.unit_id(unit)
c.open()
print(ip)

# Тип счетчика

device_type_reg = c.read_holding_registers(46082, 2)
if device_type_reg:
    device_type = str((np.uint16(device_type_reg[1]) << 16) + np.uint16(device_type_reg[0]))
    print('Type: ', device_type[:3])
else:
    print("read error")

# Серийный номер

device_sn_reg = c.read_holding_registers(46080, 2)
if device_sn_reg:
    device_sn = str((np.uint16(device_sn_reg[1]) << 16) + np.uint16(device_sn_reg[0]))
    print('S/N ', device_sn)
else:
    print("read error")

# Прошивка

device_fw_reg = c.read_holding_registers(46100, 4)
if device_fw_reg:
    device_fw = str(np.uint16(device_fw_reg[0])) + str(np.uint16(device_fw_reg[1]))
    print('firmware ', device_fw)
else:
    print("read error")

# Прошивка модуля связи (ПРОВЕРИТЬ!!!!!)

# deviceEht_fw_reg = c.read_holding_registers(46100, 4)
# if deviceEht_fw_reg:
#     deviceEth_fw = str(np.uint16(deviceEht_fw_reg[2])) + str(np.uint16(deviceEht_fw_reg[3]))
#     print('firmware Eth ', deviceEth_fw)
# else:
#     print("read error")

# Время на приборе (ПРОВЕРИТЬ!!!!!)

# time_reg = c.read_holding_registers(46416, 12)
# print(time_reg)
# ts = (np.uint16(time_reg[3]) << 16) + np.uint16(time_reg[2])
# print("Time - %s" % (dt.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')))


# Uab 720

uab_reg_720 = c.read_holding_registers(13888, 2)
if uab_reg_720:
    uab = (np.uint16(uab_reg_720[1]) << 16) + np.uint16(uab_reg_720[0])
    print('Uab ', uab)
else:
    print("read error")

# Ubc 720

ubc_reg_720 = c.read_holding_registers(13890, 2)
if ubc_reg_720:
    ubc = (np.uint16(ubc_reg_720[1]) << 16) + np.uint16(ubc_reg_720[0])
    print('Ubc ', ubc)
else:
    print("read error")

# Uca 720

uca_reg_720 = c.read_holding_registers(13892, 2)
if uca_reg_720:
    uca = (np.uint16(uca_reg_720[1]) << 16) + np.uint16(uca_reg_720[0])
    print('Uca ', uca)
else:
    print("read error")


# Angle_Uab 720

angle_uab_reg_720 = c.read_holding_registers(13904, 2)
if angle_uab_reg_720:
    angle_uab = (np.int16(angle_uab_reg_720[1]) << 16) + np.uint16(angle_uab_reg_720[0])
    print('Angle_Uab ', angle_uab)
else:
    print("read error")

# Angle_Ubc 720

angle_ubc_reg_720 = c.read_holding_registers(13906, 2)
if angle_uab_reg_720:
    angle_ubc = (np.int16(angle_ubc_reg_720[1]) << 16) + np.uint16(angle_ubc_reg_720[0])
    print('Angle_Ubc ', angle_ubc)
else:
    print("read error")

# Angle_Uca 720

angle_uca_reg_720 = c.read_holding_registers(13908, 2)
if angle_uab_reg_720:
    angle_uca = (np.int16(angle_uca_reg_720[1]) << 16) + np.uint16(angle_uca_reg_720[0])
    print('Angle_Uca ', angle_uca)
else:
    print("read error")


# La 720

la_reg_720 = c.read_holding_registers(13896, 2)
if la_reg_720:
    la = (np.uint16(la_reg_720[1]) << 16) + np.uint16(la_reg_720[0])
    print('La ', la)
else:
    print("read error")

# Lb 720

lb_reg_720 = c.read_holding_registers(13898, 2)
if lb_reg_720:
    lb = (np.uint16(lb_reg_720[1]) << 16) + np.uint16(lb_reg_720[0])
    print('Lb ', lb)
else:
    print("read error")

# Lc 720

lc_reg_720 = c.read_holding_registers(13900, 2)
if lc_reg_720:
    lc = (np.uint16(lc_reg_720[1]) << 16) + np.uint16(lc_reg_720[0])
    print('Lc ', lc)
else:
    print("read error")
print(uab_reg_720)
# regs = c.read_holding_registers(36100, 4)
# if regs:
#     print("elements num %d" % (len(regs)))
#     print("element's type %s" % (type(regs[0])))
#     print(regs)
#
#     kvar = (np.int16(regs[1]) << 16) + np.uint16(regs[0])
#     ts = (np.uint16(regs[3]) << 16) + np.uint16(regs[2])
#
#     print("example INT32 - kvar - %d" % (kvar))
#     print("example UINT32 - timestamp - %s" % (dt.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')))
#
# else:
#     print("read error")

c.close()
