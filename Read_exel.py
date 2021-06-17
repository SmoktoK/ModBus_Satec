import pandas as pd
import numpy as np
from pyModbusTCP.client import ModbusClient
from datetime import datetime as dt

read_params = pd.read_excel('satec.xlsx')  # Файл для считывания
print('Количество строк:', read_params.shape[0])  # Считаем количество строк в файле


class Satec:
    def __init__(self, port, unit, ip, desc):
        self.c = ModbusClient()
        self.c.host(ip)
        self.c.port(port)
        self.c.unit_id(unit)
        self.c.open()
        self.device_type = self.uint32_reg(46082, 'type')  # Тип счетчика
        print(f'{self.device_type=}')
        self.device_sn = self.uint32_reg(46080, 'sn')  # Серийный номер
        self.device_fw = self.uint16_reg(46100, 'fw', reg_count=4)
        self.desc = desc
        self.ip = ip
        self.port = port
        self.unit = unit
        if self.device_type == 72000:
            self.numbers = self.get_number720()
        elif self.device_type == 13330:
            self.numbers = self.get_number133()
        elif self.device_type == 175:
            self.numbers = self.get_number175()
        elif self.device_type == 18000:
            self.numbers = self.get_number180()
        else:
            print('Type error')
        self.c.close()

    def get_number720(self):
        return {
            'uab': self.uint32_reg(13888, 'uab'),
            'ubc': self.uint32_reg(13890, 'ubc'),
            'uca': self.uint32_reg(13892, 'uca'),
            'angle_uab': self.int32_reg(13904, 'angle_uab'),
            'angle_ubc': self.int32_reg(13906, 'angle_ubc'),
            'angle_uca': self.int32_reg(13908, 'angle_uca'),
            'la': self.uint32_reg(13896, 'la'),
            'lb': self.uint32_reg(13898, 'lb'),
            'lc': self.uint32_reg(13900, 'lc'),
            'angle_la': self.int32_reg(13912, 'angle_la'),
            'angle_lb': self.int32_reg(13914, 'angle_lb'),
            'angle_lc': self.int32_reg(13916, 'angle_lc'),
            'Time': self.uint16_time(46416, 'Time')
        }

    def get_number133(self):
        return {
            'uab': self.uint32_reg(13864, 'uab'),
            'ubc': self.uint32_reg(13866, 'ubc'),
            'uca': self.uint32_reg(13888, 'uca'),
            'angle_uab': self.int32_reg(13880, 'angle_uab'),
            'angle_ubc': self.int32_reg(13882, 'angle_ubc'),
            'angle_uca': self.int32_reg(13884, 'angle_uca'),
            'la': self.uint32_reg(13872, 'la'),
            'lb': self.uint32_reg(13874, 'lb'),
            'lc': self.uint32_reg(13876, 'lc'),
            'angle_la': self.int32_reg(13888, 'angle_la'),
            'angle_lb': self.int32_reg(13890, 'angle_lb'),
            'angle_lc': self.int32_reg(13892, 'angle_lc'),
        }

    def get_number175(self):
        return {
            'uab': self.uint32_reg(13888, 'uab'),
            'ubc': self.uint32_reg(13890, 'ubc'),
            'uca': self.uint32_reg(13892, 'uca'),
            'angle_uab': self.int32_reg(13904, 'angle_uab'),
            'angle_ubc': self.int32_reg(13906, 'angle_ubc'),
            'angle_uca': self.int32_reg(13908, 'angle_uca'),
            'la': self.uint32_reg(13896, 'la'),
            'lb': self.uint32_reg(13898, 'lb'),
            'lc': self.uint32_reg(13900, 'lc'),
            'angle_la': self.int32_reg(13912, 'angle_la'),
            'angle_lb': self.int32_reg(13914, 'angle_lb'),
            'angle_lc': self.int32_reg(13916, 'angle_lc'),
        }

    def get_number180(self):
        return {
            'uab': self.uint32_reg(13888, 'uab'),
            'ubc': self.uint32_reg(13890, 'ubc'),
            'uca': self.uint32_reg(13892, 'uca'),
            'angle_uab': self.int32_reg(13904, 'angle_uab'),
            'angle_ubc': self.int32_reg(13906, 'angle_ubc'),
            'angle_uca': self.int32_reg(13908, 'angle_uca'),
            'la': self.uint32_reg(13896, 'la'),
            'lb': self.uint32_reg(13898, 'lb'),
            'lc': self.uint32_reg(13900, 'lc'),
            'angle_la': self.int32_reg(13912, 'angle_la'),
            'angle_lb': self.int32_reg(13914, 'angle_lb'),
            'angle_lc': self.int32_reg(13916, 'angle_lc'),
        }

    def uint32_reg(self, reg, default, reg_count=2):

        uab_reg = self.c.read_holding_registers(reg, reg_count)
        if uab_reg:
            uab = (np.uint16(uab_reg[1]) << 16) + np.uint16(uab_reg[0])
            print(default, uab)
        else:
            return f'{default} error'

        return uab

    def int32_reg(self, reg, default, reg_count=2):

        uab_reg = self.c.read_holding_registers(reg, reg_count)
        if uab_reg:
            uab = (np.int16(uab_reg[1]) << 16) + np.uint16(uab_reg[0])
            print(default, uab)
        else:
            return f'{default} error'

        return uab

    def uint16_reg(self, reg, default, reg_count=2):

        uab_reg = self.c.read_holding_registers(reg, reg_count)
        if uab_reg:
            uab = str(np.uint16(uab_reg[0])) + str(np.uint16(uab_reg[1]))
            print(default, uab)
        else:
            return f'{default} error'

        return uab

    def uint16_time(self, reg, default, reg_count=2):

        time_reg = self.c.read_holding_registers(reg, reg_count)
        if time_reg:
            ts = (np.uint16(time_reg[1]) << 16) + np.uint16(time_reg[0])
            ts_out = dt.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')
            # print(np.uint16(time_reg[10]), np.uint16(time_reg[9]), np.uint16(time_reg[8]), np.uint16(time_reg[7]),
            #       np.uint16(time_reg[6]), np.uint16(time_reg[5]))
            print(default, ts_out)
        else:
            return f'{default} error'

        return ts

    @property
    def out_to_excel(self):
        res = {
            'Desc': self.desc,
            'Name': str(self.device_type)[:3],
            'ip': self.ip,
            'FW': self.device_fw,
            'port': self.port,
            'id': self.unit,
        }
        res.update(self.numbers)

        return res
