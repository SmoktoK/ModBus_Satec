import pandas as pd
import numpy as np
from pyModbusTCP.client import ModbusClient

read_params = pd.read_excel('satec.xlsx')
print('Количество строк:', read_params.shape[0])

SATEC_PORTS = {'720'}


class MainCounter:
    def __init__(self, port, unit, ip, desc):
        self.c = ModbusClient()
        self.c.host(ip)
        self.c.port(port)
        self.c.unit_id(unit)
        self.c.open()

        self.device_type = self.uab_reg(46082, 2, 'type')
        self.device_sn = self.uab_reg(46080, 2, 'sn')
        self.device_fw = self.uab_reg(46100, 4, 'fw')
        self.desc = desc
        self.ip = ip
        self.port = port
        self.unit = unit

        self.numers = self.get_number()

        self.c.close()

    def get_number(self):
        return {}

    def uab_reg(self, reg, reg_coun, default):

        uab_reg = self.c.read_holding_registers(reg, reg_coun)
        if uab_reg:
            uab = (np.uint16(uab_reg[1]) << 16) + np.uint16(uab_reg[0])
            print(default, uab)
        else:
            return f'{default} error'

        return uab

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
        res.update(self.numers)

        return res


class Satec(MainCounter):

    def get_number(self):
        return {
            'uab': self.uab_reg(13888, 2, 'uab'),
            'ubc': self.uab_reg(13890, 2, 'ubc'),
            'uca': self.uab_reg(13892, 2, 'uca'),
            'angle_uab': self.uab_reg(13904, 2, 'angle_uab'),
            'angle_ubc': self.uab_reg(13906, 2, 'angle_ubc'),
            'angle_uca': self.uab_reg(13908, 2, 'angle_uca'),
            'la': self.uab_reg(13896, 2, 'la'),
            'lb': self.uab_reg(13898, 2, 'lb'),
            'lc': self.uab_reg(13900, 2, 'lc'),
            'angle_la': self.uab_reg(13912, 2, 'angle_la'),
            'angle_lb': self.uab_reg(13914, 2, 'angle_lb'),
            'angle_lc': self.uab_reg(13916, 2, 'angle_lc'),
        }


class Alpha(Satec):

    def get_number(self):
        return {
            'uab': self.uab_reg(13888, 2, 'uab'),
            'ubc': self.uab_reg(13890, 2, 'ubc'),
            'uca': self.uab_reg(13892, 2, 'uca'),
            'angle_uab': self.uab_reg(13904, 2, 'angle_uab'),
            'angle_ubc': self.uab_reg(13906, 2, 'angle_ubc'),
            'angle_uca': self.uab_reg(13908, 2, 'angle_uca'),
            'la': self.uab_reg(13896, 2, 'la'),
            'lb': self.uab_reg(13898, 2, 'lb'),
            'lc': self.uab_reg(13900, 2, 'lc'),
            'angle_la': self.uab_reg(13912, 2, 'angle_la'),
            'angle_lb': self.uab_reg(13914, 2, 'angle_lb'),
            'angle_lc': self.uab_reg(13916, 2, 'angle_lc'),
        }
