import datetime
import openpyxl
import time
import visa

import numpy as np
import pandas as pd

from openpyxl.chart import LineChart, Reference

instruments = {
    'Осциллограф': 'GPIB1::7::INSTR',
    'P LO': 'GPIB1::6::INSTR',
    'P RF': 'GPIB1::20::INSTR',
    'Источник': 'GPIB1::3::INSTR',
    'Мультиметр': 'GPIB1::2::INSTR',
    'Анализатор': 'GPIB1::18::INSTR',
}

# --- параметры измерения ---
# генератор
f_start = 0.45   # начальная частота диапазона (ГГц)
f_end = 15.0   # конечная частота диапазона (ГГЦ)
f_step = 0.1   # шаг перестройки частоты (ГГц)
p_in = 0   # мощность дБм

# исочник питания
u_source = 5.0   # напряжение питания (В)
i_source = 100   # макс ток потребления (мА)

# коэфициент деления
coeff = 2
# ---------------------------

rm = visa.ResourceManager()

src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])

file_name = f'xlsx/divisor-1-coeff_{coeff}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'

# TODO compression point calc: decrease scale before point, increase after

def measure_1():
    src.write('*RST')
    mult.write('*RST')

    secondary = {'Plo_min': -10.0, 'Plo_max': 10.0, 'Plo_delta': 1.0, 'Flo_min': 1.0, 'Flo_max': 3.0,
                 'Flo_delta': 0.1, 'Prf': -10.0, 'Frf_min': 1.0, 'Frf_max': 3.0, 'Frf_delta': 0.1, 'Usrc': 5.0,
                 'OscAvg': False, 'Loss': 0.82}

    src_i = 200   # mA
    u_start = 4.75
    u_end = 5.25
    u_step = 0.05

    u_values = [round(x, 3) for x in np.arange(start=u_start, stop=u_end + 0.002, step=u_step)]

    res = []

    # mult.write(f'sens:CURRent:AC:RANGe:AUTO')

    for u in u_values:
        src.write(f'APPLY p6v,{u}V,{src_i}mA')
        src.write('OUTPut ON')

        time.sleep(0.5)

        # u_mul_read = float(mult.query('MEAS:VOLT?'))
        i_mul_read = float(mult.query('MEAS:CURR:DC? 1A,DEF'))

        raw_point = {
            'u_mul': u,
            # 'u_mul': u_mul_read,
            'i_mul': i_mul_read,
        }

        print(raw_point)

        res.append(raw_point)

    with open('out.txt', mode='wt', encoding='utf-8') as f:
        f.write(str(res))

    return res


if __name__ == '__main__':
    measure_1()


"""

"""
