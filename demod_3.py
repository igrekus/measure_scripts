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

gen_lo = rm.open_resource(instruments['P LO'])
gen_rf = rm.open_resource(instruments['P RF'])
src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['Анализатор'])

file_name = f'xlsx/divisor-1-coeff_{coeff}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'

# TODO compression point calc: decrease scale before point, increase after

def measure_1():
    gen_lo.write('*RST')
    gen_rf.write('*RST')
    src.write('*RST')
    mult.write('*RST')
    sa.write('*RST')

    secondary = {'Plo_min': -10.0, 'Plo_max': 10.0, 'Plo_delta': 1.0, 'Flo_min': 1.0, 'Flo_max': 3.0,
                 'Flo_delta': 0.1, 'Prf': -10.0, 'Frf_min': 1.0, 'Frf_max': 3.0, 'Frf_delta': 0.1, 'Usrc': 5.0,
                 'OscAvg': False, 'Loss': 0.82}

    src_u = 5
    src_i = 200   # mA
    pow_lo_start = -5
    pow_lo_end = -5
    pow_lo_step = 5
    freq_lo_start = 1.0   # GHz
    freq_lo_end = 3.0
    freq_lo_step = 0.5

    pow_rf_start = -20
    pow_rf_end = 6
    pow_rf_step = 2
    freq_rf_start = 0.06
    freq_rf_end = 3.06
    freq_rf_step = 0.5

    pow_lo = -5
    pow_rf = -5

    Pbal = 0.82
    ref_level = 10
    scale_y = 5

    pow_rf_values = [round(x, 3) for x in np.arange(start=pow_rf_start, stop=pow_rf_end + 0.2, step=pow_rf_step)]
    freq_lo_values = [round(x, 3) for x in np.arange(start=freq_lo_start, stop=freq_lo_end + 0.2, step=freq_lo_step)]
    freq_rf_deltas = [x / 1_000 for x in [5, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 150, 200, 250, 300, 350, 400, 450]]

    src.write(f'APPLY p6v,{src_u}V,{src_i}mA')

    sa.write(':CAL:AUTO OFF')
    sa.write(':SENS:FREQ:SPAN 1MHz')
    sa.write(f'DISP:WIND:TRAC:Y:RLEV {ref_level}')
    sa.write(f'DISP:WIND:TRAC:Y:PDIV {scale_y}')

    gen_lo.write(f'SOUR:POW {pow_lo}dbm')
    gen_rf.write(f'SOUR:POW {pow_rf}dbm')

    res = []
    for freq_lo in freq_lo_values:
        gen_lo.write(f'SOUR:FREQ {freq_lo}GHz')

        for freq_rf_delta in freq_rf_deltas:
            freq_rf = freq_lo + freq_rf_delta
            gen_rf.write(f'SOUR:FREQ {freq_rf}GHz')

            src.write('OUTPut ON')

            gen_lo.write(f'OUTP:STAT ON')
            gen_rf.write(f'OUTP:STAT ON')

            time.sleep(0.5)

            # TODO send to ui
            u_mul_read = float(mult.query('MEAS:VOLT?'))
            i_mul_read = float(mult.query('MEAS:CURR?'))

            center_freq = freq_rf - freq_lo
            sa.write(':CALC:MARK1:MODE POS')
            sa.write(f':SENSe:FREQuency:CENTer {center_freq}GHz')
            sa.write(f':CALCulate:MARKer1:X:CENTer {center_freq}GHz')

            time.sleep(0.5)

            pow_read = float(sa.query(':CALCulate:MARKer:Y?'))

            raw_point = {
                'f_lo': freq_lo,
                'f_rf': freq_lo + freq_rf_delta,
                'p_lo': pow_lo,
                'p_rf': pow_rf,
                'u_mul': u_mul_read,
                'i_mul': i_mul_read,
                'pow_read': pow_read,
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
