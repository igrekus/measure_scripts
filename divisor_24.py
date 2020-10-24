import datetime
import time
import visa

import pandas as pd
import numpy as np

from config import instruments

# --- параметры измерения ---
# генератор
f_start = 0.45   # начальная частота диапазона (ГГц)
f_end = 15.0   # конечная частота диапазона (ГГЦ)
f_step = 0.05   # шаг перестройки частоты (ГГц)

p_in = 10.0   # мощность дБм

# исочник питания
i_source = 100   # макс ток потребления (мА)
# ---------------------------

rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

file_name = f'xlsx/divisor-24-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    result = []

    fs = [0.45, 1.0, 5.0, 10.0, 15.0]
    us = [4.7, 5.0, 5.3]

    gen.write(f':POW:POW {p_in}dbm')
    gen.write('OUTP ON')

    for f_gen in fs:
        print('set freq', f_gen)

        gen.write(f'SOUR:FREQ:CW {f_gen}ghz')
        for u_src in us:
            src.write(f'APPLY {u_src}V,{i_source}ma,1')
            src.write('OUTP:CHAN1 ON')
            src.write('OUTP:MAST ON')

            time.sleep(0.6)

            curr = src.query('MEAS:CURR?')

            print('read curr', curr)

            result.append([f_gen, u_src, float(curr) * 1_000])

    df = pd.DataFrame(result,
                      columns=['F, GHz',
                               f'Uin, V',
                               f'I, mA'])
    print(df)

    df.to_excel(file_name)


if __name__ == '__main__':
    measure_1()
