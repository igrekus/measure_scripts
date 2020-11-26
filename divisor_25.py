import datetime
import openpyxl
import time
import visa

import pandas as pd
import numpy as np

from openpyxl.chart import Reference, LineChart
from string import ascii_uppercase

from config import instruments

# --- параметры измерения ---
# генератор
f_in = 2.0

p_in = 0.0

# питание
u_source = 5.0   # напряжение питания (В)
i_source = 100   # макс ток потребления (мА)

# цифровой вход управления
uc_start = 0.0
uc_end = 5.0
uc_step = 0.1

# коэфициент деления
coeff = 16
# ---------------------------

rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

file_name = f'xlsx/divisor-25-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    src.write(f'APPLY {u_source}V,{i_source}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')
    sa.write('frequency:start 1mhz')
    sa.write('frequency:stop 30ghz')

    result = []

    gen.write(f'SOUR:FREQ:CW {f_in}ghz')
    gen.write(f':POW:POW {p_in}dbm')
    gen.write('OUTP ON')

    ucs = [round(x, 1) for x in np.linspace(uc_start, uc_end, int((uc_end - uc_start) / uc_step) + 1, endpoint=True)]

    for uc in ucs:
        print('set Uc', uc)

        src.write(f'APPLY {uc}V,{i_source}ma,2')

        time.sleep(0.5)

        f_sa = f_in / coeff

        f_mark1 = f_sa
        sa.write('CALC1:MARK1 ON')
        sa.write(f'CALC1:MARK1:X {f_mark1}ghz')

        time.sleep(0.5)

        pw1 = sa.query(f'CALC1:MARK1:Y?')

        print('read power', pw1)

        result.append([uc, float(pw1)])

    cols = ['Uc, V', f'Pout@F/16, dB']
    df = pd.DataFrame(result, columns=cols)
    print(df)

    df.to_excel(file_name)

    wb = openpyxl.open(file_name)
    ws = wb.active

    rows = len(df)
    data = Reference(ws, range_string=f'{ws.title}!C1:{ascii_uppercase[len(cols)]}{rows + 1}')
    xs = Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}')

    chart = LineChart()
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(xs)

    ws.add_chart(chart, f'E4')

    wb.save(file_name)
    wb.close()


if __name__ == '__main__':
    measure_1()
