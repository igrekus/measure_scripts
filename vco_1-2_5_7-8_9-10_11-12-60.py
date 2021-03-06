import datetime
import openpyxl
import time
import visa

import numpy as np
import pandas as pd

from openpyxl.chart import LineChart, Reference

from config import instruments

# --- параметры измерения ---
# питание
u_source = 5.0   # напряжение питания (В)
i_source = 100   # макс ток потребления (мА)

# цифровой вход управления
uc_start = 0.0
uc_end = 20.0
uc_step = 0.5

# полоса АС
band_start = 1   # MHz
band_end = 440   # MHz
# ---------------------------

rm = visa.ResourceManager()

sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

file_name = f'xlsx/vco-1-2_5_7-8_9-10_11-12-{datetime.datetime.now().isoformat().replace(":", ".")}-60.xlsx'


def measure_1():
    src.write(f'APPLY {u_source}V,{i_source}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')
    time.sleep(0.5)
    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')
    sa.write(f'frequency:start {band_start}mhz')
    sa.write(f'frequency:stop {band_end}mhz')

    result = []

    ucs = [round(x, 2) for x in np.arange(start=uc_start, stop=uc_end + 0.2, step=uc_step)]
    # ucs = [round(x, 1) for x in np.linspace(uc_start, uc_end, int((uc_end - uc_start) / uc_step) + 1, endpoint=True)]

    for uc in ucs:
        src.write(f'APPLY {uc}V,{i_source}ma,2')

        sa.write('CALC1:MARK1 ON')
        sa.write(f'CALC1:MARK1:MAX:AUTO ON')

        time.sleep(0.5)

        freq = float(sa.query(f'CALC1:MARK1:x?'))

        f_x2 = freq * 2
        f_x3 = freq * 3

        sa.write('CALC1:MARK2 ON')
        sa.write(f'CALC1:MARK2:X {f_x2}hz')
        sa.write('CALC1:MARK3 ON')
        sa.write(f'CALC1:MARK3:X {f_x3}hz')

        time.sleep(0.5)

        pw1 = sa.query(f'CALC1:MARK1:Y?')
        pw2 = sa.query(f'CALC1:MARK2:Y?')
        pw3 = sa.query(f'CALC1:MARK3:Y?')

        curr = src.query('MEAS:CURR?')
        print('read pows', pw1, pw2, pw3)
        print('read curr', curr)

        result.append([uc, freq / 1_000_000, float(pw1), float(pw2), float(pw3), float(curr) * 1_000])

    us = [row[0] for row in result]
    fs = [row[1] for row in result]

    u_pairs = [(u1, u2) for u1, u2 in zip(us, us[1:])]
    f_pairs = [(f1, f2) for f1, f2 in zip(fs, fs[1:])]

    ss = [round(_calc_s(fp, up), 1) for fp, up in zip(f_pairs, u_pairs)]

    for r, s in zip(result, ss):
        r.append(s)

    df = pd.DataFrame(result, columns=['Uc, V', f'F, MHz', f'Px1, bDm', f'Px2, bDm', f'Px3, bDm', 'Isrc, mA', 'Tune, MHz/V'])
    print(df)

    df.to_excel(file_name)

    # wb = openpyxl.open(file_name)
    # ws = wb.active
    #
    # rows = len(df)
    # data = Reference(ws, range_string=f'{ws.title}!C1:C{rows + 1}')
    # xs = Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}')
    #
    # chart = LineChart()
    # chart.add_data(data, titles_from_data=True)
    # chart.set_categories(xs)
    #
    # ws.add_chart(chart, f'E4')
    #
    # wb.close()
    # wb.save(file_name)


def _calc_s(f_pair, u_pair):
    f1, f2 = f_pair
    u1, u2 = u_pair
    return (f2 - f1) / (u2 - u1)


if __name__ == '__main__':
    measure_1()
