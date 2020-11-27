import datetime
from itertools import groupby

import openpyxl
import time
import visa

import numpy as np
import pandas as pd

from openpyxl.chart import LineChart, Reference

from config import instruments

# --- параметры измерения ---
# питание
u_sources = [4.7, 5.0, 5.3]   # напряжение питания (В)
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

file_name = f'xlsx/vco-3-4-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')
    sa.write(f'frequency:start {band_start}mhz')
    sa.write(f'frequency:stop {band_end}mhz')

    result = []

    ucs = [round(x, 2) for x in np.arange(start=uc_start, stop=uc_end + 0.2, step=uc_step)]
    # ucs = [round(x, 1) for x in np.linspace(uc_start, uc_end, int((uc_end - uc_start) / uc_step) + 1, endpoint=True)]

    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')
    for u_src in u_sources:
        print('set source', u_src)
        src.write(f'APPLY {u_src}V,{i_source}ma,1')

        for uc in ucs:
            src.write(f'APPLY {uc}V,{i_source}ma,2')

            sa.write('CALC1:MARK1 ON')
            sa.write(f'CALC1:MARK1:MAX:AUTO ON')

            time.sleep(0.5)

            freq = float(sa.query(f'CALC1:MARK1:x?'))

            print('read freq', freq)

            result.append([u_src, uc, freq / 1_000_000])

    u_srcs = sorted({el[0] for el in result})
    res = {el[0]: list(el[1]) for el in groupby(result, key=lambda el: el[0])}
    res = {k: [el[2] for el in v] for k, v in res.items()}

    cols = ['Uc, V'] + [f'F@Usrc={u}, MHz' for u in res.keys()]
    df = pd.DataFrame(
        [[u] + freqs for u, *freqs in zip(u_srcs, *res.values())],
        columns=cols
    )
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


if __name__ == '__main__':
    measure_1()
