import datetime
import openpyxl
import time
import visa

import pandas as pd

from string import ascii_uppercase
from itertools import groupby
from openpyxl.chart import Reference, LineChart

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

    res = sorted(result, key=lambda el: el[1])

    freqs = sorted({el[0] for el in result})

    res = {el[0]: list(el[1]) for el in groupby(res, key=lambda el: el[1])}
    res = {k: [el[2] for el in v] for k, v in res.items()}

    cols = ['Fin, GHz'] + [f'I@Uin={u}V, mA' for u in res]
    df = pd.DataFrame(
        [[f] + currs for f, *currs in zip(freqs, *res.values())],
        columns=cols,
    )

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

    ws.add_chart(chart, f'G4')

    wb.save(file_name)
    wb.close()


if __name__ == '__main__':
    measure_1()
