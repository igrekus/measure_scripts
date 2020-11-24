import datetime
import openpyxl
import time
import visa

import pandas as pd
import numpy as np

from itertools import groupby
from openpyxl.chart import LineChart, Reference
from string import ascii_uppercase

from config import instruments

# --- параметры измерения ---
# генератор
f_start = 0.45   # начальная частота диапазона (ГГц)
f_end = 15.0   # конечная частота диапазона (ГГЦ)
f_step = 0.1   # шаг перестройки частоты (ГГц)

p_start = -15   # мощность дБм
p_end = 5
p_step = 5

# исочник питания
u_source = 5.0   # напряжение питания (В)
i_source = 100   # макс ток потребления (мА)

# коэфициент деления
coeff = 2
# ---------------------------

rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

file_name = f'xlsx/divisor-6-coeff_{coeff}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_6():
    src.write(f'APPLY {u_source}V,{i_source}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')
    sa.write('frequency:start 1mhz')
    sa.write('frequency:stop 30ghz')

    result = []

    fs = [round(x, 1) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]
    ps = [round(x, 1) for x in np.linspace(p_start, p_end, int((p_end - p_start) / p_step) + 1, endpoint=True)]

    for p_gen in ps:
        print('set pow', p_gen)

        gen.write(f':POW:POW {p_gen}dbm')
        gen.write('OUTP ON')

        for f_gen in fs:
            print('set freq', f_gen)

            gen.write(f'SOUR:FREQ:CW {f_gen}ghz')

            time.sleep(0.5)

            f_sa = f_gen / coeff

            sa.write('CALC1:MARK1 ON')
            sa.write(f'CALC1:MARK1:X {f_sa}ghz')

            time.sleep(0.5)

            pw1 = sa.query(f'CALC1:MARK1:Y?')

            print('read power', pw1)

            result.append([f_gen, p_gen, float(pw1)])

    freqs = sorted({el[0] for el in result})

    res = {el[0]: list(el[1]) for el in groupby(result, key=lambda el: el[1])}
    res = {k: [el[2] for el in v] for k, v in res.items()}

    cols = ['Fin, GHz'] + [f'Pout@F/{coeff}&Pin={p}, dBm' for p in res.keys()]
    df = pd.DataFrame(
        [[f] + pows for f, *pows in zip(freqs, *res.values())],
        columns=cols
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

    ws.add_chart(chart, f'I4')

    wb.save(file_name)
    wb.close()


if __name__ == '__main__':
    measure_6()
