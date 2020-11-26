import datetime
import openpyxl
import time
import visa

import pandas as pd
import numpy as np

from openpyxl.chart import Reference, LineChart

from config import instruments

# --- параметры измерения ---
# генератор
f_start = 0.45   # начальная частота диапазона (ГГц)
f_end = 15.0   # конечная частота диапазона (ГГЦ)
f_step = 0.05   # шаг перестройки частоты (ГГц)

p_start = -15.0   # мощность дБм
p_end = 10.0
p_step = 1.0

# исочник питания
u_source = 5.0   # напряжение питания (В)
i_source = 100   # макс ток потребления (мА)
# ---------------------------

rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

file_name = f'xlsx/divisor-23-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    src.write(f'APPLY {u_source}V,{i_source}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    result = []

    fs = [2.0]
    # fs = [round(x, 1) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]
    ps = [round(x, 1) for x in np.linspace(p_start, p_end, int((p_end - p_start) / p_step) + 1, endpoint=True)]

    for f_gen in fs:
        print('set freq', f_gen)

        for p_gen in ps:
            gen.write(f'SOUR:FREQ:CW {f_gen}ghz')
            gen.write(f':POW:POW {p_gen}dbm')
            gen.write('OUTP ON')

            time.sleep(0.6)

            curr = src.query('MEAS:CURR?')

            print('read curr', curr)

            result.append([p_gen, f_gen, float(curr) * 1_000])

    cols = [f'Pin, dB', 'F, GHz', f'I@F=2GHz, mA']
    df = pd.DataFrame(result, columns=cols)
    print(df)

    df.to_excel(file_name)

    wb = openpyxl.open(file_name)
    ws = wb.active

    rows = len(df)
    data = Reference(ws, range_string=f'{ws.title}!D1:D{rows + 1}')
    xs = Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}')

    chart = LineChart()
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(xs)

    ws.add_chart(chart, f'F4')

    wb.save(file_name)
    wb.close()


if __name__ == '__main__':
    measure_1()
