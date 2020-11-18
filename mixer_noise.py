import datetime
import openpyxl
import time
import visa

import numpy as np
import pandas as pd

from openpyxl.chart import LineChart, Reference

from config import instruments

# --- параметры измерения ---
f_start = 10   # MHz
f_step = 100   # MHz
f_end = 1600   # MHz
# ---------------------------

rm = visa.ResourceManager()

sa = rm.open_resource(instruments['анализатор спектра'])

file_name = f'xlsx/mixer-noise-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    raw = sa.query('TRAC1:DATA? TRACE1')

    result = [float(v) for v in raw.split(',')]
    fs = [round(x, 2) for x in np.arange(start=f_start, stop=f_end + 100, step=f_step)]

    result = [[f, r] for f, r in zip(fs, result)]

    df = pd.DataFrame(result, columns=['F, MHz', 'Noise, dB'])
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
