import datetime
import time
import visa

import pandas as pd
import numpy as np

from itertools import groupby

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

file_name = f'xlsx/divisor-22-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    src.write(f'APPLY {u_source}V,{i_source}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')
    sa.write('frequency:start 1mhz')
    sa.write('frequency:stop 30ghz')

    result = []

    fs = [0.45, 1.0, 5.0, 10.0, 15.0]
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

            result.append([f_gen, p_gen, float(curr) * 1_000])

    freqs = sorted({el[0] for el in result})

    res = {el[0]: list(el[1]) for el in groupby(sorted(result, key=lambda el: el[1]), key=lambda el: el[1])}
    res = {k: [el[2] for el in v] for k, v in res.items()}

    df = pd.DataFrame(
        [[f] + currs for f, *currs in zip(freqs, *res.values())],
        columns=['Fin, GHz'] + [f'I@Pin={p}, mA' for p in res],
    )

    print(df)
    df.to_excel(file_name)


if __name__ == '__main__':
    measure_1()
