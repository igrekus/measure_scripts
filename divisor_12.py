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
f_step = 0.1   # шаг перестройки частоты (ГГц)

p_in = 0   # мощность дБм

# исочник питания
u_start = 5.0   # напряжение питания (В)
u_end = 5.0
u_step = 5.0
i_source = 100   # макс ток потребления (мА)

# коэфициент деления
coeff = 2
# ---------------------------

rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

file_name = f'xlsx/divisor-12-coeff_{coeff}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_12():
    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')

    gen.write(f':POW:POW {p_in}dbm')
    gen.write('OUTP ON')

    sa.write('frequency:start 1mhz')
    sa.write('frequency:stop 30ghz')

    result = []

    fs = [round(x, 1) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]
    # us = [round(x, 1) for x in np.linspace(u_start, u_end, int((u_end - u_start) / u_step) + 1, endpoint=True)]
    us = [4.75, 5.0, 5.25]

    for u_source in us:
        print('set U source', u_source)

        src.write(f'APPLY {u_source}V,{i_source}ma,1')
        src.write('OUTP:CHAN1 ON')
        src.write('OUTP:MAST ON')

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

            result.append([f_gen, u_source, float(pw1)])

    freqs = sorted({el[0] for el in result})

    res = {el[0]: list(el[1]) for el in groupby(result, key=lambda el: el[1])}
    res = {k: [el[2] for el in v] for k, v in res.items()}

    df = pd.DataFrame(
        [[f] + pows for f, *pows in zip(freqs, *res.values())],
        columns=['Fin, GHz'] + [f'Pout@F/{coeff}&Uin={u}, dBm' for u in res],
    )

    print(df)

    df.to_excel(file_name)


if __name__ == '__main__':
    measure_12()
