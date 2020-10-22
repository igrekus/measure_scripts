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
f_step = 0.1   # шаг перестройки частоты (ГГц)
p_in = 10.0   # мощность дБм

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

file_name = f'xlsx/divisor-16-coeff_{coeff}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    src.write(f'APPLY {u_source}V,{i_source}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')
    sa.write('frequency:start 1mhz')
    sa.write('frequency:stop 30ghz')

    result = []

    fs = [round(x, 1) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]

    for f_gen in fs:
        print('set freq', f_gen)

        gen.write(f'SOUR:FREQ:CW {f_gen}ghz')
        gen.write(f':POW:POW {p_in}dbm')
        gen.write('OUTP ON')

        time.sleep(0.5)

        f_sa = f_gen / coeff

        f_mark1 = f_sa * 2
        f_mark2 = f_sa * 3
        f_mark3 = f_sa * 4
        f_mark4 = f_sa * 5
        sa.write('CALC1:MARK1 ON')
        sa.write(f'CALC1:MARK1:X {f_mark1}ghz')
        sa.write('CALC1:MARK2 ON')
        sa.write(f'CALC1:MARK2:X {f_mark2}ghz')
        sa.write('CALC1:MARK3 ON')
        sa.write(f'CALC1:MARK3:X {f_mark3}ghz')
        sa.write('CALC1:MARK4 ON')
        sa.write(f'CALC1:MARK4:X {f_mark4}ghz')

        time.sleep(0.5)

        pw1 = sa.query(f'CALC1:MARK1:Y?')
        pw2 = sa.query(f'CALC1:MARK2:Y?')
        pw3 = sa.query(f'CALC1:MARK3:Y?')
        pw4 = sa.query(f'CALC1:MARK4:Y?')

        print('read power', pw1, pw2, pw3, pw4)

        result.append([f_gen, float(pw1), float(pw2), float(pw3), float(pw4)])

    df = pd.DataFrame(result,
                      columns=['F, GHz',
                               f'Pout@F/{coeff} x2, dB',
                               f'Pout@F/{coeff} x3, dB',
                               f'Pout@F/{coeff} x4, dB',
                               f'Pout@F/{coeff} x5, dB'])
    print(df)

    df.to_excel(file_name)


if __name__ == '__main__':
    measure_1()
