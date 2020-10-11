import datetime
import time
import visa

import pandas as pd
import numpy as np

from config import instruments


rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['генератор 1'])

f_start = 0.1   # 0.1 GHz = 100 MHz
f_end = 6.5   # 6.5 GHz
f_step = 0.1   # 0.1 GHz = 100 MHz


def measure_4pow():
    src.write('APPLY 3.3v,50ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    # sa.write(f'DISP:TRAC1:Y:RLEV -5dbm')
    sa.write(f'BAMD 200khz')

    result = []

    fs = [round(x, 1) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]

    for f_gen in fs:
        print('set freq', f_gen)

        gen.write(f'SOUR:FREQ:CW {f_gen}ghz')
        gen.write(':POW:POW 13.0dbm')
        gen.write('OUTP ON')

        time.sleep(0.5)

        f_sa_x1 = f_gen
        f_sa_x2 = f_gen * 2
        f_sa_x3 = f_gen * 3
        f_sa_x4 = f_gen * 4

        sa.write('frequency:start 20mhz')
        sa.write('frequency:stop 30ghz')
        sa.write('CALC1:MARK1 ON')
        sa.write(f'CALC1:MARK1:X {f_sa_x1}ghz')
        sa.write('CALC1:MARK2 ON')
        sa.write(f'CALC1:MARK2:X {f_sa_x2}ghz')
        sa.write('CALC1:MARK3 ON')
        sa.write(f'CALC1:MARK3:X {f_sa_x3}ghz')
        sa.write('CALC1:MARK4 ON')
        sa.write(f'CALC1:MARK4:X {f_sa_x4}ghz')

        time.sleep(0.5)

        pw1 = sa.query(f'CALC1:MARK1:Y?')
        pw2 = sa.query(f'CALC1:MARK1:Y?')
        pw3 = sa.query(f'CALC1:MARK1:Y?')
        pw4 = sa.query(f'CALC1:MARK1:Y?')

        print('read power', pw1, pw2, pw3, pw4)

        result.append([f_gen, float(pw1), float(pw2), float(pw3), float(pw4)])

    df = pd.DataFrame(result, columns=['F, GHz', 'Px1, dB', 'Px2, dB', 'Px3, dB', 'Px4, dB'])
    print(pd)

    df.to_excel(f'xlsx/умножитель-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')


if __name__ == '__main__':
    print('Прогрмаа для измерения 4х гармоник умножителя')
    measure_4pow()
