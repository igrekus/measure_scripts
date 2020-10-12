import datetime
import time
import visa

import pandas as pd

from config import instruments


rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['генератор 1'])

fs = [1, 2, 3, 4, 5, 6]
ps = [-10, -9, -8, -7, -6, -5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]


def measure_current():
    src.write('APPLY 3.3v,50ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    gen.write(f'SOUR:FREQ:CW 100mhz')
    gen.write(':POW:POW 0dbm')
    gen.write('OUTP ON')

    result = []
    for f in fs:
        gen.write(f'SOUR:FREQ:CW {f}GHz')
        for p in ps:
            print('set f=', f, 'p=', p)
            gen.write(f':POW:POW {p}dbm')

            time.sleep(0.5)

            curr = float(src.query('MEAS:CURR?'))
            print('read current=', curr)
            result.append([f, p, curr])

    df = pd.DataFrame(result, columns=['F, GHz', 'P, dB', 'I, mA'])
    print(df)

    df.to_excel(f'xlsx/mult-current-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')


if __name__ == '__main__':
    measure_current()
