import datetime
import time
import visa

import pandas as pd

from config import instruments


rm = visa.ResourceManager()

gen = rm.open_resource(instruments['генератор 1'])
sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['генератор 1'])


def pairwise(iterable):
    evens = iterable[::2]
    odds = iterable[1::2]
    return zip(evens, odds)


def measure_phase_noise():
    src.write('APPLY 3.3v,50ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    gen.write(f'SOUR:FREQ:CW 1ghz')
    gen.write(':POW:POW 13dbm')
    gen.write('OUTP ON')

    time.sleep(0.5)

    sa.write('INST:SEL PNOISE')
    sa.write('ADJ:ALL')

    time.sleep(15)

    res = sa.query('TRAC? TRACE1')

    df = pd.DataFrame(pairwise([float(v) for v in res.split(',')]), columns=['F, GHz', 'Noise'])
    print(df)

    df.to_excel(f'xlsx/mult-phasenoise-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')


if __name__ == '__main__':
    measure_phase_noise()
