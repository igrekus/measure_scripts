import datetime
import math
import time
import visa

import numpy as np
import pandas as pd

from config import instruments

# --- параметры измерения ---
# volt format:
# I, mA : U, V
u_src = {
    '22.5': 3.16,
    '25': 3.204,
    '27.5': 3.237,
}
# ---------------------------

rm = visa.ResourceManager()

sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])


def measure_1():
    # measure 14
    for amp, volt in u_src.items():
        print(f'set voltage: {volt}')
        src.write(f'APPLY {volt}V,{35}ma,1')
        src.write('OUTP:CHAN1 ON')
        src.write('OUTP:MAST ON')

        time.sleep(1)

        sa.write(':INIT:IMM;*WAI')

        time.sleep(15)

        freqs = sa.query('FREQ:TABL:DATA?').split(',')
        noise = sa.query('TRAC? TRACE1,NOISE').split(',')

        result = []
        for f, n in zip(freqs, noise):
            result.append([float(f), float(n)])

        file_name = f'xlsx/amp-UV3-14-{amp}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
        df = pd.DataFrame(result, columns=['F, GHz', 'Noise, dB'])
        df.to_excel(file_name)


if __name__ == '__main__':
    measure_1()
