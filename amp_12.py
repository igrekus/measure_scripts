import datetime
import math
import time
import visa

import numpy as np
import pandas as pd

from config import instruments

# --- параметры измерения ---

# ---------------------------

rm = visa.ResourceManager()

src = rm.open_resource(instruments['источник питания 1'])
mult = rm.open_resource(instruments['мультиметр'])

def measure_12_13():
    # measure 11-13
    is_ = [
        80.00, 81.00, 82.00, 83.00, 84.00, 85.00, 86.00, 87.00, 88.00, 89.00,
        90.00, 91.00, 92.00, 93.00, 94.00, 95.00, 96.00, 97.00, 98.00, 99.00,
        100.0, 101.0, 102.0, 103.0, 104.0, 105.0, 106.0, 107.0, 108.0, 109.0,
        110.0, 111.0, 112.0, 113.0, 114.0, 115.0, 116.0, 117.0, 118.0, 119.0,
        120.0
    ]
    us = [
        4.775, 4.794, 4.815, 4.824, 4.839, 4.852, 4.866, 4.889, 4.905, 4.922,
        4.949, 4.956, 4.972, 4.990, 5.007, 5.025, 5.042, 5.052, 5.077, 5.094,
        5.112, 5.129, 5.145, 5.164, 5.180, 5.198, 5.216, 5.235, 5.252, 5.270,
        5.288, 5.305, 5.325, 5.343, 5.360, 5.377, 5.395, 5.413, 5.431, 5.450,
        5.469
    ]
    result = []
    for u, i in zip(us, is_):
        print(f'set voltage: {u}')
        src.write(f'APPLY {u}V,{125}ma,1')
        src.write('OUTP:CHAN1 ON')
        src.write('OUTP:MAST ON')

        time.sleep(1)

        curr = src.query('MEAS:CURR?')
        u_out = mult.query('MEASURE:VOLTAGE:DC?')
        res = [u, i, float(u_out)]

        print('read params:', res)
        result.append(res)

    file_name = f'xlsx/amp-UV1-12-in-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
    df = pd.DataFrame(result, columns=['Usrc, V', 'Isrc, mA', 'Uin, V'])
    df.to_excel(file_name)


if __name__ == '__main__':
    measure_12_13()
