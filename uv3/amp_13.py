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
    # measure 12-13
    is_ = [
         20.0,  21.0,  22.0,  23.0,  24.0,  25.0,  26.0,  27.0,  28.0,  29.0,  30.0
    ]
    us = [
        3.113, 3.131, 3.150, 3.167, 3.185, 3.201, 3.220, 3.235, 3.254, 3.270, 3.285
    ]
    result = []
    for u, i in zip(us, is_):
        print(f'set voltage: {u}')
        src.write(f'APPLY {u}V,{35}ma,1')
        src.write('OUTP:CHAN1 ON')
        src.write('OUTP:MAST ON')

        time.sleep(1)

        curr = src.query('MEAS:CURR?')
        u_out = mult.query('MEASURE:VOLTAGE:DC?')
        res = [u, i, float(u_out)]

        print('read params:', res)
        result.append(res)

    file_name = f'xlsx/amp-UV3-13-out-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
    df = pd.DataFrame(result, columns=['Usrc, V', 'Isrc, mA', 'Uout, V'])
    df.to_excel(file_name)


if __name__ == '__main__':
    measure_12_13()
