import datetime
import math
import time
import visa

import numpy as np
import pandas as pd

from config import instruments

# --- параметры измерения ---
# freq
f_start = 0.01
f_end = 3.0
f_step = 0.01

# freqs for pow-pow measurement
freqs = [0.01, 1, 1.5]
# ---------------------------

rm = visa.ResourceManager()

pna = rm.open_resource(instruments['анализатор цепей'])
src = rm.open_resource(instruments['источник питания 1'])
mult = rm.open_resource(instruments['мультиметр'])

znx_comp_in = [
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\5-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\6-90.znx'",
]
znx_comp_val = [1, 3]


def measure_16_17():
    # measure 16 - 17
    fs = [round(x, 3) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]
    result_comp = []

    volt = 4.950
    amp = 90
    print(f'set voltage: {volt}')
    src.write(f'APPLY {volt}V,{amp + 10}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')
    time.sleep(1)

    for comp_in, comp_val in zip(znx_comp_in, znx_comp_val):
        print('load', comp_in)
        pna.write(comp_in)
        time.sleep(1)

        for f in fs:
            print('set F:', f)
            pna.write(f'FREQ:CW {f}GHz')
            time.sleep(1)
            res = pna.query('CALC:STAT:NLIN:COMP:RES?')
            print('read comp in/out:', res)
            _, comp_out = [float(s) for s in res.split(',')]

            # pae = 100 * (cmp_out - cmp_in) / (10 * math.log10(volt * amp * 1_000))
            # out = [f, comp_val, cmp_in, cmp_out, pae]

            out = [f, comp_val, comp_out]
            print(out)
            result_comp.append(out)

    file_name = f'xlsx/amp-UV1-16-17-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
    df = pd.DataFrame(result_comp, columns=['F, GHz', 'Comp value', 'Cmp out, dBm'])
    df.to_excel(file_name)


if __name__ == '__main__':
    measure_16_17()
