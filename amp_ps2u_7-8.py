import datetime
import time
import visa

import numpy as np
import pandas as pd

from config import instruments

rm = visa.ResourceManager()

pna = rm.open_resource(instruments['анализатор цепей'])

# --- параметры измерения ---
# freq
f_start = 0.5
f_end = 7.5
f_step = 0.1

znx_comp_setting = [
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps1u_ps2u\7-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps1u_ps2u\7-500.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps1u_ps2u\7-1000.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps1u_ps2u\8-12.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps1u_ps2u\8-14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps1u_ps2u\8-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps1u_ps2u\8-18.znx'",
]
znx_comp_value = [1, 1, 1, 1, 1, 1, 1]

# ---------------------------


def measure_7_8():
    # measure 7 - 8
    fs = [round(x, 3) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]
    result = []
    for comp_in, comp_val in zip(znx_comp_setting, znx_comp_value):
        print('load setting:', comp_in)
        pna.write(comp_in)
        time.sleep(1)

        for f in fs:
            print('set F:', f)
            pna.write(f'FREQ:CW {f}GHz')
            time.sleep(1)

            res = pna.query('CALC:STAT:NLIN:COMP:RES?')
            print('read comp in/out:', res)
            cmp_in, cmp_out = [float(s) for s in res.split(',')]

            out = [f, comp_val, cmp_in, cmp_out]
            print(out)
            result.append(out)

    file_name = f'xlsx/amp-ps1u-7-8-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
    df = pd.DataFrame(result, columns=['F, GHz', 'Comp val', 'Cmp in, dBm', 'Cmp out, dBm'])
    df.to_excel(file_name)


if __name__ == '__main__':
    measure_7_8()
