import datetime
import math
import time
import visa

import numpy as np
import pandas as pd

from config import instruments

rm = visa.ResourceManager()

pna = rm.open_resource(instruments['анализатор цепей'])
src = rm.open_resource(instruments['источник питания 1'])
mult = rm.open_resource(instruments['мультиметр'])

# --- параметры измерения ---

in_files = [
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\1-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\2-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\3-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\4-90.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\20.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\21.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\22.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\23.csv',FORM,LOGP,COMM",
]

# ---------------------------


def measure_20_23():
    # measure 20 - 23
    volt = 4.95
    amp = 90
    print(f'set voltage: {volt}')
    src.write(f'APPLY {volt}V,{amp + 10}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')
    time.sleep(1)

    for in_file, out_file in zip(in_files, out_files):
        print('load', in_file)
        pna.write(in_file)
        time.sleep(1)
        print('save', out_file)
        pna.write(out_file)
        time.sleep(1)


if __name__ == '__main__':
    measure_20_23()
