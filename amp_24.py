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
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\24-10-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\24-10-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\24-10-110.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\24-200-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\24-200-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv1u\24-200-110.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\24-10-90.s2p',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\24-10-100.s2p',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\24-10-110.s2p',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\24-200-90.s2p',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\24-200-100.s2p',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv1u\24-200-110.s2p',FORM,LOGP,COMM",
]

volts = [4.96, 5.15, 5.32] * 2

# ---------------------------


def measure_20_23():
    # measure 24
    for in_file, out_file, volt in zip(in_files, out_files, volts):
        print(f'set voltage: {volt}')
        src.write(f'APPLY {volt}V,{120}ma,1')
        src.write('OUTP:CHAN1 ON')
        src.write('OUTP:MAST ON')
        time.sleep(1)

        print('load', in_file)
        pna.write(in_file)
        time.sleep(1)
        print('save', out_file)
        pna.write(out_file)
        time.sleep(1)


if __name__ == '__main__':
    measure_20_23()
