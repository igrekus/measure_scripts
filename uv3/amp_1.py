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
    '90': 3.16,
    '100': 3.204,
    '110': 3.237,
}
subs = {
    '90': 22.5,
    '100': 25.0,
    '110': 27.5
}

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

in_files = [
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\1-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\1-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\1-110.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\2-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\2-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\2-110.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\3-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\3-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\3-110.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\4-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\4-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\4-110.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\1-90.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\1-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\1-110.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\2-90.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\2-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\2-110.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\3-90.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\3-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\3-110.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\4-90.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\4-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\4-110.csv',FORM,LOGP,COMM",
]

znx_comp_in = [
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\5-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\5-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\5-110.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\6-90.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\6-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\6-110.znx'",
]
znx_comp_val = [1, 1, 1, 3, 3, 3]

znx_9 = r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\9.znx'"
znx_10 = r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\10.znx'"

znx_15_in = r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\uv3u\15.znx'"
znx_15_out = r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\uv3u\15.csv',FORM,LOGP,COMM"


def _get_volt_from_file_name(filename):
    return u_src[filename.split('-')[1][:-5]]


def _get_curr_from_file_name(filename):
    return subs[filename.split('-')[1][:-5]]


def measure_1():
    # measure 1 - 4
    for in_file, out_file in zip(in_files, out_files):
        volt = _get_volt_from_file_name(in_file)
        print(f'set voltage: {volt}')
        src.write(f'APPLY {volt}V,{35}ma,1')
        src.write('OUTP:CHAN1 ON')
        src.write('OUTP:MAST ON')

        time.sleep(1)

        print('load', in_file)
        pna.write(in_file)
        time.sleep(1)
        print('save', out_file)
        pna.write(out_file)
        time.sleep(1)

    # measure 5 - 6
    fs = [round(x, 3) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]
    result_comp = []
    for comp_in, comp_val in zip(znx_comp_in, znx_comp_val):
        volt = _get_volt_from_file_name(comp_in)
        amp = _get_curr_from_file_name(comp_in)
        print(f'set voltage: {volt}')
        src.write(f'APPLY {volt}V,{35}ma,1')
        src.write('OUTP:CHAN1 ON')
        src.write('OUTP:MAST ON')

        time.sleep(1)

        print('load', comp_in)
        pna.write(comp_in)
        time.sleep(1)

        for f in fs:
            print('set F:', f)
            pna.write(f'FREQ:CW {f}GHz')
            time.sleep(1)
            res = pna.query('CALC:STAT:NLIN:COMP:RES?')
            print('read comp in/out:', res)
            cmp_in, cmp_out = [float(s) for s in res.split(',')]

            pae = 100 * (cmp_out - cmp_in) / (10 * math.log10(volt * amp * 1_000))

            out = [f, comp_val, cmp_in, cmp_out, pae]
            print(out)
            result_comp.append(out)

    file_name = f'xlsx/amp-UV3-5-4-6-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
    df = pd.DataFrame(result_comp, columns=['F, GHz', 'Comp val','Cmp in, dBm', 'Cmp out, dBm', 'PAE, %'])
    df.to_excel(file_name)

    # measure 9
    print('load znx 9:', znx_9)
    pna.write(znx_9)
    for f in freqs:
        pna.write(f'FREQ:CW {f}GHz')
        time.sleep(1)
        pna.write(f"MMEM:STOR:TRAC 'Trc1','c:\\users\\instrument\\desktop\\res\\uv3u\\9-{f}.csv',FORM,LOGP,COMM",)

    # measure 10
    print('load znx pow:', znx_10)
    pna.write(znx_10)
    for f in freqs:
        pna.write(f'FREQ:CW {f}GHz')
        time.sleep(1)
        pna.write(f"MMEM:STOR:TRAC 'Trc1','c:\\users\\instrument\\desktop\\res\\uv3u\\10-{f}.csv',FORM,LOGP,COMM",)

    # # measure 11
    # select pae from measure 9 for freq = [0.01, 1, 1.5] -> plot 3 freqs

    # measure 15
    volt = 3.16
    amp = 22.5
    print(f'set {volt}V, {amp}mA')
    src.write(f'APPLY {volt}V,{amp}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    print('load state for 15')
    pna.write(znx_15_in)
    time.sleep(2)
    print('save state for 15')
    pna.write(znx_15_out)
    time.sleep(1)


if __name__ == '__main__':
    measure_1()
