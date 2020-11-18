import datetime
import time
import visa

from config import instruments

# --- параметры измерения ---
f_start = 10   # MHz
f_step = 100   # MHz
f_end = 1600   # MHz
# ---------------------------

rm = visa.ResourceManager()

pna = rm.open_resource(instruments['анализатор цепей'])

file_name = f'xlsx/mixer-noise-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'

in_files = [
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps7u\1-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\1-20.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps7u\1-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps7u\1-500.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\1-10.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\1-20.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\1-100.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\1-500.s1p',FORM,LOGP",
]


def measure_1():
    for in_file, out_file in zip(in_files, out_files):
        print('load', in_file)
        pna.write(in_file)
        time.sleep(1)
        print('save', out_file)
        pna.write(out_file)
        time.sleep(1)


if __name__ == '__main__':
    measure_1()
