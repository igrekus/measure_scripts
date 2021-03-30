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
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\test.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Trc1','c:\users\instrument\desktop\res\test.s1p',FORM,LOGP",
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
