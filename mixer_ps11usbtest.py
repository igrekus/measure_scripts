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
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\1-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\1-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\1-1500.znx'",
]
out_files = [ 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\tlsb\1-10.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\tlsb\1-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\tlsb\1-1500.csv',FORM,LOGP,COMM",
]


def measure_1():
    for in_file, out_file in zip(in_files, out_files):
        print('load', in_file)
        pna.write(in_file)
        time.sleep(5)
        print('save', out_file)
        pna.write(out_file)
        time.sleep(1)


if __name__ == '__main__':
    measure_1()
