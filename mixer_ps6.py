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
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\1-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\1-50.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\1-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\2.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\3-0.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\3-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\3+10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\4.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\5-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\5-1000.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\5-1500.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\6.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\7.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\17-1.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\17-2.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\17-3.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\17-4.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\20-1.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\20-2.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\20-3.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\23-1.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\23-2.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\24.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\26-1.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps6u\26-2.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\1-10.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\1-50.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\1-100.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\2.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\3-0.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\3-10.s1p',FORM,LOGP", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\3+10.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\4.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\5-10.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\5-1000.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\5-1500.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\6.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\7.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc8','c:\users\instrument\desktop\res\17-1.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc9','c:\users\instrument\desktop\res\17-2.s1p',FORM,LOGP", 
    r"MMEM:STOR:TRAC 'Trc10','c:\users\instrument\desktop\res\17-3.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc11','c:\users\instrument\desktop\res\17-4.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc8','c:\users\instrument\desktop\res\20-1.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc9','c:\users\instrument\desktop\res\20-2.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc10','c:\users\instrument\desktop\res\20-3.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc10','c:\users\instrument\desktop\res\23-1.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'Trc11','c:\users\instrument\desktop\res\23-2.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'LO_leak','c:\users\instrument\desktop\res\24.s1p',FORM,LOGP",
    r"MMEM:STOR:TRAC 'RF_REFL','c:\users\instrument\desktop\res\26-1.s1p',FORM,LOGP", 
    r"MMEM:STOR:TRAC 'Trc11','c:\users\instrument\desktop\res\26-2.s1p',FORM,LOGP",
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
