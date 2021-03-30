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
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\2-12.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\2-14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\2-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\2-18.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\3.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\4.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\6-5.5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\6-6.5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\6-7.5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\7-12.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\7-14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\7-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\7-18.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\8.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\9.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\11-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\11-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\11-1500.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\12-12.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\12-14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\12-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\12-18.znx'",    
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\13.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\15.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\26-12.znx'",   
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\26-14.znx'",    
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\26-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\26-18.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\27.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\28.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\29.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\30.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\31.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\32.znx'",    
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\lsb\33.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\1-10.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\1-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\1-1500.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\2-12.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\2-14.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\2-16.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\2-18.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\3.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\4.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\6-5.5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\6-6.5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\6-7.5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\7-12.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\7-14.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\7-16.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\7-18.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\8.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\9.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\10.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\11-10.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\11-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\11-1500.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\12-12.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\12-14.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\12-16.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\12-18.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\13.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\14.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\lsb\15.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\26-12.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\26-14.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\26-16.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\26-18.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\27.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\28.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\29.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\30.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\31.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\32.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\lsb\33.csv',FORM,LOGP,COMM",
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
