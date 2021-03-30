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
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\1-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\1-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\1-1500.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\2-12.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\2-14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\2-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\2-18.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\3.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\4.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\6-5.5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\6-6.5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\6-7.5.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\7-12.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\7-14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\7-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\7-18.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\8.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\9.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\11-10.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\11-100.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\11-1500.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\12-12.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\12-14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\12-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\12-18.znx'",    
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\13.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\14.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\15.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\26-12.znx'",   
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\26-14.znx'",    
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\26-16.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\26-18.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\27.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\28.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\29.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\30.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\31.znx'",
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\32.znx'",    
    r"MMEM:LOAD:STATE 1,'c:\users\instrument\desktop\ps11u\usb\33.znx'",
]
out_files = [
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\1-10.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\1-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\1-1500.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\2-12.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\2-14.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\2-16.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\2-18.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\3.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\4.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\6-5.5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\6-6.5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\6-7.5.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\7-12.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\7-14.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\7-16.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\7-18.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\8.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\9.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\10.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\11-10.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\11-100.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\11-1500.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\12-12.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\12-14.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\12-16.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\12-18.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\13.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\14.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Trc4','c:\users\instrument\desktop\res\usb\15.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\26-12.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\26-14.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\26-16.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\26-18.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\27.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\28.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\29.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\30.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\31.csv',FORM,LOGP,COMM",
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\32.csv',FORM,LOGP,COMM", 
    r"MMEM:STOR:TRAC 'Conv','c:\users\instrument\desktop\res\usb\33.csv',FORM,LOGP,COMM",
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
