import datetime
import openpyxl
import time
import visa

import numpy as np
import pandas as pd

from openpyxl.chart import LineChart, Reference

instruments = {
    'анализатор спектра': 'GPIB1::18::INSTR',
    'источник питания 1': 'GPIB1::4::INSTR',
}

# --- параметры измерения ---
params = {
    'u_src': 4.7,
    'i_src_max': 300.0,
    'u_vco_min': 0.0,
    'u_vco_max': 10.0,
    'u_vco_delta': 1.0,
    'sa_min': 9.0,
    'sa_max': 12.5,
    'sa_rlev': 30.0,
    'is_harm_relative': False,
    'is_u_src_drift': False,
    'u_src_drift': 10.0,
    'is_p_out_2': False,
}

# ---------------------------

rm = visa.ResourceManager()

sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

sa.send = sa.write
src.send = src.write

file_name = f'xlsx/vco-pulsar-1-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'

GIGA = 1_000_000_000
MEGA = 1_000_000


def measure():
    print(sa.query('*IDN?'))
    print(src.query('*IDN?'))

    sa.send('*RST')
    src.send('*RST')

    u_src = params['u_src']
    i_src_max = params['i_src_max'] / 1_000

    u_tune_min = params['u_vco_min']
    u_tune_max = params['u_vco_max']
    u_tune_step = params['u_vco_delta']
    i_tune_max = 10 / 1_000

    sa_f_start = params['sa_min'] * GIGA
    sa_f_stop = params['sa_max'] * GIGA
    sa_rlev = params['sa_rlev']

    """
    1. src chan 1 - set source voltage (4.7V), current limit - 300mA - APPLY p6v,4.7V,300mA
    2. src chan 1 - set 0V, limit 10 mA  (0 - 10V) - APPLY p25v,0V,10mA
    3. sa set window freq limits - 
       DISP:WIND:TRAC:Y:RLEV 30
       :SENS:FREQ:STAR 9GHz
       :SENS:FREQ:STOP 11GHz
       sa.send(':CAL:AUTO OFF')
       sa.send(':CALC:MARK1:MODE POS')
    4. src - output on
    5. go measure cycle
       +step
       APPLY p6v,4.7V,300mA
       
       find peak - CALC:MARK1:MAX
       CALC1:MARK1:X?
       CALC1:MARK1:Y?
       MEAS:CURR1?
    """

    result = []

    ucs = [round(x, 2) for x in np.arange(start=u_tune_min, stop=u_tune_max + 0.002, step=u_tune_step)]
    # ucs = [round(x, 1) for x in np.linspace(uc_start, uc_end, int((uc_end - uc_start) / uc_step) + 1, endpoint=True)]
    print(ucs)

    src.send(f'APPLY p6v,{u_src}V,{i_src_max}A')
    src.send(f'APPLY p25v,{ucs[0]}V,{i_tune_max}A')

    sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
    sa.send(f':SENS:FREQ:STAR {sa_f_start}Hz')
    sa.send(f':SENS:FREQ:STOP {sa_f_stop}Hz')
    sa.send(':CAL:AUTO OFF')
    sa.send(':CALC:MARK1:MODE POS')

    src.send('OUTP ON')

    for uc in ucs:
        src.send(f'APPLY p6v,{u_src}V,{i_src_max}A')
        src.send(f'APPLY p25v,{uc}V,{i_tune_max}A')

        time.sleep(1)

        sa.send('CALC:MARK1:MAX')

        time.sleep(0.1)

        read_f = float(sa.query(f'CALC1:MARK1:X?')) / MEGA
        read_p = float(sa.query(f'CALC1:MARK1:Y?'))
        read_i = float(src.query('MEAS:CURR? p6v')) * 1_000

        res = [u_src, uc, read_f , read_p, read_i]
        print('read values', res)

        result.append(res)

    df = pd.DataFrame(result, columns=['Uпит, В', 'Uупр, В', 'Fвых, МГц', 'Pвых, дБм', 'Iпот, мА', ])
    df.to_excel(file_name, engine='openpyxl', index=False)

    sa.send(':CAL:AUTO ON')
    src.send('OUTP OFF')


if __name__ == '__main__':
    measure()
