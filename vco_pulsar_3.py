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
    'sa_span': 50.0,
    'is_harm_relative': False,
    'is_u_src_drift': False,
    'u_src_drift1': 4.7,
    'u_src_drift2': 0,
    'u_src_drift3': 0,
    'is_p_out_2': False,
}

# ---------------------------

rm = visa.ResourceManager()

sa = rm.open_resource(instruments['анализатор спектра'])
src = rm.open_resource(instruments['источник питания 1'])

sa.send = sa.write
src.send = src.write

file_name = f'xlsx/vco-pulsar-3-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'

GIGA = 1_000_000_000
MEGA = 1_000_000


def measure():
    print(sa.query('*IDN?'))
    print(src.query('*IDN?'))

    src.send('*RST')
    src.send('OUTP OFF')
    sa.send('*RST')

    u_src = params['u_src']
    i_src_max = params['i_src_max'] / 1_000

    u_tune_min = params['u_vco_min']
    u_tune_max = params['u_vco_max']
    u_tune_step = params['u_vco_delta']
    i_tune_max = 10 / 1_000

    sa_f_start = params['sa_min'] * GIGA
    sa_f_stop = params['sa_max'] * GIGA
    sa_rlev = params['sa_rlev']
    sa_span = params['sa_span'] * MEGA

    u_dr_1 = params['u_src_drift1']
    u_dr_2 = params['u_src_drift2']
    u_dr_3 = params['u_src_drift3']

    result = []

    ucs = [round(x, 2) for x in np.arange(start=u_tune_min, stop=u_tune_max + 0.002, step=u_tune_step)]
    # ucs = [round(x, 1) for x in np.linspace(uc_start, uc_end, int((uc_end - uc_start) / uc_step) + 1, endpoint=True)]
    print(ucs)

    u_drs = [u for u in [u_dr_1, u_dr_2, u_dr_3] if u]

    src.send(f'APPLY p6v,{u_dr_1}V,{i_src_max}A')
    src.send(f'APPLY p25v,{ucs[0]}V,{i_tune_max}A')

    sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
    sa.send(f':SENS:FREQ:STAR {sa_f_start}Hz')
    sa.send(f':SENS:FREQ:STOP {sa_f_stop}Hz')
    sa.send(':CAL:AUTO OFF')
    sa.send(':CALC:MARK1:MODE POS')

    src.send('OUTP ON')

    for u_dr in u_drs:
        for uc in ucs:
            src.send(f'APPLY p6v,{u_dr}V,{i_src_max}A')
            src.send(f'APPLY p25v,{uc}V,{i_tune_max}A')

            time.sleep(1)

            sa.send('CALC:MARK1:MAX')

            time.sleep(0.1)

            read_f = float(sa.query(f'CALC1:MARK1:X?')) / MEGA
            read_p = float(sa.query(f'CALC1:MARK1:Y?'))

            read_i = float(src.query('MEAS:CURR? p6v')) * 1_000

            res = [u_dr, uc, read_f , read_p, read_i]
            print('read values', res)

            result.append(res)

    pairs = [[row[1], row[2] * MEGA] for row in result if row[0] == u_dr_1]

    result_harmonics_x2 = []
    sa.send(f':SENS:FREQ:SPAN {sa_span}')
    for uc, f in pairs:
        src.send(f'APPLY p6v,{u_dr_1}V,{i_src_max}A')
        src.send(f'APPLY p25v,{uc}V,{i_tune_max}A')

        time.sleep(1)

        fx3 = f * 2
        sa.send(f':SENS:FREQ:CENT {fx3}Hz')

        time.sleep(0.5)

        sa.send('CALC:MARK1:MAX')

        time.sleep(0.05)

        read_px3 = float(sa.query(f'CALC1:MARK1:Y?'))

        res = [uc, read_px3]
        result_harmonics_x2.append(res)

    result_harmonics_x3 = []
    sa.send(f':SENS:FREQ:SPAN {sa_span}')
    for uc, f in pairs:
        src.send(f'APPLY p6v,{u_dr_1}V,{i_src_max}A')
        src.send(f'APPLY p25v,{uc}V,{i_tune_max}A')

        time.sleep(1)

        fx3 = f * 3
        sa.send(f':SENS:FREQ:CENT {fx3}Hz')

        time.sleep(0.1)

        sa.send('CALC:MARK1:MAX')

        time.sleep(0.1)

        read_px3 = float(sa.query(f'CALC1:MARK1:Y?'))

        res = [uc, read_px3]
        result_harmonics_x3.append(res)

    sa.send(':CAL:AUTO ON')
    src.send('OUTP OFF')

    print(result_harmonics_x2)
    print(result_harmonics_x3)

    # TODO group u_dr by measure blocks, don't create tab for drift == 0
    df = pd.DataFrame(result, columns=['Uпит, В', 'Uупр, В', 'Fвых, МГц', 'Pвых, дБм', 'Iпот, мА', ])
    df.to_excel(file_name, engine='openpyxl', index=False)

    # [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0]
    # read values [4.7, 0.0, 9583.0, -0.991, 73.80771999999999]
    # read values [4.7, 1.0, 10003.0, 3.365, 70.38627]
    # read values [4.7, 2.0, 10447.0, 6.424, 69.66817999999999]
    # read values [4.7, 3.0, 10832.0, 7.924, 68.51714]
    # read values [4.7, 4.0, 11263.0, 7.002, 70.75939000000001]
    # read values [4.7, 5.0, 11602.0, 9.263, 75.85638]
    # read values [4.7, 6.0, 11847.0, 9.997, 78.60902999999999]
    # read values [4.7, 7.0, 12010.0, 8.718, 79.35879]
    # read values [4.7, 8.0, 12103.0, 8.829, 79.73191]
    # read values [4.7, 9.0, 12138.0, 8.834, 79.91144]
    # read values [4.7, 10.0, 12150.0, 8.643, 79.92199]
    # [[0.0, -22.006], [1.0, -25.104], [2.0, -23.253], [3.0, -24.527], [4.0, -21.237], [5.0, -23.507], [6.0, -28.615], [7.0, -37.745], [8.0, -38.494], [9.0, -42.876], [10.0, -38.98]]
    # [[0.0, -50.554], [1.0, -46.267], [2.0, -47.991], [3.0, -39.016], [4.0, -34.675], [5.0, -42.85], [6.0, -37.567], [7.0, -39.202], [8.0, -36.298], [9.0, -37.13], [10.0, -41.076]]


if __name__ == '__main__':
    measure()
