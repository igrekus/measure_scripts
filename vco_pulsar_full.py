import datetime
import openpyxl
import time
import visa

import numpy as np
import pandas as pd

from openpyxl.chart import LineChart, Reference
from openpyxl.chart import Series


instruments = {
    'анализатор спектра': 'GPIB1::18::INSTR',
    'источник питания 1': 'GPIB1::4::INSTR',
}

# --- параметры измерения ---
params = {
    'u_src_drift1': 4.7,
    'u_src_drift2': 0,
    'u_src_drift3': 0,
    'i_src_max': 300.0,
    'u_vco_min': 0.0,
    'u_vco_max': 10.0,
    'u_vco_delta': 1.0,
    'sa_min': 9.0,
    'sa_max': 12.5,
    'sa_rlev': 30.0,
    'sa_span': 50.0,
}

# ---------------------------


file_name = f'xlsx/vco-pulsar-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'

GIGA = 1_000_000_000
MEGA = 1_000_000


def measure():

    def prepare_rig():
        src.send(f'APPLY p6v,{u_dr_1}V,{i_src_max}A')
        src.send(f'APPLY p25v,{ucs[0]}V,{i_tune_max}A')

        sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
        sa.send(f':SENS:FREQ:STAR {sa_f_start}Hz')
        sa.send(f':SENS:FREQ:STOP {sa_f_stop}Hz')
        sa.send(':CAL:AUTO OFF')
        sa.send(':CALC:MARK1:MODE POS')

        src.send('OUTP ON')

    def measure_drift():
        result = []
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

                res = [u_dr, uc, read_f, read_p, read_i]
                print('read values', res)

                result.append(res)

        return result

    def measure_harmonics(multiplier):
        result = []
        sa.send(f':SENS:FREQ:SPAN {sa_span}')
        for uc, f in pairs:
            src.send(f'APPLY p6v,{u_dr_1}V,{i_src_max}A')
            src.send(f'APPLY p25v,{uc}V,{i_tune_max}A')

            time.sleep(1)

            f_xmul = f * multiplier
            sa.send(f':SENS:FREQ:CENT {f_xmul}Hz')

            time.sleep(0.5)

            sa.send('CALC:MARK1:MAX')

            time.sleep(0.05)

            read_p = float(sa.query(f'CALC1:MARK1:Y?'))

            res = [uc, read_p]
            result.append(res)

        return result

    rm = visa.ResourceManager()

    sa = rm.open_resource(instruments['анализатор спектра'])
    src = rm.open_resource(instruments['источник питания 1'])

    sa.send = sa.write
    src.send = src.write

    print(sa.query('*IDN?'))
    print(src.query('*IDN?'))

    src.send('*RST')
    src.send('OUTP OFF')
    sa.send('*RST')

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

    ucs = [round(x, 2) for x in np.arange(start=u_tune_min, stop=u_tune_max + 0.002, step=u_tune_step)]
    u_drs = [u for u in [u_dr_1, u_dr_2, u_dr_3] if u]

    prepare_rig()
    result = measure_drift()

    pairs = [[row[1], row[2] * MEGA] for row in result if row[0] == u_dr_1]

    result_harmonics_x2 = measure_harmonics(multiplier=2)
    result_harmonics_x3 = measure_harmonics(multiplier=3)

    sa.send(':CAL:AUTO ON')
    src.send('OUTP OFF')

    print(result_harmonics_x2)
    print(result_harmonics_x3)

    # TODO group u_dr by measure blocks, don't create tab for drift == 0
    df = pd.DataFrame(result, columns=['Uпит, В', 'Uупр, В', 'Fвых, МГц', 'Pвых, дБм', 'Iпот, мА', ])
    df.to_excel(file_name, engine='openpyxl', index=False)

    # drift single measurement
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

    # fx2
    # [[0.0, -22.006], [1.0, -25.104], [2.0, -23.253], [3.0, -24.527], [4.0, -21.237], [5.0, -23.507], [6.0, -28.615], [7.0, -37.745], [8.0, -38.494], [9.0, -42.876], [10.0, -38.98]]
    # fx3
    # [[0.0, -50.554], [1.0, -46.267], [2.0, -47.991], [3.0, -39.016], [4.0, -34.675], [5.0, -42.85], [6.0, -37.567], [7.0, -39.202], [8.0, -36.298], [9.0, -37.13], [10.0, -41.076]]


def to_excel():
    u_dr_1 = 4.7
    u_dr_2 = 5.0
    u_dr_3 = 5.3

    result_harmonics_x2 = [[0.0, -22.006], [1.0, -25.104], [2.0, -23.253], [3.0, -24.527], [4.0, -21.237], [5.0, -23.507], [6.0, -28.615], [7.0, -37.745], [8.0, -38.494], [9.0, -42.876], [10.0, -38.98]]
    result_harmonics_x3 = [[0.0, -50.554], [1.0, -46.267], [2.0, -47.991], [3.0, -39.016], [4.0, -34.675], [5.0, -42.85], [6.0, -37.567], [7.0, -39.202], [8.0, -36.298], [9.0, -37.13], [10.0, -41.076]]
    result_harmonics_x2 = [[u_dr_1] + row for row in result_harmonics_x2]
    result_harmonics_x3 = [[u_dr_1] + row for row in result_harmonics_x3]

    df_harm_2 = pd.DataFrame(result_harmonics_x2, columns=['Uпит, В', 'Uупр, В', 'Pвых_2, дБм'])
    df_harm_3 = pd.DataFrame(result_harmonics_x3, columns=['Uпит, В', 'Uупр, В', 'Pвых_3, дБм'])

    df = pd.read_excel('./xlsx/vco-pulsar-2-2021-08-11T16.15.57.458610.xlsx', engine='openpyxl')

    df['fdiff'] = df.groupby('Uпит, В')['Fвых, МГц'].diff().shift(-1)
    df['udiff'] = df.groupby('Uпит, В')['Uупр, В'].diff().shift(-1)
    df['S, МГц/В'] = df[df['fdiff'].notna()].apply(lambda row: (row['fdiff'] / row['udiff']) * 100, axis = 1)
    df = df.drop(['fdiff', 'udiff'], axis=1)

    df = pd.merge(df, df_harm_2, how='left', on=['Uпит, В', 'Uупр, В'])
    df = pd.merge(df, df_harm_3, how='left', on=['Uпит, В', 'Uупр, В'])

    df['Pвых_2отн, дБм'] = df[df['Pвых_2, дБм'].notna()].apply(lambda row: -(row['Pвых, дБм'] - row['Pвых_2, дБм']), axis=1)
    df['Pвых_3отн, дБм'] = df[df['Pвых_3, дБм'].notna()].apply(lambda row: -(row['Pвых, дБм'] - row['Pвых_3, дБм']), axis=1)
    df['S, МГц/В'] = df['S, МГц/В'].fillna(0)

    df_udr_1 = df[df['Uпит, В'] == u_dr_1]
    df_udr_2 = df[df['Uпит, В'] == u_dr_2]
    df_udr_3 = df[df['Uпит, В'] == u_dr_3]

    udr_1_s = [df_udr_1.columns.values.tolist()] + df_udr_1.values.tolist()
    udr_2_s = [df_udr_2.columns.values.tolist()] + df_udr_2.values.tolist()
    udr_3_s = [df_udr_3.columns.values.tolist()] + df_udr_3.values.tolist()

    out = []
    for r1, r2, r3 in zip(udr_1_s, udr_2_s, udr_3_s):
        row = r1 + ['', ''] + r2 + ['', ''] + r3
        out.append(row)

    wb = openpyxl.Workbook()
    ws = wb.active

    for row in out:
        ws.append(row)

    rows = len(udr_1_s)

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!C1:C{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!O1:O{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AA1:AA{rows + 1}'),
        ],
        title='Диапазон перестройки',
        loc='B15',
        curve_labels=['Uпит = 4.7В', 'Uпит = 5.0В', 'Uпит = 5.3В']
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!D1:D{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!P1:P{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AB1:AB{rows + 1}'),
        ],
        title='Мощность',
        loc='M15',
        curve_labels=['Uпит = 4.7В', 'Uпит = 5.0В', 'Uпит = 5.3В']
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!I1:I{rows + 1}'),
        ],
        title='Относительный уровень 2й гармоники',
        loc='B30',
        curve_labels=['Uпит = 4.7В']
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!J1:J{rows + 1}'),
        ],
        title='Относительный уровень 3й гармоники',
        loc='M30',
        curve_labels=['Uпит = 4.7В']
    )

    wb.save('out.xlsx')


def _add_chart(ws, xs, ys, title, loc, curve_labels=None):
    chart = LineChart()

    for y, label in zip(ys, curve_labels):
        ser = Series(y, title=label)
        chart.append(ser)

    chart.set_categories(xs)
    chart.title = title

    ws.add_chart(chart, loc)


if __name__ == '__main__':
    # measure()

    to_excel()
