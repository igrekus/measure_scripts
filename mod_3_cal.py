import datetime
import time
import visa

import numpy as np


instruments = {
    'Осциллограф': 'GPIB1::7::INSTR',
    'P LO': 'GPIB1::6::INSTR',
    'P RF': 'GPIB1::20::INSTR',
    'Источник': 'GPIB1::3::INSTR',
    'Мультиметр': 'GPIB1::22::INSTR',
    'Анализатор': 'GPIB1::18::INSTR',
}

rm = visa.ResourceManager()

gen_lo = rm.open_resource(instruments['P LO'])
gen_mod = rm.open_resource(instruments['P RF'])
src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['Анализатор'])

secondary = {
    'Plo': -5.0,
    'Flo_min': 0.6,
    'Flo_max': 6.6,
    'Flo_delta': 1.0,
    'is_Flo_div2': False,
    'Fmod_min': 1.0,   # MHz
    'Fmod_max': 501.0,   # MHz
    'Fmod_delta': 10.0,   # MHz
    'Uoffs': 250,   # mV
    'Usrc': 5.0,
    'sa_rlev': 10.0,
    'sa_scale_y': 10.0,
    'sa_span': 10.0,   # MHz
}


def measure_1():

    gen_lo.write('*RST')
    gen_mod.write('*RST')
    src.write('*RST')
    mult.write('*RST')
    sa.write('*RST')

    mod_f_min = secondary['Fmod_min']
    mod_f_max = secondary['Fmod_max']
    mod_f_delta = secondary['Fmod_delta']
    mod_p = -5

    sa_rlev = secondary['sa_rlev']
    sa_scale_y = secondary['sa_scale_y']
    sa_span = secondary['sa_span']

    mod_f_values = [
        round(x, 3)for x in
        np.arange(start=mod_f_min, stop=mod_f_max + 0.0002, step=mod_f_delta)
    ]

    gen_mod.write(f'SOUR:POW {mod_p}dbm')

    sa.write(':CAL:AUTO OFF')
    sa.write(f':SENS:FREQ:SPAN {sa_span}MHz')
    sa.write(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
    sa.write(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')
    sa.write(':CALC:MARK1:MODE POS')

    gen_mod.write(f'OUTP:STAT ON')

    res = []
    for mod_f in mod_f_values:
        gen_mod.write(f'SOUR:FREQ {mod_f}MHz')

        time.sleep(0.8)

        sa_freq = mod_f
        sa.write(f':SENSe:FREQuency:CENTer {sa_freq}MHz')

        time.sleep(0.2)

        sa_p_out = float(sa.query(':CALCulate:MARKer:Y?'))
        loss = mod_p - sa_p_out

        out = [mod_f, sa_p_out, loss]
        res.append(out)
        print(out)

    gen_mod.write(f'OUTP:STAT OFF')
    gen_mod.write(f'SOUR:FREQ {mod_f_min}GHz')

    sa.write(':CAL:AUTO ON')

    with open('out.txt', mode='wt', encoding='utf-8') as f:
        f.write(str(res))
    print(res)
    return res


if __name__ == '__main__':
    measure_1()
