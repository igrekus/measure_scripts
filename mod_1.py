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
src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['Анализатор'])

secondary = {
    'Plo_min': -10.0,
    'Plo_max': 0.0,
    'Plo_delta': 5.0,
    'Flo_min': 0.05,
    'Flo_max': 6.05,
    'Flo_delta': 0.1,
    'is_Flo_div2': False,
    'Fmod': 1.0,   # MHz
    'Umod': 30,   # %
    'Uoffs': 250,   # mV
    'Usrc': 5.0,
    'sa_rlev': 10.0,
    'sa_scale_y': 10.0,
    'sa_span': 10.0,   # MHz
}


def measure_1():

    # gen_lo.write('*RST')
    src.write('*RST')
    mult.write('*RST')
    sa.write('*RST')

    lo_pow_start = secondary['Plo_min']
    lo_pow_end = secondary['Plo_max']
    lo_pow_step = secondary['Plo_delta']
    lo_f_start = secondary['Flo_min']
    lo_f_end = secondary['Flo_max']
    lo_f_step = secondary['Flo_delta']

    lo_f_is_div2 = secondary['is_Flo_div2']

    mod_f = secondary['Fmod']
    mod_u = secondary['Umod']
    mod_u_offs = secondary['Uoffs']
    mod_f_offs_0 = 0.5   # MHz

    src_u = secondary['Usrc']
    src_i_max = 200   # mA

    sa_rlev = secondary['sa_rlev']
    sa_scale_y = secondary['sa_scale_y']
    sa_span = secondary['sa_span']

    pow_lo_values = [
        round(x, 3)for x in
        np.arange(start=lo_pow_start, stop=lo_pow_end + 0.002, step=lo_pow_step)
    ] if lo_pow_start != lo_pow_end else [lo_pow_start]

    freq_lo_values = [
        round(x, 3) for x in
        np.arange(start=lo_f_start, stop=lo_f_end + 0.0001, step=lo_f_step)
    ]

    file_name = 'WFM1:SINE_TEST_WFM'

    gen_lo.write(f':OUTP:MOD:STAT OFF')
    gen_lo.write(f':RAD:ARB OFF')
    gen_lo.write(f':RAD:ARB:WAV "{file_name}"')
    gen_lo.write(f':RAD:ARB:BASE:FREQ:OFFS {mod_f + mod_f_offs_0}MHz')
    gen_lo.write(f':RAD:ARB:RSC {mod_u}')
    gen_lo.write(f':DM:IQAD:EXT:COFF {mod_u_offs}mV')
    gen_lo.write(f':DM:IQAD ON')
    gen_lo.write(f':DM:STAT ON')
    gen_lo.write(f':DM:IQAD:EXT:IQAT 0db')

    src.write(f'APPLY p6v,{src_u}V,{src_i_max}mA')

    sa.write(':CAL:AUTO OFF')
    sa.write(f':SENS:FREQ:SPAN {sa_span}MHz')
    sa.write(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
    sa.write(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')
    sa.write(':CALC:MARK1:MODE POS')

    res = []
    for pow_lo in pow_lo_values:

        for freq_lo in freq_lo_values:

            if lo_f_is_div2:
                freq_lo /= 2

            gen_lo.write(f'SOUR:POW {pow_lo}dbm')
            gen_lo.write(f'SOUR:FREQ {freq_lo}GHz')

            # TODO hoist out of the loops
            src.write('OUTPut ON')

            gen_lo.write(f'OUTP:STAT ON')
            gen_lo.write(f':RAD:ARB ON')

            time.sleep(0.5)

            sa.write(f':SENSe:FREQuency:CENTer {freq_lo}GHz')

            f_out = freq_lo - mod_f / 1_000
            sa.write(f':CALCulate:MARKer1:X {f_out}GHz')
            time.sleep(0.1)
            sa_p_out = float(sa.query(':CALCulate:MARKer:Y?'))

            f_carr = freq_lo
            sa.write(f':CALCulate:MARKer1:X {f_carr}GHz')
            time.sleep(0.1)
            sa_p_carr = float(sa.query(':CALCulate:MARKer:Y?'))

            f_sb = freq_lo + mod_f / 1_000
            sa.write(f':CALCulate:MARKer1:X {f_sb}GHz')
            time.sleep(0.1)
            sa_p_sb = float(sa.query(':CALCulate:MARKer:Y?'))

            f_3_harm = freq_lo + 3 * (mod_f / 1_000)
            sa.write(f':CALCulate:MARKer1:X {f_3_harm}GHz')
            time.sleep(0.1)
            sa_p_mod_f_x3 = float(sa.query(':CALCulate:MARKer:Y?'))

            lo_p_read = float(gen_lo.query('SOUR:POW?'))
            lo_f_read = float(gen_lo.query('SOUR:FREQ?'))

            src_u_read = src_u
            src_i_read = float(mult.query('MEAS:CURR:DC? 1A,DEF'))

            raw_point = {
                'lo_p': lo_p_read,
                'lo_f': lo_f_read,
                'src_u': src_u_read,   # power source voltage as set in GUI
                'src_i': src_i_read,
                'sa_p_out': sa_p_out,
                'sa_p_carr': sa_p_carr,
                'sa_p_sb': sa_p_sb,
                'sa_p_mod_f_x3': sa_p_mod_f_x3,
            }

            print(raw_point)
            res.append(raw_point)

    gen_lo.write(f'OUTP:STAT OFF')
    time.sleep(0.5)
    src.write('OUTPut OFF')

    gen_lo.write(f'SOUR:POW {lo_pow_start}dbm')
    gen_lo.write(f'SOUR:FREQ {lo_f_start}GHz')

    with open('out.txt', mode='wt', encoding='utf-8') as f:
        f.write(str(res))

    print(res)

    gen_lo.write(f'OUTP:STAT OFF')
    gen_lo.write(f':RAD:ARB OFF')
    src.write('OUTP OFF')
    sa.write(':CAL:AUTO ON')
    return res


if __name__ == '__main__':
    measure_1()


"""

"""
