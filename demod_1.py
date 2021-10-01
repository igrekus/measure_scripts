import ast
import time
import visa

import numpy as np


class Token:
    cancelled = False


instruments = {
    # 'Осциллограф': 'GPIB1::7::INSTR',
    'Анализатор': 'GPIB1::18::INSTR',
    'P LO': 'GPIB1::19::INSTR',
    'P RF': 'GPIB1::6::INSTR',
    'Источник': 'GPIB1::3::INSTR',
    'Мультиметр': 'GPIB1::22::INSTR',
}

rm = visa.ResourceManager()

gen_lo = rm.open_resource(instruments['P LO'])
gen_mod = rm.open_resource(instruments['P RF'])
# osc = rm.open_resource(instruments['Осциллограф'])
src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['Анализатор'])

gen_lo.send = gen_lo.write
gen_mod.send = gen_mod.write
# osc.send = osc.write
src.send = src.write
mult.send = mult.write
sa.send = sa.write

mock_enabled = False

token = Token()

MILLI = 1 / 1_000
GIGA = 1_000_000_000
MEGA = 1_000_000
KILO = 1_000

calibrated_pows_lo = {}
calibrated_pows_rf = {}

gen_lo.send('*RST')
gen_mod.send('*RST')
# osc = rm.open_resource(instruments['Осциллограф'])
src.send('*RST')
mult.send('*RST')
sa.send('*RST')


def measure_1():
    secondary = {'Plo_min': -10.0,
                 'Plo_max': 0.0,
                 'Plo_delta': 5.0,
                 'Flo_min': 0.05,
                 'Flo_max': 6.05,
                 'Flo_delta': 0.2,
                 'Fmod': 1.0,
                 'Umod': 30.0,
                 'Uoffs': 250.0,
                 'is_Flo_div2': False,
                 'D': True,
                 'Usrc': 5.0,
                 'UsrcD': 3.3,
                 'sa_rlev': 10.0,
                 'sa_scale_y': 10.0,
                 'sa_span': 10.0,
                 'sa_avg_state': True,
                 'sa_avg_count': 16.0}

    def set_read_marker(freq):
        sa.send(f':CALCulate:MARKer1:X {freq}Hz')
        if not mock_enabled:
            time.sleep(0.01)
        return float(sa.query(':CALCulate:MARKer:Y?'))

    lo_pow_start = secondary['Plo_min']
    lo_pow_end = secondary['Plo_max']
    lo_pow_step = secondary['Plo_delta']
    lo_f_start = secondary['Flo_min'] * GIGA
    lo_f_end = secondary['Flo_max'] * GIGA
    lo_f_step = secondary['Flo_delta'] * GIGA

    lo_f_is_div2 = secondary['is_Flo_div2']
    d = secondary['D']

    mod_f = secondary['Fmod'] * MEGA
    mod_u = secondary['Umod']  # %
    mod_u_offs = secondary['Uoffs'] * MILLI
    mod_f_offs_0 = 0.5 * MEGA

    src_u = secondary['Usrc']
    src_i_max = 200 * MILLI
    src_u_d = secondary['UsrcD']
    src_i_d_max = 20 * MILLI

    sa_rlev = secondary['sa_rlev']
    sa_scale_y = secondary['sa_scale_y']
    sa_span = secondary['sa_span'] * MEGA

    sa_avg_state = 'ON' if secondary['sa_avg_state'] else 'OFF'
    sa_avg_count = secondary['sa_avg_count']

    pow_lo_values = [
        round(x, 3) for x in
        np.arange(start=lo_pow_start, stop=lo_pow_end + 0.002, step=lo_pow_step)
    ] if lo_pow_start != lo_pow_end else [lo_pow_start]

    freq_lo_values = [
        round(x, 3) for x in
        np.arange(start=lo_f_start, stop=lo_f_end + 0.0001, step=lo_f_step)
    ]

    waveform_filename = 'WFM1:SINE_TEST_WFM'

    gen_lo.send(f':OUTP:MOD:STAT OFF')
    gen_mod.send(f':OUTP:MOD:STAT OFF')

    gen_f_mul = 2 if d else 1
    gen_lo.send(f':FREQ:MULT {gen_f_mul}')
    # gen_mod.send(f':FREQ:MULT {gen_f_mul}')

    gen_mod.send(f':RAD:ARB OFF')
    gen_mod.send(f':RAD:ARB:WAV "{waveform_filename}"')
    gen_mod.send(f':RAD:ARB:BASE:FREQ:OFFS {mod_f + mod_f_offs_0}Hz')
    gen_mod.send(f':RAD:ARB:RSC {mod_u}')
    gen_mod.send(f':DM:IQAD:EXT:COFF {mod_u_offs}V')
    gen_mod.send(f':DM:IQAD ON')
    gen_mod.send(f':DM:STAT ON')
    gen_mod.send(f':DM:IQAD:EXT:IQAT 0db')

    src.send(f'APPLY p6v,{src_u}V,{src_i_max}A')
    src.send(f'APPLY p25v,{src_u_d}V,{src_i_d_max}A')

    sa.send(':CAL:AUTO OFF')
    sa.send(f':SENS:FREQ:SPAN {sa_span}Hz')
    sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
    sa.send(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')
    sa.send(':CALC:MARK1:MODE POS')
    sa.send(f'AVER:COUNT {sa_avg_count}')
    sa.send(f'AVER {sa_avg_state}')

    res = []
    for lo_pow in pow_lo_values:

        for lo_freq in freq_lo_values:

            freq_sa = lo_freq
            if lo_f_is_div2:
                lo_freq *= 2

            if token.cancelled:
                gen_lo.send(f'OUTP:STAT OFF')
                gen_mod.send(f'OUTP:STAT OFF')
                gen_mod.send(f':RAD:ARB OFF')
                if not mock_enabled:
                    time.sleep(0.5)
                src.send('OUTPut OFF')

                gen_lo.send(f'SOUR:POW {lo_pow_start}dbm')
                gen_lo.send(f'SOUR:FREQ {lo_f_start}Hz')

                sa.send(':CAL:AUTO ON')
                raise RuntimeError('measurement cancelled')

            pow_loss = calibrated_pows_lo.get(lo_pow, dict()).get(lo_freq, 0) / 2
            gen_lo.send(f'SOUR:POW {lo_pow + pow_loss}dbm')
            gen_lo.send(f'SOUR:FREQ {lo_freq}Hz')

            # TODO hoist out of the loops
            src.send('OUTPut ON')

            gen_lo.send(f'OUTP:STAT ON')
            gen_mod.send(f'OUTP:STAT ON')
            gen_mod.send(f':RAD:ARB ON')

            # time.sleep(0.1)
            if not mock_enabled:
                time.sleep(0.6)

            sa.send(f'DISP:WIND:TRAC:X:OFFS {0}Hz')
            center_f = freq_sa / 2 if d else freq_sa
            sa.send(f':SENSe:FREQuency:CENTer {center_f}Hz')
            offset = freq_sa / 2 if d else 0
            sa.send(f'DISP:WIND:TRAC:X:OFFS {offset}Hz')

            print('>>>>>>>>>>>>>>>>>>>>>>>>', lo_freq, freq_sa, center_f, offset)

            if not mock_enabled:
                time.sleep(1)

            if lo_f_is_div2:
                f_out = freq_sa + mod_f
                sa_p_out = set_read_marker(f_out)

                f_carr = freq_sa
                sa_p_carr = set_read_marker(f_carr)

                f_sb = freq_sa - mod_f
                sa_p_sb = set_read_marker(f_sb)

                f_3_harm = freq_sa - 3 * mod_f
                sa_p_3_harm = set_read_marker(f_3_harm)
            else:
                f_out = freq_sa - mod_f
                sa_p_out = set_read_marker(f_out)

                f_carr = freq_sa
                sa_p_carr = set_read_marker(f_carr)

                f_sb = freq_sa + mod_f
                sa_p_sb = set_read_marker(f_sb)

                f_3_harm = freq_sa + 3 * mod_f
                sa_p_3_harm = set_read_marker(f_3_harm)

            # lo_p_read = float(gen_lo.query('SOUR:POW?'))
            # lo_f_read = float(gen_lo.query('SOUR:FREQ?'))

            src_u_read = src_u
            src_i_read = float(mult.query('MEAS:CURR:DC? 1A,DEF'))

            raw_point = {
                'lo_p': lo_pow,
                'lo_f': lo_freq,
                'src_u': src_u_read,  # power source voltage as set in GUI
                'src_i': src_i_read,
                'sa_p_out': sa_p_out,
                'sa_p_carr': sa_p_carr,
                'sa_p_sb': sa_p_sb,
                'sa_p_3_harm': sa_p_3_harm,
                'loss': pow_loss,
            }

            if mock_enabled:
                raw_point = mocked_raw_data[index]
                raw_point['loss'] = pow_loss
                raw_point['lo_f'] = lo_freq
                index += 1

            print(raw_point)
            res.append(raw_point)

    gen_lo.send(f'OUTP:STAT OFF')
    gen_mod.send(f'OUTP:STAT OFF')
    gen_mod.send(f':RAD:ARB OFF')
    if not mock_enabled:
        time.sleep(0.5)
    src.send('OUTPut OFF')

    gen_lo.send(f'SOUR:POW {lo_pow_start}dbm')
    gen_lo.send(f'SOUR:FREQ {lo_f_start}Hz')

    sa.send(':CAL:AUTO ON')

    if not mock_enabled:
        with open('out.txt', mode='wt', encoding='utf-8') as f:
            f.write(str(res))

    return res


if __name__ == '__main__':
    measure_1()

"""

"""
