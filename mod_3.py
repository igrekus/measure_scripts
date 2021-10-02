import datetime
import time
import visa

import numpy as np


class Token:
    cancelled = False


token = Token()


instruments = {
    'Осциллограф': 'GPIB1::7::INSTR',
    'P LO': 'GPIB1::19::INSTR',
    'P MOD': 'GPIB1::6::INSTR',
    'Источник': 'GPIB1::3::INSTR',
    'Мультиметр': 'GPIB1::22::INSTR',
    'Анализатор': 'GPIB1::18::INSTR',
}


rm = visa.ResourceManager()

gen_lo = rm.open_resource(instruments['P LO'])
gen_mod = rm.open_resource(instruments['P MOD'])
src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['Анализатор'])

gen_lo.send = gen_lo.write
gen_mod.send = gen_mod.write
src.send = src.write
mult.send = mult.write
sa.send = sa.write


secondary = {
    'Plo': 0,
    'Pmod': -5.0,
    'Flo_min': 5,
    'Flo_max': 5,
    'Flo_delta': 1.0,
    'is_Flo_div2': False,
    'D': False,
    'Fmod_min': 1.0,   # MHz
    'Fmod_max': 1301,   # MHz
    'Fmod_delta': 10,   # MHz
    'Uoffs': 600,  # mV
    'Usrc': 4.75,
    'UsrcD': 3.1,
    'sa_rlev': 10.0,
    'sa_scale_y': 10.0,
    'sa_span': 10.0,
    'sa_avg_state': False,
    'sa_avg_count': 16,
    'sep_1': None,
    'u_min': 4.75,
    'u_max': 5.25,
    'u_delta': 0.05
}

mock_enabled = False
GIGA = 1_000_000_000
MEGA = 1_000_000
MILLI = 1 / 1_000

calibrated_pows_rf = {}
calibrated_pows_mod = {}
calibrated_pows_lo = {}


def measure_1():

    def set_read_marker(freq):
        sa.send(f':CALCulate:MARKer1:X {freq}Hz')
        if not mock_enabled:
            time.sleep(0.01)
        return float(sa.query(':CALCulate:MARKer:Y?'))

    lo_pow = secondary['Plo']
    lo_f_start = secondary['Flo_min'] * GIGA
    lo_f_end = secondary['Flo_max'] * GIGA
    lo_f_step = secondary['Flo_delta'] * GIGA

    lo_f_is_div2 = secondary['is_Flo_div2']
    d = secondary['D']

    mod_f_min = secondary['Fmod_min'] * MEGA
    mod_f_max = secondary['Fmod_max'] * MEGA
    mod_f_delta = secondary['Fmod_delta'] * MEGA
    mod_u_offs = secondary['Uoffs'] * MILLI
    mod_pow = secondary['Pmod']

    src_u = secondary['Usrc']
    src_i_max = 200   # mA
    src_u_d = secondary['UsrcD']
    src_i_d_max = 20   # mA

    sa_rlev = secondary['sa_rlev']
    sa_scale_y = secondary['sa_scale_y']
    sa_span = secondary['sa_span'] * MEGA
    sa_avg_state = 'ON' if secondary['sa_avg_state'] else 'OFF'
    sa_avg_count = secondary['sa_avg_count']

    u_start = secondary['u_min']
    u_end = secondary['u_max']
    u_step = secondary['u_delta']

    mod_f_values = [
        round(x, 3)for x in
        np.arange(start=mod_f_min, stop=mod_f_max + 0.0002, step=mod_f_delta)
    ]

    freq_lo_values = [
        round(x, 3) for x in
        np.arange(start=lo_f_start, stop=lo_f_end + 0.0001, step=lo_f_step)
    ]

    u_values = [round(x, 3) for x in np.arange(start=u_start, stop=u_end + 0.002, step=u_step)]

    # region main measure
    gen_f_mul = 2 if d else 1
    gen_lo.send(f':FREQ:MULT {gen_f_mul}')
    gen_lo.send(f':OUTP:MOD:STAT OFF')
    gen_lo.send(f':RAD:ARB OFF')
    # gen_lo.send(f':DM:IQAD:EXT:COFF {mod_u_offs}')

    src.send(f'APPLY p25v,{src_u_d}V,{src_i_d_max}mA')
    src.send(f'APPLY p6v,{src_u}V,{src_i_max}mA')
    src.send('OUTPut ON')

    # gen_lo.send(f':DM:IQAD ON')
    # gen_lo.send(f':DM:STAT ON')

    sa.send(':CAL:AUTO OFF')
    sa.send(f':SENS:FREQ:SPAN {sa_span}')
    sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_rlev}')
    sa.send(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')
    sa.send(f'AVER:COUNT {sa_avg_count}')
    sa.send(f'AVER {sa_avg_state}')
    sa.send(':CALC:MARK1:MODE POS')

    gen_lo.send(f'OUTP:STAT ON')
    gen_mod.send(f'OUTP:STAT ON')

    res = []
    for lo_freq in freq_lo_values:

        sa_freq = lo_freq

        if lo_f_is_div2:
            lo_freq *= 2

        gen_lo.send(f'SOUR:FREQ {lo_freq}')
        # gen_lo.send(f':DM:IQAD OFF')
        # gen_lo.send(f':DM:IQAD ON')

        for mod_f in mod_f_values:

            if token.cancelled:
                gen_mod.send(f'OUTP:STAT OFF')
                gen_lo.send(f'OUTP:STAT OFF')

                time.sleep(0.5)

                src.send('OUTPut OFF')

                gen_lo.send(f':DM:IQAD OFF')
                gen_lo.send(f':DM:STAT OFF')
                gen_lo.send(f'SOUR:POW {lo_pow}dbm')
                gen_lo.send(f'SOUR:FREQ {lo_f_start}')

                gen_mod.send(f'SOUR:FREQ {mod_f_min}')

                sa.send(':CAL:AUTO ON')
                raise RuntimeError('measurement cancelled')

            lo_loss = calibrated_pows_lo.get(lo_pow, dict()).get(lo_freq, 0) / 2
            mod_loss = calibrated_pows_mod.get(mod_pow, dict()).get(mod_f, 0)
            out_loss = calibrated_pows_rf.get(lo_freq, dict()).get(mod_f, 0) / 2

            gen_mod.send(f'SOUR:POW {mod_pow + mod_loss}dbm')
            gen_lo.send(f'SOUR:POW {lo_pow + lo_loss}dbm')

            gen_mod.send(f'SOUR:FREQ {mod_f}')
            # sa.send(f':SENS:FREQ:SPAN {mod_f * 10}')

            if not mock_enabled:
                time.sleep(0.3)

            if lo_f_is_div2:
                sa_center_freq = sa_freq + mod_f
            else:
                sa_center_freq = sa_freq - mod_f

            sa.send(f'DISP:WIND:TRAC:X:OFFS {0}Hz')
            center_f = (lo_freq / 2 - mod_f) if d else sa_center_freq
            sa.send(f':SENSe:FREQuency:CENTer {center_f}Hz')
            offset = (lo_freq / 2) if d else 0
            sa.send(f'DISP:WIND:TRAC:X:OFFS {offset}Hz')

            if not mock_enabled:
                time.sleep(0.3)

            if lo_f_is_div2:
                f_out = sa_freq + mod_f
                sa_p_out = set_read_marker(f_out)
            else:
                f_out = (center_f + offset) if d else (sa_freq - mod_f)
                sa_p_out = set_read_marker(f_out)

            src_u_read = src_u
            src_i_read = float(mult.query('MEAS:CURR:DC? 1A,DEF'))

            raw_point = {
                'lo_p': lo_pow,
                'lo_f': lo_freq,
                'mod_f': mod_f,
                'src_u': src_u_read,   # power source voltage as set in GUI
                'src_i': src_i_read,
                'sa_p_out': sa_p_out,
                'out_loss': out_loss,
            }

            print(raw_point)
            res.append(raw_point)

    gen_mod.send(f'OUTP:STAT OFF')
    gen_lo.send(f'OUTP:STAT OFF')

    time.sleep(0.5)

    src.send('OUTPut OFF')

    gen_lo.send(f':DM:STAT OFF')
    gen_lo.send(f'SOUR:POW {lo_pow}dbm')
    gen_lo.send(f'SOUR:FREQ {lo_f_start}')
    gen_mod.send(f'SOUR:FREQ {mod_f_min}')

    sa.send(':CAL:AUTO ON')


    with open('out.txt', mode='wt', encoding='utf-8') as f:
        f.write(str(res))
    print(res)
    return res


if __name__ == '__main__':
    measure_1()
