import datetime
import random
import time
import visa

import numpy as np


class Token:
    cancelled = False


token = Token()


instruments = {
    'Осциллограф': 'GPIB1::7::INSTR',
    'Gen RF': 'GPIB1::19::INSTR',
    'Gen ref': 'GPIB1::6::INSTR',
    'Источник': 'GPIB1::3::INSTR',
    'Мультиметр': 'GPIB1::22::INSTR',
    'АС': 'GPIB1::18::INSTR',
    'АЦ': 'GPIB1::9::INSTR',
}


rm = visa.ResourceManager()

# osc = rm.open_resource(instruments['Осциллограф'])
# gen_rf = rm.open_resource(instruments['Gen RF'])
gen_rf = rm.open_resource(instruments['Gen ref'])
# mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['АС'])
src = rm.open_resource(instruments['Источник'])
# pna = rm.open_resource(instruments['АЦ'])

# osc.send = osc.write
# gen_rf.send = gen_rf.write
gen_rf.send = gen_rf.write
# mult.send = mult.write
sa.send = sa.write
src.send = src.write
# pna.send = pna.write

mock_enabled = False
GIGA = 1_000_000_000
MEGA = 1_000_000
MILLI = 1 / 1_000
MICRO = 1 / 1_000_000
NANO = 1 / 1_000_000_000

calibrated_pows_rf = {}
calibrated_pows_mod = {}
calibrated_pows_lo = {}

secondary = {
    'rf_f_min': 1.15 * GIGA,
    'rf_f_max': 1.65 * GIGA,
    'rf_f_step': 0.10 * GIGA,
    'rf_p_min': -30.0,
    'rf_p_max': 0.0,
    'rf_p_step': 10.0,

    'src_u': 3.3,
    'src_i_max': 60,

    'sa_span': 100 * MEGA,
    'sa_ref_lev': 10,
}


def measure_1():
    rf_f_min = secondary['rf_f_min']
    rf_f_max = secondary['rf_f_max']
    rf_f_step = secondary['rf_f_step']
    rf_p_min = secondary['rf_p_min']
    rf_p_max = secondary['rf_p_max']
    rf_p_step = secondary['rf_p_step']

    src_v = secondary['src_u']
    src_i_max = secondary['src_i_max']

    sa_span = secondary['sa_span']
    sa_ref_lev = secondary['sa_ref_lev']

    gen_rf.send('*RST')
    sa.send('*RST')
    src.send('*RST')

    # setup
    gen_rf.send(f':OUTP:MOD:STAT OFF')

    sa.send(':CAL:AUTO OFF')
    sa.send(':CALC:MARK1:MODE POS')
    sa.send(f':SENS:FREQ:SPAN {sa_span}Hz')
    sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_ref_lev}')
    # sa.send(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')

    src.send(f'APPLY p6v,{src_v}V,{src_i_max}mA')
    src.send('OUTP ON')

    freq_rf_values = [round(x, 3) for x in np.arange(start=rf_f_min, stop=rf_f_max + 0.0001, step=rf_f_step)] \
        if rf_f_min != rf_f_max else [rf_f_min]
    pow_rf_values = [round(x, 3) for x in np.arange(start=rf_p_min, stop=rf_p_max + 0.0001, step=rf_p_step)] \
        if rf_p_min != rf_p_max else [rf_p_min]

    # measurement
    for _ in range(3):
        for rf_freq in freq_rf_values:
            gen_rf.send(f'SOUR:FREQ {rf_freq}Hz')
            for rf_pow in pow_rf_values:
                gen_rf.send(f'SOUR:POW {rf_pow}dbm')
                gen_rf.send(f'OUTP:STAT ON')

                if not mock_enabled:
                    time.sleep(0.5)

                center_freq = rf_freq
                sa.send(':CALC:MARK1:MODE POS')
                sa.send(f':SENSe:FREQuency:CENTer {center_freq}Hz')
                p = sa.send(f':CALCulate:MARKer1:X:CENTer {center_freq}Hz')

    src.send('OUTP OFF')
    sa.send(':CAL:AUTO ON')
    gen_rf.send(f'OUTP:STAT OFF')
    return []


if __name__ == '__main__':
    measure_1()
