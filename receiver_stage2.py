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

osc = rm.open_resource(instruments['Осциллограф'])
gen_rf = rm.open_resource(instruments['Gen RF'])
gen_ref = rm.open_resource(instruments['Gen ref'])
# mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['АС'])
src = rm.open_resource(instruments['Источник'])
# pna = rm.open_resource(instruments['АЦ'])

osc.send = osc.write
gen_rf.send = gen_rf.write
gen_ref.send = gen_ref.write
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
    'rf_f_max': 6.5 * GIGA,
    'rf_f_step': 0.05 * GIGA,
    'rf_p_min': -100,
    'rf_p_max': -60,
    'rf_p_step': 5,

    'ref_f': 100 * MEGA,
    'ref_p': 0,  # dBm

    'src_v': 3.3,
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

    ref_f = secondary['ref_f']
    ref_p = secondary['ref_p']

    src_v = secondary['src_v']
    src_i_max = secondary['src_i_max']

    sa_span = secondary['sa_span']
    sa_ref_lev = secondary['sa_ref_lev']

    gen_ref.send('*RST')
    gen_rf.send('*RST')
    sa.send('*RST')
    src.send('*RST')

    # setup
    sa.send(':CAL:AUTO OFF')
    sa.send(':CALC:MARK1:MODE POS')
    sa.send(f':SENS:FREQ:SPAN {sa_span}Hz')
    sa.send(f'DISP:WIND:TRAC:Y:RLEV {sa_ref_lev}')
    # sa.send(f'DISP:WIND:TRAC:Y:PDIV {sa_scale_y}')

    # osc.send(f':ACQ:AVERage {osc_`avg`}')

    osc.send(f':CHANnel1:DISPlay ON')
    osc.send(f':CHANnel2:DISPlay ON')

    osc.send(f':CHAN1:SCALE 500')  # V
    osc.send(f':CHAN2:SCALE 500')
    # osc.send(':TIMEBASE:SCALE 10E-8')  # ms / div

    # osc.send(':TRIGger:MODE EDGE')
    # osc.send(':TRIGger:EDGE:SOURCe CHANnel1')
    # osc.send(':TRIGger:LEVel CHANnel1,0')
    # osc.send(':TRIGger:EDGE:SLOPe POSitive')

    # osc.send(':MEASure:VAMPlitude channel1')
    # osc.send(':MEASure:VAMPlitude channel2')
    # osc.send(':MEASure:PHASe CHANnel1,CHANnel2')
    # osc.send(':MEASure:FREQuency CHANnel1')

    # gen_rf.send(f':OUTP:MOD:STAT OFF')
    src.send(f'APPLY p6v,{src_v}V,{src_i_max}mA')
    # src.send(f'APPLY p25v,{src_u_d}V,{src_i_d}mA')

    # pow_rf_values = [round(x, 3) for x in np.arange(start=rf_p_min, stop=rf_p_max + 0.0001, step=rf_p_step)] \
    #     if rf_p_min != rf_p_max else [rf_p_min]
    # freq_rf_values = [round(x, 3) for x in np.arange(start=rf_f_min, stop=rf_f_max + 0.0001, step=rf_f_step)] \
    #     if rf_f_min != rf_f_max else [rf_f_min]

    freq_rf_values = ['0_5', '1', '2', '5', '10', '20', '50']
    pow_rf_values = ['1', '1_25', '1_5']
    pows = {
        '1': -104,
        '1_25': -102,
        '1_5': -104,
    }


    r"""
    1. load waveform - :WMEMory1:LOAD "C:\Users\Administrator\Documents\Yastrebov\test2.wfm"
    2. turn 1 chan off - :CHAN1:DISP OFF
    3. add chan3 to chan1 - :FUNCtion1:ADD channel3,channel1
    4. show result - :FUNCtion1:DISPlay
    """
    # gen_ref.send(f':OUTP:MOD:STAT OFF')
    gen_ref.send(f'SOUR:POW {ref_p}dbm')
    gen_ref.send(f'SOUR:FREQ {ref_f}dbm')
    gen_ref.send(f'OUTP:STAT ON')

    # measurement
    timbases = [
        2.5 * 2 * MICRO, 2.5 * 2 * MICRO, 2.5 * 2 * MICRO,
        2.5 * 2 * MICRO, 2.5 * 2 * MICRO, 2.5 * 2 * MICRO,
        1 * 2 * MICRO, 1 * 2 * MICRO, 1 * 2 * MICRO,
        500 * 2 * NANO, 500 * 2 * NANO, 500 * 2 * NANO,
        500 * 2 * NANO, 250 * 2 * NANO, 250 * 2 * NANO,
        100 * 2 * NANO, 100 * 2 * NANO, 100 * 2 * NANO,
        50 * 2 * NANO,  50 * 2 * NANO,  50 * 2 * NANO,
    ]
    osc.send(':CHAN1:DISP OFF')
    osc.send(':CHAN2:DISP OFF')

    index = 0
    for rf_freq in freq_rf_values:
        rf_freq_num = (float(rf_freq.replace('_', '.')) + 1_600) * MEGA
        gen_rf.send(f'SOUR:FREQ {rf_freq_num}')
        for rf_p in pow_rf_values:
            timebase = timbases[index]
            osc.send(f':TIMEBASE:RANGE {timebase}s')
            index += 1
            osc.send(f':chan1:SCALE 1V')  # V
            osc.send(f':chan2:SCALE 1V')

            gen_rf.send(f'SOUR:POW {pows[rf_p]}dbm')
            gen_ref.send(f'OUTP:STAT ON')
            # TODO skip 5 - 1
            if not mock_enabled:
                time.sleep(0.1)

            chan = 1
            osc.send(f':WMEMory{chan}:LOAD "C:\\Users\\Administrator\\Documents\\Yastrebov\\{rf_freq}M{rf_p}V{chan}ch.csv"')
            osc.send(f':WMEMory{chan}:YOFFset 0')
            osc.send(f':WMEMory{chan}:YRANge 2V')

            chan = 2
            osc.send(f':WMEMory{chan}:LOAD "C:\\Users\\Administrator\\Documents\\Yastrebov\\{rf_freq}M{rf_p}V{chan}ch.csv"')
            osc.send(f':WMEMory{chan}:YOFFset 0')
            osc.send(f':WMEMory{chan}:YRANge 2V')

            if not mock_enabled:
                time.sleep(1.3)

            r"""
            1. load waveform - :WMEMory1:LOAD "C:\Users\Administrator\Documents\Yastrebov\test2.wfm"
            2. turn 1 chan off - :CHAN1:DISP OFF
            3. add chan3 to chan1 - :FUNCtion1:ADD channel3,channel1
            4. show result - :FUNCtion1:DISPlay
            """

            # center_Freq = 0
            # sa.send(':CALC:MARK1:MODE POS')
            # sa.send(f':SENSe:FREQuency:CENTer {center_freq}GHz')
            # sa.send(f':CALCulate:MARKer1:X:CENTer {center_freq}GHz')

    src.send('OUTP OFF')
    sa.send(':CAL:AUTO ON')
    osc.send('*RST')
    osc.send(':CHAN1:DISP OFF')
    osc.send(':CHAN2:DISP OFF')
    return []


if __name__ == '__main__':
    measure_1()
