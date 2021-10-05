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
    'P LO': 'GPIB1::19::INSTR',
    'P MOD': 'GPIB1::6::INSTR',
    'Источник': 'GPIB1::4::INSTR',
    'Мультиметр': 'GPIB1::22::INSTR',
    'АС': 'GPIB1::18::INSTR',
    'АЦ': 'GPIB1::9::INSTR',
}


rm = visa.ResourceManager()

# gen_lo = rm.open_resource(instruments['P LO'])
# gen_mod = rm.open_resource(instruments['P MOD'])
# mult = rm.open_resource(instruments['Мультиметр'])
# sa = rm.open_resource(instruments['АС'])
src = rm.open_resource(instruments['Источник'])
pna = rm.open_resource(instruments['АЦ'])

# gen_lo.send = gen_lo.write
# gen_mod.send = gen_mod.write
# mult.send = mult.write
# sa.send = sa.write
src.send = src.write
pna.send = pna.write

mock_enabled = False
GIGA = 1_000_000_000
MEGA = 1_000_000
MILLI = 1 / 1_000

calibrated_pows_rf = {}
calibrated_pows_mod = {}
calibrated_pows_lo = {}

secondary = {
    'sweep_points': 401,
    'f_min': 1 * GIGA,
    'f_max': 2 * GIGA,
    'pow_in': -20,
    'src_v': 3.3,
    'src_i_max': 60,
}


def measure_1():
    sweep_points = secondary['sweep_points']
    pna_f_min = secondary['f_min']
    pna_f_max = secondary['f_max']
    pow_in = secondary['pow_in']
    src_v = secondary['src_v']
    src_i_max = secondary['src_i_max']

    pna.send('SYST:PRES')
    pna.query('*OPC?')
    # pna.send('SENS1:CORR ON')

    pna.send(f'SYSTem:FPRESet')

    pna.send('CALC1:PAR:DEF:EXT "CH1_S11",S11')
    # pna.send('CALC1:PAR:DEF:EXT "CH1_S21",S21')

    pna.send(f'DISPlay:WINDow1:STATe ON')
    pna.send(f"DISPlay:WINDow1:TRACe1:FEED 'CH1_S11'")
    # pna.send(f"DISPlay:WINDow1:TRACe2:FEED 'CH1_S21'")

    # pna.send(f'SENSe{chan}:SWEep:TRIGger:POINt OFF')
    pna.send(f'SOUR1:POW1 {pow_in}dbm')

    pna.send(f'SENS1:SWE:POIN {sweep_points}')

    pna.send(f'SENS1:FREQ:STAR {pna_f_min}')
    pna.send(f'SENS1:FREQ:STOP {pna_f_max}')
    # pna.send(f'SENS1:POW:ATT AREC, {primary["Pin"]}')

    pna.send('SENS1:SWE:MODE CONT')
    pna.send(f'FORM:DATA ASCII')

    src.send('INST:SEL OUTP1')
    src.send(f'APPLY {src_v}V,{src_i_max}mA')
    src.send('OUTP ON')

    # measurement
    out = []

    for p in [-15, -16, -15, -16, -15, -16, -15, -16, -15, -16, -15, -16, -15, -16, -15, -16]:
        pna.send(f'CALC1:PAR:SEL "CH1_S11"')
        pna.query('*OPC?')
        # res = pna.query(f'CALC1:DATA:SNP? 1')

        if not mock_enabled:
            time.sleep(0.5)

        pna.send(f'SOUR1:POW1 {p}dbm')

        # pna.send(f'CALC1:PAR:SEL "CH1_S21"')
        # pna.query('*OPC?')

        if not mock_enabled:
            time.sleep(0.5)

        offs = random.randint(-4, 4)
        pna.send(f'CALC:OFFS:MAGN {offs}')
        slop = random.random()
        pna.send(f'CALC:OFFS:MAGN:SLOP {slop}')

        if not mock_enabled:
            time.sleep(0.1)

        pna.send(f'DISP:WIND:TRAC:Y:AUTO')

        out.append(p)

    src.send('OUTP OFF')
    pna.send('SYST:PRES')
    return out


if __name__ == '__main__':
    measure_1()
