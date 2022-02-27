import datetime
import time
import visa

import numpy as np


instruments = {
    'Генератор': 'ASRL6::INSTR',
    'Мощность': 'GPIB2::1::INSTR',
    'Источник': 'GPIB3::4::INSTR',
}

rm = visa.ResourceManager()

src = rm.open_resource(instruments['Источник'])
meter = rm.open_resource(instruments['Мощность'])
gen = rm.open_resource(instruments['Генератор'])

MHz = 1_000_000
GIGA = 1_000_000_000

src.send = src.write
gen.send = gen.write
meter.send = meter.write


def measure():
    '''
    pulse mode:

    chanel - trace setup - setup x/y max/min
    trigger - continuouse - trigger level - display type - trace -
    setup markers from gui
    trigger measure
    read AVG
    read source current

    ---
    INIT:CONT ON
    TRAC:STAT ON
    ? AVER:STAT OFF
    TRIG:SOUR INT1

    INIT - trigger new measure

    DISP:WIND1:TRAC:FEED "SENS1"
    DISP:WIND1:FORM TRAC
    DISP:SCR:FORM FSCR

    x start - SENS1:TRAC:OFFS:TIME -0.01 (secs)
    x scale - SENS1:TRAC:X:SCAL:PDIV 0.1 (seconds)

    y max - SENS1:TRAC:LIM:UPP -5 (dbm)
    y scale - SENS1:TRAC:Y:SCAL:PDIV 3 (dbm)

    trigger level - TRIG:SEQ:LEV -20 (dbm)

    set marker 1 - SENS1:SWE1:OFFS:TIME 0.00001
    set marker 2 - SENS1:SWE1:TIME 0.00002

    read marker: SENS1:SWE1:MARK1|2:POW?
    '''

    secondary = {
        'x_start': -0.0000002,
        'x_scale': 0.0000002,
        'y_max': 0.0,
        'y_scale': 2.0,
        'trig_level': -10.0,
        'mark_1': 0.0000002,
        'mark_2': 0.0000006,
        'avg': 2,
    }

    avg = secondary['avg']
    x_start = secondary['x_start']
    x_scale = secondary['x_scale']
    y_max = secondary['y_max']
    y_scale = secondary['y_scale']
    trig_level = secondary['trig_level']
    mark_1 = secondary['mark_1']
    mark_2 = secondary['mark_2']

    task = [
        (2.7, 13.63812149),
        (2.9, 13.13896211),
        (3.1, 13.317021930000001),
        (2.7, 15.675503699999998),
        (2.9, 15.1485758),
        (3.1, 15.307607),
        (2.7, 17.6418364),
        (2.9, 17.116042),
        (3.1, 17.34813),
        (2.7, 19.6842686),
        (2.9, 19.093634),
        (3.1, 19.3848347),
    ]

    gen.send('*RST')
    meter.send('*RST')

    meter.send(f'SENS1:AVER:COUN {avg}')
    meter.send('FORMat ASCII')

    meter.send('INIT:CONT ON')
    # meter.send('TRAC:STAT ON')
    meter.send('TRIG:SOUR INT1')

    meter.send('DISP:WIND1:TRAC:FEED "SENS1"')
    meter.send('DISP:WIND1:FORM TRAC')
    meter.send('DISP:SCR:FORM FSCR')

    meter.send(f'SENS1:TRAC:OFFS:TIME {x_start}')
    meter.send(f'SENS1:TRAC:X:SCAL:PDIV {x_scale}')
    meter.send(f'SENS1:TRAC:LIM:UPP {y_max}')
    meter.send(f'SENS1:TRAC:Y:SCAL:PDIV {y_scale}')

    meter.send(f'TRIG:SEQ:LEV {trig_level}')

    meter.send(f'SENS1:SWE1:OFFS:TIME {mark_1}')
    meter.send(f'SENS1:SWE1:TIME {mark_2}')

    # автоматическое измерение ошибается в первой точке, измеряем пустышку
    # почему - хз
    f1 = task[0][0] * GIGA
    p1 = task[0][1]
    gen.send(f'POW {p1}dbm')
    gen.send(f'FREQ {f1}')
    meter.send(f'SENS1:FREQ {f1}')
    gen.send('OUTP ON')
    time.sleep(0.1)
    meter.query('FETCH?')

    result = []
    for f, p in task:
        f = f * 1_000_000_000
        gen.send(f'POW {p}dbm')
        gen.send(f'FREQ {f}')
        meter.send(f'SENS1:FREQ {f}')
        gen.send('OUTP ON')

        # meter.send('ABORT')
        # meter.send('INIT')

        time.sleep(0.2)

        avg_pow = float(meter.query('FETCH?').strip())

        point = {
            'f': f,
            'p': p,
            'avg_pow': avg_pow,
        }
        print(point)

        result.append(point)

    gen.send('OUTP OFF')

    with open('out.txt', mode='wt', encoding='utf-8') as f:
        f.write(str(result))


if __name__ == '__main__':
    measure()
