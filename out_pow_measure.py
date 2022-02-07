import datetime
import time
import visa

import numpy as np


instruments = {
    'Генератор': 'ASRL6::INSTR',
    'Мощность': 'GPIB3::1::INSTR',
    'Источник': 'GPIB3::4::INSTR',
}

rm = visa.ResourceManager()

src = rm.open_resource(instruments['Источник'])
meter = rm.open_resource(instruments['Мощность'])
gen = rm.open_resource(instruments['Генератор'])

MHz = 1_000_000

secondary = {
    'f_start': 100 * MHz,
    'f_end': 300 * MHz,
    'f_step': 100 * MHz,
    'p_start': 0.0,
    'p_end': 10.0,
    'p_step': 5.0,
    'avg': 10,
}


src.send = src.write
gen.send = gen.write
meter.send = meter.write

def measure():
    '''
    pow measure:
    manual avg: SENSe1:AVERage:COUNt 10
    set freq: SENSe1:FREQuency 100000000

    gen:
    set freq: FREQuency 100000000
    set pow: pow 5dbm
    output on: output on

    '''
    avg = secondary['avg']
    f_start = secondary['f_start']
    f_end = secondary['f_end']
    f_step = secondary['f_step']
    p_start = secondary['p_start']
    p_end = secondary['p_end']
    p_step = secondary['p_step']

    pows = [round(x, 1) for x in np.arange(start=p_start, stop=p_end + 0.0001, step=p_step)]
    freqs = [round(x, 1) for x in np.arange(start=f_start, stop=f_end + 0.0001, step=f_step)]

    gen.send('*RST')
    meter.send('*RST')

    meter.send(f'SENS1:AVER:COUN {avg}')
    meter.send('FORMat ASCII')
    # meter.send('TRIG:SOUR INT1')
    # meter.send('INIT:CONT ON')

    task = [
        {'f': 100000000.0, 'p': 0.0, 'read_pow': 0.0228594364, 'delta_in': 0.164355169, 'delta_out': 0.3717806899999996},
        {'f': 200000000.0, 'p': 0.0, 'read_pow': 0.0140764248, 'delta_in': 0.215280425, 'delta_out': 0.24844625000000065},
        {'f': 300000000.0, 'p': 0.0, 'read_pow': 0.01597751, 'delta_in': 0.208624168, 'delta_out': 0.5548249199999997},
        {'f': 100000000.0, 'p': 5.0, 'read_pow': 5.00273474, 'delta_in': 0.19960721999999986, 'delta_out': 0.3717806899999996},
        {'f': 200000000.0, 'p': 5.0, 'read_pow': 5.00701105, 'delta_in': 0.22105469999999983, 'delta_out': 0.24844625000000065},
        {'f': 300000000.0, 'p': 5.0, 'read_pow': 4.9882887, 'delta_in': 0.11074699999999993, 'delta_out': 0.5548249199999997},
        {'f': 100000000.0, 'p': 10.0, 'read_pow': 10.0053209, 'delta_in': 0.18388404000000058, 'delta_out': 0.3717806899999996},
        {'f': 200000000.0, 'p': 10.0, 'read_pow': 9.99854645, 'delta_in': 0.21019315999999932, 'delta_out': 0.24844625000000065},
        {'f': 300000000.0, 'p': 10.0, 'read_pow': 9.99458571, 'delta_in': 0.06325619999999965, 'delta_out': 0.5548249199999997},
    ]

    point = task[0]
    # автоматическое измерение ошибается в первой точке, измеряем пустышку
    # почему - хз
    gen.send(f'POW {point["p"]}dbm')
    gen.send(f'FREQ {point["f"]}')
    meter.send(f'SENS1:FREQ {point["f"]}')
    gen.send('OUTP ON')
    meter.send('ABORT')
    meter.send('INIT')
    time.sleep(0.1)
    meter.query('FETCH?')

    result = []
    for row in task:
        p = row['p']
        f = row['f']
        delta_in = row['delta_in']
        delta_out = row['delta_out']

        gen.send(f'POW {p + delta_in}dbm')
        gen.send(f'FREQ {f}')
        meter.send(f'SENS1:FREQ {f}')
        gen.send('OUTP ON')

        meter.send('ABORT')
        meter.send('INIT')

        time.sleep(0.1)

        read_pow = float(meter.query('FETCH?'))
        adjusted_pow = read_pow + delta_out

        point = {
            'f': f,
            'p': p,
            'read_pow': read_pow,
            'adjusted_pow': adjusted_pow,
        }
        print(point)

        result.append(point)

    gen.send('OUTP OFF')


if __name__ == '__main__':
    measure()
