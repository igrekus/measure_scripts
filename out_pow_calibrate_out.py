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

    calibrated_in = [
        {'f': 100000000.0, 'p': 0.0, 'read_pow': 0.0228594364, 'delta': 0.164355169},
        {'f': 200000000.0, 'p': 0.0, 'read_pow': 0.0140764248, 'delta': 0.215280425},
        {'f': 300000000.0, 'p': 0.0, 'read_pow': 0.01597751, 'delta': 0.208624168},
        {'f': 100000000.0, 'p': 5.0, 'read_pow': 5.00273474, 'delta': 0.19960721999999986},
        {'f': 200000000.0, 'p': 5.0, 'read_pow': 5.00701105, 'delta': 0.22105469999999983},
        {'f': 300000000.0, 'p': 5.0, 'read_pow': 4.9882887, 'delta': 0.11074699999999993},
        {'f': 100000000.0, 'p': 10.0, 'read_pow': 10.0053209, 'delta': 0.18388404000000058},
        {'f': 200000000.0, 'p': 10.0, 'read_pow': 9.99854645, 'delta': 0.21019315999999932},
        {'f': 300000000.0, 'p': 10.0, 'read_pow': 9.99458571, 'delta': 0.06325619999999965},
    ]
    max_p = max(el['p'] for el in calibrated_in)
    calibrated_in = list(filter(lambda el: el['p'] == max_p, calibrated_in))
    point = calibrated_in[0]

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
    for point in calibrated_in:
        p = point['p']
        f = point['f']

        gen.send(f'POW {p}dbm')
        gen.send(f'FREQ {f}')
        meter.send(f'SENS1:FREQ {f}')
        gen.send('OUTP ON')

        meter.send('ABORT')
        meter.send('INIT')

        time.sleep(0.1)

        read_pow = float(meter.query('FETCH?'))
        delta = p - read_pow

        point = {
            'f': f,
            'p': p,
            'read_pow': read_pow,
            'delta': delta,
        }
        print(point)

        result.append(point)

    # {'f': 100000000.0, 'p': 10.0, 'read_pow': 9.62821931, 'delta': 0.3717806899999996}
    # {'f': 200000000.0, 'p': 10.0, 'read_pow': 9.75155375, 'delta': 0.24844625000000065}
    # {'f': 300000000.0, 'p': 10.0, 'read_pow': 9.44517508, 'delta': 0.5548249199999997}

    gen.send('OUTP OFF')


if __name__ == '__main__':
    measure()
