import datetime
import openpyxl
import time
import visa

import numpy as np
import pandas as pd

from openpyxl.chart import LineChart, Reference

instruments = {
    'Осциллограф': 'GPIB1::7::INSTR',
    'P LO': 'GPIB1::6::INSTR',
    'P RF': 'GPIB1::20::INSTR',
    'Источник': 'GPIB1::3::INSTR',
    'Мультиметр': 'GPIB1::23::INSTR',
    'Анализатор': 'GPIB1::18::INSTR',
}

# --- параметры измерения ---
# генератор
f_start = 0.45   # начальная частота диапазона (ГГц)
f_end = 15.0   # конечная частота диапазона (ГГЦ)
f_step = 0.1   # шаг перестройки частоты (ГГц)
p_in = 0   # мощность дБм

# исочник питания
u_source = 5.0   # напряжение питания (В)
i_source = 100   # макс ток потребления (мА)

# коэфициент деления
coeff = 2
# ---------------------------

rm = visa.ResourceManager()

gen_lo = rm.open_resource(instruments['P LO'])
gen_rf = rm.open_resource(instruments['P RF'])
osc = rm.open_resource(instruments['Осциллограф'])
src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['Анализатор'])

file_name = f'xlsx/divisor-1-coeff_{coeff}-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure_1():
    gen_lo.write('*RST')
    gen_rf.write('*RST')
    osc.write('*RST')   # TODO doesnt know reset
    src.write('*RST')
    mult.write('*RST')
    sa.write('*RST')

    secondary = {'Plo_min': -10.0, 'Plo_max': 10.0, 'Plo_delta': 1.0, 'Flo_min': 1.0, 'Flo_max': 3.0,
                 'Flo_delta': 0.1, 'Prf': -10.0, 'Frf_min': 1.0, 'Frf_max': 3.0, 'Frf_delta': 0.1, 'Usrc': 5.0,
                 'OscAvg': False, 'Loss': 0.82}

    src_u = 5
    src_i = 200   # mA
    pow_lo_start = -10
    pow_lo_end = -5
    pow_lo_step = 5
    freq_lo_start = 0.1   # GHz
    freq_lo_end = 3
    freq_lo_step = 0.1

    pow_rf = -10
    freq_rf_start = 0.11
    freq_rf_end = 3
    freq_rf_step = 0.1

    osc_avg = 'ON'

    src.write(f'APPLY p6v,{src_u}V,{src_i}mA')

    osc.write(f':ACQ:AVERage {osc_avg}')

    osc.write(f':CHANnel1:DISPlay ON')
    osc.write(f':CHANnel2:DISPlay ON')

    osc.write(':CHAN1:SCALE 0.05')   # V
    osc.write(':CHAN2:SCALE 0.05')
    osc.write(':TIMEBASE:SCALE 10E-8')   # ms / div

    osc.write(':TRIGger:MODE EDGE')
    osc.write(':TRIGger:EDGE:SOURCe CHANnel1')
    osc.write(':TRIGger:LEVel CHANnel1,0')
    osc.write(':TRIGger:EDGE:SLOPe POSitive')

    osc.write(':MEASure:VAMPlitude channel1')
    osc.write(':MEASure:VAMPlitude channel2')
    osc.write(':MEASure:PHASe CHANnel1,CHANnel2')
    osc.write(':MEASure:FREQuency CHANnel1')

    # osc_ch1_amp = float(osc.query(':MEASure:VAMPlitude? channel1'))
    # osc.write(f':CHANnel1:RANGe {osc_ch1_amp + 0.3 * osc_ch1_amp}')

    pow_lo_values = [round(x, 3) for x in np.arange(start=pow_lo_start, stop=pow_lo_end + 0.2, step=pow_lo_step)]
    freq_lo_values = [round(x, 3) for x in np.arange(start=freq_lo_start, stop=freq_lo_end + 0.2, step=freq_lo_step)]
    freq_rf_values = [round(x, 3) for x in np.arange(start=freq_rf_start, stop=freq_rf_end + 0.2, step=freq_rf_step)]

    gen_rf.write(f'SOUR:POW {pow_rf}dbm')

    res = []
    for pow_lo in pow_lo_values:
        gen_lo.write(f'SOUR:POW {pow_lo}dbm')

        for freq_lo, freq_rf in zip(freq_lo_values, freq_rf_values):
            # if token.cancelled:
            #     raise RuntimeError('measurement cancelled')

            gen_lo.write(f'SOUR:FREQ {freq_lo}GHz')
            gen_rf.write(f'SOUR:FREQ {freq_rf}GHz')

            # TODO hoist out of the loops
            src.write('OUTPut ON')

            gen_lo.write(f'OUTP:STAT ON')
            gen_rf.write(f'OUTP:STAT ON')

            # osc.write(':MEASure:CLEar')
            # osc.write(':MEASure:VAMPlitude channel1')
            # osc.write(':MEASure:VAMPlitude channel2')
            # osc.write(':MEASure:PHASe CHANnel1,CHANnel2')
            # osc.write(':MEASure:FREQuency CHANnel1')
            osc.write(':CDISplay')

            time.sleep(3)
            # if not mock_enabled:
            #     time.sleep(3)

            # label, current, result(-), min, max, mean, std dev, times
            # TODO read mean instead of current
            stats = osc.query(':MEASure:RESults?')

            osc_ch1_amp = float(osc.query(':MEASure:VAMPlitude? channel1'))
            osc_ch2_amp = float(osc.query(':MEASure:VAMPlitude? channel2'))
            osc_phase = float(osc.query(':MEASure:PHASe? CHANnel1,CHANnel2'))
            osc_ch1_freq = float(osc.query(':MEASure:FREQuency? CHANnel1'))

            timebase = (1 / (abs(freq_rf - freq_lo) * 10_000_000)) * 0.01
            osc.write(f':TIMEBASE:SCALE {timebase}')   # ms / div
            osc.write(f':CHANnel1:OFFSet 0')
            osc.write(f':CHANnel2:OFFSet 0')

            rng = osc_ch1_amp + 0.3 * osc_ch1_amp
            osc.write(f':CHANnel1:RANGe {rng}')
            osc.write(f':CHANnel2:RANGe {rng}')

            while rng > 1_000_000:
                osc.write(f':CHANnel1:RANGe 0.2')
                osc.write(f':CHANnel2:RANGe 0.2')

                osc.write(':CDIS')

                time.sleep(2)
                osc_ch1_amp = float(osc.query(':MEASure:VAMPlitude? channel1'))
                rng = osc_ch1_amp + 0.3 * osc_ch1_amp
                osc.write(f':CHANnel1:RANGe {rng}')
                osc.write(f':CHANnel2:RANGe {rng}')

            # TODO record live data
            p_lo_read = float(gen_lo.query('SOUR:POW?'))
            f_lo_read = float(gen_lo.query('SOUR:FREQ?'))

            p_rf_read = float(gen_rf.query('SOUR:POW?'))
            f_rf_read = float(gen_rf.query('SOUR:FREQ?'))

            u_src_read = float(src.query('MEAS:VOLT?'))
            i_src_read = float(src.query('MEAS:CURR?'))   # TODO read measured values

            # u_src_read = random.randint(1, 100)

            raw_point = {
                'p_lo': p_lo_read,
                'f_lo': f_lo_read,
                'p_rf': p_rf_read,
                'f_rf': f_rf_read,
                'u_src': u_src_read,
                'i_src': i_src_read,
                'ch1_amp': osc_ch1_amp,
                'ch2_amp': osc_ch2_amp,
                'phase': osc_phase,
                'ch1_freq': osc_ch1_freq,
            }
            # to show:
            # + p_lo, f_lo
            # + p_rf, f_rf
            # + u_src, i_src
            # + ch1_amp. ch2_amp, ch2_amp - ch1_amp, phase, ch1_freq

            # расчеты:
            # мощность сигнала пч по каналу 1: Ppch = 30 + 1 * log10(((ch1_amp/2 * 0.001) ^ 2) / 100)
            # к-т передачи с учетом потерь: Kp = Ppch - Prf + Pbal
            # амп. ошибка в разах: aerr_times = ch2_amp / chi1_amp
            # амп. ош в дБ: aerr_db = 20 * log10(ch2_amp * 0.001) - 20 * log10(ch1_amp * 0.001)
            # фаз. ош в град: pherr = delta_pherr + 90
            # подавление зерк. канала: azk = 10 * log10((1 + aerr_times ^ 2 - 2 * aerr_times * cos(rad(pherr)))
            # / (1 + aerr_times ^ 2 + 2 * aerr_times * cos(rad(pherr))))

            print(raw_point)
            # self._add_measure_point(raw_point)

            res.append([raw_point, stats])

    with open('out.txt', mode='wt', encoding='utf-8') as f:
        f.write(str(res))

    return res


if __name__ == '__main__':
    measure_1()


"""

"""
