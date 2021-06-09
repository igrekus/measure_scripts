import time
import visa

import numpy as np

instruments = {
    'Осциллограф': 'GPIB1::7::INSTR',
    'P LO': 'GPIB1::6::INSTR',
    'P RF': 'GPIB1::20::INSTR',
    'Источник': 'GPIB1::3::INSTR',
    'Мультиметр': 'GPIB1::2::INSTR',
    'Анализатор': 'GPIB1::18::INSTR',
}

rm = visa.ResourceManager()

gen_lo = rm.open_resource(instruments['P LO'])
gen_rf = rm.open_resource(instruments['P RF'])
osc = rm.open_resource(instruments['Осциллограф'])
src = rm.open_resource(instruments['Источник'])
mult = rm.open_resource(instruments['Мультиметр'])
sa = rm.open_resource(instruments['Анализатор'])


def measure_1():
    src_u = 5.0
    src_i = 200  # mA
    pow_lo_start = -10.0
    pow_lo_end = 0.0
    pow_lo_step = 5.0
    freq_lo_start = 0.09
    freq_lo_end = 5.99
    freq_lo_step = 0.1
    freq_lo_x2 = False

    pow_rf = -10.0
    freq_rf_start = 0.1
    freq_rf_end = 6.0
    freq_rf_step = 0.1

    loss = 0.82

    osc_avg = 'ON'

    osc_scale = 0.1
    osc_timebase_coeff = 1.0

    mock_enabled = False

    gen_lo.write('*RST')
    gen_rf.write('*RST')
    osc.write('*RST')
    src.write('*RST')
    mult.write('*RST')
    sa.write('*RST')

    src.write(f'APPLY p6v,{src_u}V,{src_i}mA')

    osc.write(f':ACQ:AVERage {osc_avg}')

    osc.write(f':CHANnel1:DISPlay ON')
    osc.write(f':CHANnel2:DISPlay ON')

    osc.write(':CHAN1:SCALE 0.05')  # V
    osc.write(':CHAN2:SCALE 0.05')
    osc.write(':TIMEBASE:SCALE 10E-8')  # ms / div

    osc.write(':TRIGger:MODE EDGE')
    osc.write(':TRIGger:EDGE:SOURCe CHANnel1')
    osc.write(':TRIGger:LEVel CHANnel1,0')
    osc.write(':TRIGger:EDGE:SLOPe POSitive')

    osc.write(':MEASure:VAMPlitude channel1')
    osc.write(':MEASure:VAMPlitude channel2')
    osc.write(':MEASure:PHASe CHANnel1,CHANnel2')
    osc.write(':MEASure:FREQuency CHANnel1')

    pow_lo_values = [round(x, 3) for x in np.arange(start=pow_lo_start, stop=pow_lo_end + 0.0001, step=pow_lo_step)] \
        if pow_lo_start != pow_lo_end else [pow_lo_start]
    freq_lo_values = [round(x, 3) for x in
                      np.arange(start=freq_lo_start, stop=freq_lo_end + 0.0001, step=freq_lo_step)]
    freq_rf_values = [round(x, 3) for x in
                      np.arange(start=freq_rf_start, stop=freq_rf_end + 0.0001, step=freq_rf_step)]

    gen_lo.write(f':OUTP:MOD:STAT OFF')
    # gen_rf.write(f':OUTP:MOD:STAT OFF')

    low_signal_threshold = 1.1
    range_ratio = 1.2

    res = []
    for pow_lo in pow_lo_values:

        for freq_lo, freq_rf in zip(freq_lo_values, freq_rf_values):

            if freq_lo_x2:
                freq_lo *= 2

            gen_lo.write(f'SOUR:POW {round(pow_lo, 2)}dbm')
            gen_rf.write(f'SOUR:POW {round(pow_rf, 2)}dbm')

            gen_lo.write(f'SOUR:FREQ {freq_lo}GHz')
            gen_rf.write(f'SOUR:FREQ {freq_rf}GHz')

            # TODO hoist out of the loops
            src.write('OUTPut ON')

            gen_lo.write(f'OUTP:STAT ON')
            gen_rf.write(f'OUTP:STAT ON')

            osc.write(':CDISplay')

            # time.sleep(0.5)
            if not mock_enabled:
                time.sleep(2)

            stats = osc.query(':MEASure:RESults?')

            stats_split = stats.split(',')
            osc_ch1_amp = float(stats_split[18])
            osc_ch2_amp = float(stats_split[25])
            osc_phase = float(stats_split[11])
            osc_ch1_freq = float(stats_split[4])

            timebase = (1 / (abs(freq_rf - ((freq_lo / 2) if freq_lo_x2 else freq_lo)) * 10_000_000)) * 0.01 * osc_timebase_coeff
            osc.write(f':TIMEBASE:SCALE {timebase}')  # ms / div
            osc.write(f':CHANnel1:OFFSet 0')
            osc.write(f':CHANnel2:OFFSet 0')

            max_amp = osc_ch1_amp if osc_ch1_amp > osc_ch2_amp else osc_ch2_amp

            if not mock_enabled:
                # check if auto-scale is needed:
                # some of the measure points go out of OSC display range resulting in incorrect measurement
                # this is correct external device behaviour, not a program bug
                if max_amp < 1_000_000:
                    # if reading is correct, check if the signal is too small
                    big_amp, ch_num = (osc_ch1_amp, 1) if osc_ch1_amp > osc_ch2_amp else (osc_ch2_amp, 2)
                    current_scale = float(osc.query(f':CHAN{ch_num}:SCALE?'))

                    # if signal fits in less than 1.5 sections of the display, is is too small, need to
                    # auto scale OSC display up
                    while big_amp / current_scale <= low_signal_threshold:

                        target_range = big_amp + big_amp * range_ratio

                        osc.write(f':CHANnel1:RANGe {target_range}')
                        osc.write(f':CHANnel2:RANGe {target_range}')

                        osc.write(':CDIS')

                        time.sleep(2)

                        autofit_stats_split = osc.query(':MEASure:RESults?').split(',')

                        osc_ch1_amp = float(autofit_stats_split[18])
                        osc_ch2_amp = float(autofit_stats_split[25])

                        big_amp, ch_num = (osc_ch1_amp, 1) if osc_ch1_amp > osc_ch2_amp else (osc_ch2_amp, 2)
                        current_scale = float(osc.query(f':CHAN{ch_num}:SCALE?'))
                else:
                    # if reading was not correct, reset OSC display range to safe level (controlled via GUI)
                    # and iterate OSC range scaling a few times
                    # to get the correct reading
                    max_amp = osc_ch1_amp if osc_ch1_amp > osc_ch2_amp else osc_ch2_amp
                    if max_amp > 1_000_000:

                        osc.write(f':CHANnel1:RANGe {osc_scale}')
                        osc.write(f':CHANnel2:RANGe {osc_scale}')

                        osc.write(':CDIS')

                        time.sleep(2)

                        autofit_stats_split = osc.query(':MEASure:RESults?').split(',')
                        osc_ch1_amp = float(autofit_stats_split[18])
                        osc_ch2_amp = float(autofit_stats_split[25])

                        # check if safe level results in too small signal
                        big_amp, ch_num = (osc_ch1_amp, 1) if osc_ch1_amp > osc_ch2_amp else (osc_ch2_amp, 2)
                        current_scale = float(osc.query(f':CHAN{ch_num}:SCALE?'))

                        # if signal fits in less than 1.5 sections of the display, auto scale OSC display up
                        while big_amp / current_scale <= low_signal_threshold:

                            if token.cancelled:
                                gen_lo.write(f'OUTP:STAT OFF')
                                gen_rf.write(f'OUTP:STAT OFF')
                                time.sleep(0.5)
                                src.write('OUTPut OFF')

                                gen_rf.write(f'SOUR:POW {pow_rf}dbm')
                                gen_lo.write(f'SOUR:POW {pow_lo_start}dbm')

                                gen_rf.write(f'SOUR:FREQ {freq_rf_start}GHz')
                                gen_lo.write(f'SOUR:FREQ {freq_rf_start}GHz')
                                raise RuntimeError('measurement cancelled')

                            target_range = big_amp + big_amp * range_ratio

                            osc.write(f':CHANnel1:RANGe {target_range}')
                            osc.write(f':CHANnel2:RANGe {target_range}')

                            osc.write(':CDIS')

                            time.sleep(2)

                            autofit_stats_split = osc.query(':MEASure:RESults?').split(',')

                            osc_ch1_amp = float(autofit_stats_split[18])
                            osc_ch2_amp = float(autofit_stats_split[25])

                            big_amp, ch_num = (osc_ch1_amp, 1) if osc_ch1_amp > osc_ch2_amp else (osc_ch2_amp, 2)
                            current_scale = float(osc.query(f':CHAN{ch_num}:SCALE?'))
                        else:
                            # if safe level is acceptable, select largest signal
                            # and fit the display to 130% of the signal
                            max_amp = osc_ch1_amp if osc_ch1_amp > osc_ch2_amp else osc_ch2_amp
                            target_range = max_amp * range_ratio
                            if target_range < 1_000_000:
                                osc.write(f':CHANnel1:RANGe {target_range}')
                                osc.write(f':CHANnel2:RANGe {target_range}')

            # read actual amp values after auto-scale (if any occured)
            osc.write(':CDIS')

            time.sleep(2)

            stats_split = osc.query(':MEASure:RESults?').split(',')
            osc_ch1_amp = float(stats_split[18])
            osc_ch2_amp = float(stats_split[25])

            f_lo_read = float(gen_lo.query('SOUR:FREQ?'))
            f_rf_read = float(gen_rf.query('SOUR:FREQ?'))

            i_src_read = float(mult.query('MEAS:CURR:DC? 1A,DEF'))

            raw_point = {
                'p_lo': pow_lo,
                'f_lo': f_lo_read,
                'p_rf': pow_rf,
                'f_rf': f_rf_read,
                'u_src': src_u,   # power source voltage
                'i_src': i_src_read,
                'ch1_amp': osc_ch1_amp,
                'ch2_amp': osc_ch2_amp,
                'phase': osc_phase,
                'ch1_freq': osc_ch1_freq,
                'loss': loss,
            }


            print(raw_point, stats)
            res.append([raw_point, stats])

    gen_lo.write(f'OUTP:STAT OFF')
    gen_rf.write(f'OUTP:STAT OFF')
    time.sleep(0.5)
    src.write('OUTPut OFF')

    gen_rf.write(f'SOUR:POW {pow_rf}dbm')
    gen_lo.write(f'SOUR:POW {pow_lo_start}dbm')

    gen_rf.write(f'SOUR:FREQ {freq_rf_start}GHz')
    gen_lo.write(f'SOUR:FREQ {freq_rf_start}GHz')


if __name__ == '__main__':
    measure_1()


"""

"""
