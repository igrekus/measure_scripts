import time
import visa

rm = visa.ResourceManager()

gen = rm.open_resource('TCPIP0::169.254.2.20::hislip0::INSTR')
sa = rm.open_resource('TCPIP0::169.254.56.102::hislip0::INSTR')
src = rm.open_resource('USB0::0x0AAD::0x0135::039964506::0::INSTR')

"""
source - set params - APPLY 3.3v,50ma,1
source - measure current - MEAS:CURR?
source - turn on output - OUTP:CHAN1 ON 
                          OUTP:MAST ON

gen - turn on output - OUTP ON
gen - set freq - SOURce1:FREQuency:CW 100mhz
gen - set power - :POW:POW 0.1dbm

sa - set start freq - frequency:start 20mhz
sa - set stop freq - frequency:stop 20ghz
sa - turn on marker - calc1:mark1 on 
sa - turn on marker - calc1:mark1:x 200mhz
sa - display reference - level DISP:TRAC1:Y:RLEV -5dbm
sa - switch to phase noise - INST:SEL PNOISE
ss - auto adjust - ADJust:ALL

pna - get trace date - TRAC?
"""
