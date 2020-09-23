import visa

rm = visa.ResourceManager()

pna = rm.open_resource('TCPIP0::169.254.130.204::hislip0::INSTR')
src = rm.open_resource('USB0::0x0AAD::0x0135::039795547::0::INSTR')


def measure_pna():
    """ZNB 40 PNA"""

    src.write('APPLY 3.3v,50ma,1')
    src.write('OUTP ON')

    pna.write('CALC1:PAR:SDEF "Trc2", "S11"')
    pna.write('CALC1:PAR:SDEF "Trc3", "S22"')
    pna.write('CALC1:PAR:SDEF "Trc4", "S12"')

    """
    CALC2:PAR:SDEF 'Trc2', 'S11'
    DISP:WIND1:TRAC2:FEED 'Trc2'
    CALC:DATA:TRAC? 'Trc2', FDAT

    CALC3:PAR:SDEF 'Trc3', 'S22'
    DISP:WIND1:TRAC3:FEED 'Trc3'
    CALC:DATA:TRAC? 'Trc3', FDAT

    CALC3:PAR:SDEF 'Trc4', 'S12'
    DISP:WIND1:TRAC4:FEED 'Trc4'
    CALC:DATA:TRAC? 'Trc4', FDAT
    """

    res1 = pna.write('CALC:DATA:TRAC? "Trc2" FDAT')
    res2 = pna.write('CALC:DATA:TRAC? "Trc3" FDAT')

