import datetime
import time
import visa

import numpy as np
import pandas as pd

from config import instruments

# --- параметры измерения ---
# питание
u_source = 5.0   # напряжение питания (В)
i_source = 100   # макс ток потребления (мА)

# найстройки генеретора
p_start = -70.0
p_end = 20.0
p_step = 2.0

f_start = 0.5   # GHz
f_end = 5.5   # GHz
f_step = 0.5   # GHz
# ---------------------------

rm = visa.ResourceManager()

src = rm.open_resource(instruments['источник питания 1'])
gen = rm.open_resource(instruments['генератор 1'])
mult = rm.open_resource(instruments['мультиметр'])

file_name = f'xlsx/log_detector-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'


def measure():
    fs = [0.1] + [round(x, 2) for x in np.linspace(f_start, f_end, int((f_end - f_start) / f_step) + 1, endpoint=True)]
    ps = [round(x, 2) for x in np.linspace(p_start, p_end, int((p_end - p_start) / p_step) + 1, endpoint=True)]

    src.write(f'APPLY {u_source}V,{i_source}ma,1')
    src.write('OUTP:CHAN1 ON')
    src.write('OUTP:MAST ON')

    result = []
    gen.write('OUTP ON')
    for f in fs:
        gen.write(f'SOUR:FREQ:CW {f}GHz')
        for p in ps:
            gen.write(f':POW:POW {p}dBm')
            print(f'set freq {f}GHz power {p}dB')

            time.sleep(1)

            u_out = mult.query('MEASURE:VOLTAGE:DC?')
            print('read u_out', u_out)
            result.append([f, p, float(u_out)])

    df = pd.DataFrame(result, columns=['F, GHz', f'P, dB', f'Uout, V'])
    print(df)

    df.to_excel(file_name)

    # wb = openpyxl.open(file_name)
    # ws = wb.active
    #
    # rows = len(df)
    # data = Reference(ws, range_string=f'{ws.title}!C1:C{rows + 1}')
    # xs = Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}')
    #
    # chart = LineChart()
    # chart.add_data(data, titles_from_data=True)
    # chart.set_categories(xs)
    #
    # ws.add_chart(chart, f'E4')
    #
    # wb.close()
    # wb.save(file_name)


if __name__ == '__main__':
    measure()
