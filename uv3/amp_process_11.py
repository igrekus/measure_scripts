import io
import pandas as pd

from math import log10

csvs = [
    '9-0.01.csv',
    '9-1.csv',
    '9-1.5.csv',
]
path = 'xlsx'

amp = 25   # mA
volt = 3.204   # V


def _calc_pae(row):
    pow_in, pow_out = row
    return 100 * (pow_out - pow_in) / (10 * log10(volt * amp * 1_000))


for csv in csvs:
    file_path = f'{path}/{csv}'
    print('process', file_path)
    with open(file_path, mode='rt', encoding='utf-8') as f:
        df = pd.read_csv(io.StringIO('\n'.join(f.readlines()[2:]).replace(',', '.')), sep=';')
        df = df.drop(['Unnamed: 2'], axis=1)
        df.applymap(lambda cell: float(cell))
        df['PAE %'] = df.apply(_calc_pae, axis=1)
        df.to_excel(csv.replace('.csv', '.xlsx').replace('9-', '11-'))
