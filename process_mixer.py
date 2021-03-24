import datetime
import io
import os

import openpyxl
import pandas as pd

from openpyxl.chart import Reference, LineChart


# todo 1-33 -- usb (lf+lf2)/2
# todo 34-66 -- lsb -||-
# todo 67-end -- 1-33, only lf1, lsb
# todo 1-33, 34-66  --

def _fileter_xlsx(path):
    print('finding .csv')
    return sorted([os.path.normpath(f'{path}/{p}') for p in os.listdir(path) if p.endswith('.csv')])


def _is_file_list_valid(files):
    print('validating file list')
    res = []
    for file in files:
        if any(ref in file for ref in ['1-10.csv', '1-100.csv', '1-1500.csv']):
            res.append(file)
        if any(ref in file for ref in ['3.csv', '4.csv', '5.csv']):
            res.append(file)
    return files == res and len(files) == 6


def _read_csv(file):
    with open(file, mode='rt', encoding='utf-8') as f:
        return pd.read_csv(io.StringIO('\n'.join(f.readlines()[2:])), sep=';').drop(['Unnamed: 2'], axis=1)


def main(path):
    files = _fileter_xlsx(path)
    if not _is_file_list_valid(files):
        raise ValueError('Неполный список файлов, в папке должны быть:\n'
                         '- 3 файла с измерниями для пунктов 1,2,7,8,9,10,11,12 для всех температур\n'
                         '- 1 файл с измерениям для пуктов 3,4\n'
                         '- 3 файла с измерениями для пункта 6 для всех температур')

    for file in files:
        if any(ref in file for ref in ['1-10.csv', '1-100.csv', '1-1500.csv']):
            if '1-10.csv' in file:
                df_1_10 = _read_csv(file)
            if '1-100.csv' in file:
                df_1_100 = _read_csv(file)
            if '1-1500.csv' in file:
                df_1_1500 = _read_csv(file)
        if any(ref in file for ref in ['3.csv', '4.csv', '5.csv']):
            if '3.csv' in file:
                df_3 = _read_csv(file)
            if '4.csv' in file:
                df_4 = _read_csv(file)
            if '5.csv' in file:
                df_5 = _read_csv(file)

    print('building data array')
    df_1_10.columns = [f'{c}-10' if c != 'freq[Hz]' else 'freq[GHz]' for c in df_1_10.columns]

    df_1_100 = df_1_100.drop(['freq[Hz]'], axis=1)
    df_1_100.columns = [f'{c}-100' for c in df_1_100.columns]

    df_1_1500 = df_1_1500.drop(['freq[Hz]'], axis=1)
    df_1_1500.columns = [f'{c}-1500' for c in df_1_1500.columns]

    df_3 = df_3.drop(['freq[Hz]'], axis=1)
    df_3.columns = [f'{c}+25' for c in df_3.columns]

    df_4 = df_4.drop(['freq[Hz]'], axis=1)
    df_4.columns = [f'{c}-60' for c in df_4.columns]

    df_5 = df_5.drop(['freq[Hz]'], axis=1)
    df_5.columns = [f'{c}+85' for c in df_5.columns]

    result = pd.concat([
        df_1_10, df_1_100, df_1_1500,
        df_3, df_4, df_5,
    ], axis=1)

    result = result.applymap(lambda x: float(x.replace(',', '.')))
    result['freq[GHz]'] = result['freq[GHz]'].apply(lambda x: x / 1_000_000_000)
    print(result)

    out_excel_name = f'mixer-result{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
    result.to_excel(out_excel_name, engine='openpyxl')

    # save plots
    print('making charts')
    wb = openpyxl.open(out_excel_name)
    ws = wb.active

    rows = len(result)

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!C1:E{rows + 1}'),
        ],
        title='Зависимость коэффициента преобразования от частоты входного сигнала при разной частоте ПЧ',
        loc='B10'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!F1:H{rows + 1}'),
        ],
        title='Зависимость коэффициента преобразования от частоты входного сигнала в диапазоне температур',
        loc='K10'
    )

    print(f'saving resulting {out_excel_name}')
    wb.save(out_excel_name)
    wb.close()


def _add_chart(ws, xs, ys, title, loc):
    chart = LineChart()
    for y in ys:
        chart.add_data(y, titles_from_data=True)
    chart.set_categories(xs)
    chart.title = title
    ws.add_chart(chart, loc)


if __name__ == '__main__':
    # path = sys.argv
    path = os.path.normpath('ref/s2p_process/mixer/lsb')
    main(path)
