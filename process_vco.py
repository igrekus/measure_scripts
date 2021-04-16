import datetime
import os

from collections import defaultdict

import openpyxl
import pandas as pd

from openpyxl.chart import Reference, LineChart


temperatures = [
    '-20',
    '-40',
    '-60',
    '+25',
    '+40',
    '+60',
    '+85',
    '0',
]


def _find_files(path):
    print('finding .xlsx')
    return {
        t: _filter_xlsx(os.path.normpath(os.path.join(path, t)))
        for t in temperatures
    }


def _filter_xlsx(path):
    all_files = [p for p in os.listdir(path) if p.endswith('.xlsx')]

    if len(all_files) != 3:
        raise ValueError(f'в папке {path} должно быть три .xlsx-файла')

    f1, f2, f3 = all_files
    if not f1.startswith('vco-1-3-4-5par-'):
        raise ValueError(f'в папке найден файл с неправильным форматом названия {f1}')
    if not f2.startswith('vco-2par-'):
        raise ValueError(f'в папке найден файл с неправильным форматом названия {f2}')
    if not f3.startswith('vco-6tok-'):
        raise ValueError(f'в папке найден файл с неправильным форматом названия {f2}')

    return sorted([os.path.normpath(f'{path}/{p}') for p in all_files])


def _is_file_list_valid(files):
    print('validating file list')
    res = []
    for file in files:
        if 'vco-1-2' in file:
            if file.endswith('+25.xlsx'):
                res.append(file)
            if file.endswith('-60.xlsx'):
                res.append(file)
            if file.endswith('+85.xlsx'):
                res.append(file)
        elif 'vco-3-4' in file:
            res.append(file)
        elif 'vco-6' in file:
            if file.endswith('+25.xlsx'):
                res.append(file)
            if file.endswith('-60.xlsx'):
                res.append(file)
            if file.endswith('+85.xlsx'):
                res.append(file)
    return files == res and len(files) == 7


def main(path):
    all_files = _find_files(path)
    prefixes = {
        0: '1par',
        1: '2par',
        2: '6par',
    }

    dfs = []

    for temp, files in list(all_files.items()):
        for idx, file in enumerate(files):
            df = pd.read_excel(file, engine='openpyxl', index_col=None)
            df = df.drop(['Unnamed: 0'], axis=1)
            if idx == 1:
                df = df.pivot(index='Uc, V', columns='Usrc, V', values='F, Mhz')
                df.columns = [f'F@Uc={c} V' for c in df.columns]
                df['Uc, V'] = df.index
                cs = list(df.columns)
                df = df[[cs[-1]] + cs[:-1]]
                df.reset_index(drop=True, inplace=True)

            pref = prefixes[idx]
            df.columns = [df.columns[0]] + [f'({temp}:{pref}) {c}' for c in df.columns[1:]]

            dfs.append(df)

    for i in range(1, len(dfs)):
        dfs[i] = dfs[i].drop(dfs[i].columns[:1], axis=1)

    print('building data array')

    out_excel_name = f'vco-result{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'

    result = pd.concat(objs=dfs, axis=1)
    result.to_excel(out_excel_name, engine='openpyxl')

    # save plots
    print('making charts')
    wb = openpyxl.open(out_excel_name)
    ws = wb.active

    rows = len(result)
    xs = Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}')
    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!C1:C{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!R1:R{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AG1:AG{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AV1:AV{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BK1:BK{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BZ1:BZ{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CO1:CO{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DD1:DD{rows + 1}'),
        ],
        title='Частота, МГц',
        loc='B24'
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
    path = os.path.normpath(r'.\ref\s2p_process\vco\2021-04-14')
    main(path)
