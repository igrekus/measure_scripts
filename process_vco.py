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

    # abs harmonic + pow - 1 = delta Px2
    cols_px1 = [f'({t}:1par) Px1, bDm' for t in temperatures]
    cols_px2 = [f'({t}:1par) Px2, bDm' for t in temperatures]
    cols_px3 = [f'({t}:1par) Px3, bDm' for t in temperatures]

    for t, px1, px2 in zip(temperatures, cols_px1, cols_px2):
        result[f'({t}) Rel Px2'] = result.apply(lambda row: -(abs(row[px2]) + row[px1]), axis=1)

    for t, px1, px3 in zip(temperatures, cols_px1, cols_px3):
        result[f'({t}) Rel Px3'] = result.apply(lambda row: -(abs(row[px3]) + row[px1]), axis=1)

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

    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!H1:H{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!W1:W{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AL1:AL{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BA1:BA{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BP1:BP{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CE1:CE{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CT1:CT{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DI1:DI{rows + 1}'),
        ],
        title='Крутизна, МГц/В',
        loc='K24'
    )

    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!Q1:Q{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AF1:AF{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AU1:AU{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BJ1:BJ{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BY1:BY{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CN1:CN{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DC1:DC{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DR1:DR{rows + 1}'),
        ],
        title='Крутизна, МГц/В',
        loc='T24'
    )

    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!D1:D{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!S1:S{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AH1:AH{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AW1:AW{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BL1:BL{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CA1:CA{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CP1:CP{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DE1:DE{rows + 1}'),
        ],
        title='Pвых, дБм',
        loc='B39'
    )

    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!E1:E{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!T1:T{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AI1:AI{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AX1:AX{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BM1:BM{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CB1:CB{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CQ1:CQ{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DF1:DF{rows + 1}'),

            Reference(ws, range_string=f'{ws.title}!F1:F{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!U1:U{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AJ1:AJ{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AY1:AY{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BN1:BN{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CC1:CC{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CR1:CR{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DG1:DG{rows + 1}'),
        ],
        title='Pвых гармоник, дБм',
        loc='K39'
    )

    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!G1:G{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!V1:V{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AK1:AK{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AZ1:AZ{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!BO1:BO{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CD1:CD{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!CS1:CS{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DH1:DH{rows + 1}'),
        ],
        title='Ток потребления, мА',
        loc='T39'
    )

    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!DS1:DS{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DT1:DT{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DU1:DU{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DV1:DV{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DW1:DW{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DX1:DX{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DY1:DY{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!DZ1:DZ{rows + 1}'),
        ],
        title='Относ. ур. 2й гарм., дБм',
        loc='B54'
    )

    _add_chart(
        ws=ws,
        xs=xs,
        ys=[
            Reference(ws, range_string=f'{ws.title}!EA1:EA{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!EB1:EB{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!EC1:EC{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!ED1:ED{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!EE1:EE{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!EF1:EF{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!EG1:EG{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!EH1:EH{rows + 1}'),
        ],
        title='Относ. ур. 3й гарм., дБм',
        loc='K54'
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
