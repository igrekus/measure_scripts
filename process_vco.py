import datetime
import os

import openpyxl
import pandas as pd

from openpyxl.chart import Reference, LineChart


def _fileter_xlsx(path):
    print('finding .xlsx')
    return sorted([os.path.normpath(f'{path}/{p}') for p in os.listdir(path) if p.endswith('.xlsx')])


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
    files = _fileter_xlsx(path)
    if not _is_file_list_valid(files):
        raise ValueError('Неполный список файлов, в папке должны быть:\n'
                         '- 3 файла с измерниями для пунктов 1,2,7,8,9,10,11,12 для всех температур\n'
                         '- 1 файл с измерениям для пуктов 3,4\n'
                         '- 3 файла с измерениями для пункта 6 для всех температур')

    for file in files:
        print(f'reading {file}')
        if 'vco-1-2' in file and file.endswith('+25.xlsx'):
            df_1_2_25 = pd.read_excel(file, engine='openpyxl')
        if 'vco-1-2' in file and file.endswith('-60.xlsx'):
            df_1_2_60 = pd.read_excel(file, engine='openpyxl')
        if 'vco-1-2' in file and file.endswith('+85.xlsx'):
            df_1_2_85 = pd.read_excel(file, engine='openpyxl')
        if 'vco-3-4' in file:
            df_3_4 = pd.read_excel(file, engine='openpyxl')
        if 'vco-6' in file and file.endswith('+25.xlsx'):
            df_6_25 = pd.read_excel(file, engine='openpyxl')
        if 'vco-6' in file and file.endswith('-60.xlsx'):
            df_6_60 = pd.read_excel(file, engine='openpyxl')
        if 'vco-6' in file and file.endswith('+85.xlsx'):
            df_6_85 = pd.read_excel(file, engine='openpyxl')

    print('building data array')
    # Usrc = 4.7V
    df_1_2_25.columns = [f'{c}+25' if 'Uc, V' not in c else c for c in df_1_2_25.columns]
    df_1_2_25 = df_1_2_25.drop(['Unnamed: 0+25'], axis=1)

    df_1_2_60.columns = [f'{c}-60' for c in df_1_2_60.columns]
    df_1_2_60 = df_1_2_60.drop(['Unnamed: 0-60', 'Uc, V-60'], axis=1)

    df_1_2_85.columns = [f'{c}+85' for c in df_1_2_85.columns]
    df_1_2_85 = df_1_2_85.drop(['Unnamed: 0+85', 'Uc, V+85'], axis=1)

    df_3_4 = df_3_4.drop(['Unnamed: 0', 'Uc, V'], axis=1)

    # Usrc = 5V
    df_6_25.columns = [f'{c}+25' for c in df_6_25.columns]
    df_6_25 = df_6_25.drop(['Unnamed: 0+25', 'Uc, V+25'], axis=1)
    df_6_25['dPx2, dBm+25'] = df_6_25.apply(lambda row: row['Px2, bDm+25'] - row['Px1, bDm+25'], axis=1)
    df_6_25['dPx3, dBm+25'] = df_6_25.apply(lambda row: row['Px3, bDm+25'] - row['Px1, bDm+25'], axis=1)

    df_6_60.columns = [f'{c}-60' for c in df_6_60.columns]
    df_6_60 = df_6_60.drop(['Unnamed: 0-60', 'Uc, V-60'], axis=1)
    df_6_60['dPx2, dBm-60'] = df_6_60.apply(lambda row: row['Px2, bDm-60'] - row['Px1, bDm-60'], axis=1)
    df_6_60['dPx3, dBm-60'] = df_6_60.apply(lambda row: row['Px3, bDm-60'] - row['Px1, bDm-60'], axis=1)

    df_6_85.columns = [f'{c}+85' for c in df_6_85.columns]
    df_6_85 = df_6_85.drop(['Unnamed: 0+85', 'Uc, V+85'], axis=1)
    df_6_85['dPx2, dBm+85'] = df_6_85.apply(lambda row: row['Px2, bDm+85'] - row['Px1, bDm+85'], axis=1)
    df_6_85['dPx3, dBm+85'] = df_6_85.apply(lambda row: row['Px3, bDm+85'] - row['Px1, bDm+85'], axis=1)

    result = pd.concat([df_1_2_25, df_1_2_60, df_1_2_85, df_3_4, df_6_25, df_6_60, df_6_85], axis=1)

    out_excel_name = f'vco-result{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx'
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
            Reference(ws, range_string=f'{ws.title}!C1:C{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!I1:I{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!O1:O{rows + 1}'),
        ],
        title='Диапазон рабочих частот',
        loc='B10'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!U1:W{rows + 1}'),
        ],
        title='Диапазон рабочих частот в зависимости от напряжения питания',
        loc='K10'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!H1:H{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!N1:N{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!T1:T{rows + 1}'),
        ],
        title='Крутизна регулировочной характеристики',
        loc='T10'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!D1:D{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!J1:J{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!P1:P{rows + 1}'),
        ],
        title='Выходная мощность',
        loc='B25'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!E1:E{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!K1:K{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!Q1:Q{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!F1:F{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!L1:L{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!R1:R{rows + 1}'),
        ],
        title='Мощность выходного сигнала паразитных гармоник (2й и 3й) в диапазоне температур)',
        loc='K25'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!G1:G{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!M1:M{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!S1:S{rows + 1}'),
        ],
        title='Ток потребления',
        loc='T25'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!AD1:AD{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AL1:AL{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AT1:AT{rows + 1}'),
        ],
        title='Относительный уровень 2й гармоники',
        loc='B40'
    )

    _add_chart(
        ws=ws,
        xs=Reference(ws, range_string=f'{ws.title}!B1:B{rows + 1}'),
        ys=[
            Reference(ws, range_string=f'{ws.title}!AE1:AE{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AM1:AM{rows + 1}'),
            Reference(ws, range_string=f'{ws.title}!AU1:AU{rows + 1}'),
        ],
        title='Относительный уровень 3й гармоники',
        loc='K40'
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
    path = os.path.normpath('ref/s2p_process/vco/xlsx')
    main(path)
