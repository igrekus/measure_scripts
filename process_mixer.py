import datetime
import io
import os

import pandas as pd

from openpyxl.chart import LineChart



csvs = [
    '1-10.csv',
    '1-100.csv',
    '1-1500.csv',
    '10.csv',
    '11-10.csv',
    '11-100.csv',
    '11-1500.csv',
    '12-12.csv',
    '12-14.csv',
    '12-16.csv',
    '12-18.csv',
    '13.csv',
    '14.csv',
    '15.csv',
    '2-12.csv',
    '2-14.csv',
    '2-16.csv',
    '2-18.csv',
    '26-12.csv',
    '26-14.csv',
    '26-16.csv',
    '26-18.csv',
    '27.csv',
    '28.csv',
    '29.csv',
    '3.csv',
    '30.csv',
    '31.csv',
    '32.csv',
    '33.csv',
    '4.csv',
    '5.csv',
    '6-5.5.csv',
    '6-6.5.csv',
    '6-7.5.csv',
    '7-12.csv',
    '7-14.csv',
    '7-16.csv',
    '7-18.csv',
    '8.csv',
    '9.csv',
]


def _build_file_list(path):
    print('finding .csv')
    root = path

    return {lf: {temp: {band: [
        os.path.normpath(f'{root}/{lf} {temp}/{band}/{c}') for c in csvs
    ] for band in ['lsb', 'usb']} for temp in ['+25', '+85', '-60']} for lf in ['LF1', 'LF2']}


def _validate_file_list(file_dict):
    print('validating file list')
    res = []
    for lf, temps in file_dict.items():
        for temp, bands in temps.items():
            for band, files in bands.items():
                for file in files:
                    if not os.path.isfile(file):
                        raise ValueError(f'нет .csv для LF={lf}, темп={temp}, полоса={band}: не хватает файла: {file}')


def _read_csvs_to_dfs(file_dict):
    out = {}
    for lf, temps in file_dict.items():
        tps = {}
        for temp, bands in temps.items():
            bds = {}
            for band, files in bands.items():
                dfs = []
                for idx, file in enumerate(files):
                    df = _read_csv(file)
                    dfs.append(df)
                    # dfs.append([file.split('\\')[-1], idx])
                bds[band] = dfs
            tps[temp] = bds
        out[lf] = tps
    return out


def _read_csv(file):
    with open(file, mode='rt', encoding='utf-8') as f:
        return pd.read_csv(io.StringIO('\n'.join(f.readlines()[2:])), sep=';').drop(['Unnamed: 2'], axis=1)


def _build_1_cat_df(dfs):
    # 1 cat - 1-5, 11-15, 30-33
    filtered = [
        dfs['LF1']['+25']['usb'][0],   # 1 - B-G
        dfs['LF1']['+25']['usb'][1],
        dfs['LF1']['+25']['usb'][2],

        dfs['LF1']['+25']['usb'][14],   # 2 - H-O
        dfs['LF1']['+25']['usb'][15],
        dfs['LF1']['+25']['usb'][16],
        dfs['LF1']['+25']['usb'][17],

        dfs['LF1']['+25']['usb'][25],   # 3 - P-U
        dfs['LF1']['-60']['usb'][25],
        dfs['LF1']['+85']['usb'][25],

        dfs['LF1']['+25']['usb'][30],   # 4 - V-AA
        dfs['LF1']['-60']['usb'][30],
        dfs['LF1']['+85']['usb'][30],

        dfs['LF1']['+25']['usb'][31],   # 5 - AB-AG
        dfs['LF1']['-60']['usb'][31],
        dfs['LF1']['+85']['usb'][31],

        dfs['LF1']['+25']['usb'][4],   # 11 - AH-AM
        dfs['LF1']['+25']['usb'][5],
        dfs['LF1']['+25']['usb'][6],

        dfs['LF1']['+25']['usb'][7],   # 12 - AN-AU
        dfs['LF1']['+25']['usb'][8],
        dfs['LF1']['+25']['usb'][9],
        dfs['LF1']['+25']['usb'][10],

        dfs['LF1']['+25']['usb'][11],   # 13 - AV-BA
        dfs['LF1']['-60']['usb'][11],
        dfs['LF1']['+85']['usb'][11],

        dfs['LF1']['+25']['usb'][12],   # 14 - BB-BG
        dfs['LF1']['-60']['usb'][12],
        dfs['LF1']['+85']['usb'][12],

        dfs['LF1']['+25']['usb'][13],   # 15 - BH-BM
        dfs['LF1']['-60']['usb'][13],
        dfs['LF1']['+85']['usb'][13],

        dfs['LF1']['+25']['usb'][26],   # 30 - BN-BS
        dfs['LF1']['-60']['usb'][26],
        dfs['LF1']['+85']['usb'][26],

        dfs['LF1']['+25']['usb'][27],   # 31 - BT-BY
        dfs['LF1']['-60']['usb'][27],
        dfs['LF1']['+85']['usb'][27],

        dfs['LF1']['+25']['usb'][28],   # 32 - BZ-CE
        dfs['LF1']['-60']['usb'][28],
        dfs['LF1']['+85']['usb'][28],

        dfs['LF1']['+25']['usb'][29],   # 33 - CF-CK
        dfs['LF1']['-60']['usb'][29],
        dfs['LF1']['+85']['usb'][29],
    ]

    for idx in range(len(filtered)):
        filtered[idx] = filtered[idx].applymap(lambda cell: float(cell.replace(',', '.')))
        filtered[idx]['freq[Hz]'] = filtered[idx]['freq[Hz]'].apply(lambda row: row / 1_000_000_000)
        filtered[idx].columns = ['freq[GHz]', filtered[idx].columns[-1]]

    return pd.concat(objs=filtered, axis=1)


def _build_2_cat_df(dfs):
    # 6-10, 26-29
    filtered = [
        dfs['LF1']['+25']['usb'][32],   # 6 - B-G
        dfs['LF1']['+25']['usb'][33],
        dfs['LF1']['+25']['usb'][34],

        dfs['LF1']['+25']['usb'][35],   # 7 - H-O
        dfs['LF1']['+25']['usb'][36],
        dfs['LF1']['+25']['usb'][37],
        dfs['LF1']['+25']['usb'][38],

        dfs['LF1']['+25']['usb'][39],   # 8 - P-U
        dfs['LF1']['-60']['usb'][39],
        dfs['LF1']['+85']['usb'][39],

        dfs['LF1']['+25']['usb'][40],   # 9 - V-AA
        dfs['LF1']['-60']['usb'][40],
        dfs['LF1']['+85']['usb'][40],

        dfs['LF1']['+25']['usb'][3],   # 10 - AB-AG
        dfs['LF1']['-60']['usb'][3],
        dfs['LF1']['+85']['usb'][3],

        dfs['LF1']['+25']['usb'][18],   # 26 - AH-AO
        dfs['LF1']['+25']['usb'][19],
        dfs['LF1']['+25']['usb'][20],
        dfs['LF1']['+25']['usb'][21],

        dfs['LF1']['+25']['usb'][22],   # 27 - AP-AU
        dfs['LF1']['-60']['usb'][22],
        dfs['LF1']['+85']['usb'][22],

        dfs['LF1']['+25']['usb'][23],   # 28 - AV-BA
        dfs['LF1']['-60']['usb'][23],
        dfs['LF1']['+85']['usb'][23],

        dfs['LF1']['+25']['usb'][24],   # 29 - BB-BG
        dfs['LF1']['-60']['usb'][24],
        dfs['LF1']['+85']['usb'][24],
    ]

    for idx in range(len(filtered)):
        filtered[idx] = filtered[idx].applymap(lambda cell: float(cell.replace(',', '.')))
        filtered[idx]['freq[Hz]'] = filtered[idx]['freq[Hz]'].apply(lambda row: row / 1_000_000_000)
        filtered[idx].columns = ['freq[GHz]', filtered[idx].columns[-1]]

    return pd.concat(objs=filtered, axis=1)


def main(path):
    files = _build_file_list(path)
    _validate_file_list(files)
    dfs = _read_csvs_to_dfs(files)
    # _build_out_dfs(dfs)

    print('extracting 1 band')
    df_for_1_cat = _build_1_cat_df(dfs)
    print('saving df for 1 band')
    df_for_1_cat.to_excel(f'mixer-result-1-5_11-15_30-33-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')

    print('extracting 2 band')
    df_for_2_cat = _build_2_cat_df(dfs)
    print('saving df for 2 band')
    df_for_2_cat.to_excel(f'mixer-result-6-10_26-29-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')


def _add_chart(ws, xs, ys, title, loc):
    chart = LineChart()
    for y in ys:
        chart.add_data(y, titles_from_data=True)
    chart.set_categories(xs)
    chart.title = title
    ws.add_chart(chart, loc)


if __name__ == '__main__':
    # path = sys.argv
    path = os.path.normpath('ref/s2p_process/mixer/measured_data/')
    main(path)
