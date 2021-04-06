import datetime
import io
import os

import pandas as pd

from openpyxl.chart import LineChart

# для категории 1, измерения 1-5, 11-15, 30-33
# 1 - B-G
# 2 - H-O
# 3 - P-U
# 4 - V-AA
# 5 - AB-AG
# 11 - AH-AM
# 12 - AN-AU
# 13 - AV-BA
# 14 - BB-BG
# 15 - BH-BM
# 30 - BN-BS
# 31 - BT-BY
# 32 - BZ-CE
# 33 - CF-CK

# для категории 2, измерения 6-10, 26-29
# 6 - B-G
# 7 - H-O
# 8 - P-U
# 9 - V-AA
# 10 - AB-AG
# 26 - AH-AO
# 27 - AP-AU
# 28 - AV-BA
# 29 - BB-BG


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


def _build_1_33_measure_1_cat_df(dfs):
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


def _build_1_33_measure_2_cat_df(dfs):
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


def _build_34_66_measure_1_cat_df(dfs):
    # 1 cat - 34-38, 44-48, 63-66
    filtered = [
        dfs['LF1']['+25']['lsb'][0],   # 34 - B-G
        dfs['LF1']['+25']['lsb'][1],
        dfs['LF1']['+25']['lsb'][2],

        dfs['LF1']['+25']['lsb'][14],   # 35 - H-O
        dfs['LF1']['+25']['lsb'][15],
        dfs['LF1']['+25']['lsb'][16],
        dfs['LF1']['+25']['lsb'][17],

        dfs['LF1']['+25']['lsb'][25],   # 36 - P-U
        dfs['LF1']['-60']['lsb'][25],
        dfs['LF1']['+85']['lsb'][25],

        dfs['LF1']['+25']['lsb'][30],   # 37 - V-AA
        dfs['LF1']['-60']['lsb'][30],
        dfs['LF1']['+85']['lsb'][30],

        dfs['LF1']['+25']['lsb'][31],   # 38 - AB-AG
        dfs['LF1']['-60']['lsb'][31],
        dfs['LF1']['+85']['lsb'][31],

        dfs['LF1']['+25']['lsb'][4],   # 44 - AH-AM
        dfs['LF1']['+25']['lsb'][5],
        dfs['LF1']['+25']['lsb'][6],

        dfs['LF1']['+25']['lsb'][7],   # 45 - AN-AU
        dfs['LF1']['+25']['lsb'][8],
        dfs['LF1']['+25']['lsb'][9],
        dfs['LF1']['+25']['lsb'][10],

        dfs['LF1']['+25']['lsb'][11],   # 46 - AV-BA
        dfs['LF1']['-60']['lsb'][11],
        dfs['LF1']['+85']['lsb'][11],

        dfs['LF1']['+25']['lsb'][12],   # 47 - BB-BG
        dfs['LF1']['-60']['lsb'][12],
        dfs['LF1']['+85']['lsb'][12],

        dfs['LF1']['+25']['lsb'][13],   # 48 - BH-BM
        dfs['LF1']['-60']['lsb'][13],
        dfs['LF1']['+85']['lsb'][13],

        dfs['LF1']['+25']['lsb'][26],   # 63 - BN-BS
        dfs['LF1']['-60']['lsb'][26],
        dfs['LF1']['+85']['lsb'][26],

        dfs['LF1']['+25']['lsb'][27],   # 64 - BT-BY
        dfs['LF1']['-60']['lsb'][27],
        dfs['LF1']['+85']['lsb'][27],

        dfs['LF1']['+25']['lsb'][28],   # 65 - BZ-CE
        dfs['LF1']['-60']['lsb'][28],
        dfs['LF1']['+85']['lsb'][28],

        dfs['LF1']['+25']['lsb'][29],   # 66 - CF-CK
        dfs['LF1']['-60']['lsb'][29],
        dfs['LF1']['+85']['lsb'][29],
    ]

    for idx in range(len(filtered)):
        filtered[idx] = filtered[idx].applymap(lambda cell: float(cell.replace(',', '.')))
        filtered[idx]['freq[Hz]'] = filtered[idx]['freq[Hz]'].apply(lambda row: row / 1_000_000_000)
        filtered[idx].columns = ['freq[GHz]', filtered[idx].columns[-1]]

    return pd.concat(objs=filtered, axis=1)


def _build_34_66_measure_2_cat_df(dfs):
    # 39-43, 59-62
    filtered = [
        dfs['LF1']['+25']['lsb'][32],   # 39 - B-G
        dfs['LF1']['+25']['lsb'][33],
        dfs['LF1']['+25']['lsb'][34],

        dfs['LF1']['+25']['lsb'][35],   # 41 - H-O
        dfs['LF1']['+25']['lsb'][36],
        dfs['LF1']['+25']['lsb'][37],
        dfs['LF1']['+25']['lsb'][38],

        dfs['LF1']['+25']['lsb'][39],   # 42 - P-U
        dfs['LF1']['-60']['lsb'][39],
        dfs['LF1']['+85']['lsb'][39],

        dfs['LF1']['+25']['lsb'][40],   # 43 - V-AA
        dfs['LF1']['-60']['lsb'][40],
        dfs['LF1']['+85']['lsb'][40],

        dfs['LF1']['+25']['lsb'][3],   # 44 - AB-AG
        dfs['LF1']['-60']['lsb'][3],
        dfs['LF1']['+85']['lsb'][3],

        dfs['LF1']['+25']['lsb'][18],   # 59 - AH-AO
        dfs['LF1']['+25']['lsb'][19],
        dfs['LF1']['+25']['lsb'][20],
        dfs['LF1']['+25']['lsb'][21],

        dfs['LF1']['+25']['lsb'][22],   # 60 - AP-AU
        dfs['LF1']['-60']['lsb'][22],
        dfs['LF1']['+85']['lsb'][22],

        dfs['LF1']['+25']['lsb'][23],   # 61 - AV-BA
        dfs['LF1']['-60']['lsb'][23],
        dfs['LF1']['+85']['lsb'][23],

        dfs['LF1']['+25']['lsb'][24],   # 62 - BB-BG
        dfs['LF1']['-60']['lsb'][24],
        dfs['LF1']['+85']['lsb'][24],
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

    print('extracting measure 1-33, 1 band')
    df_for_1_cat = _build_1_33_measure_1_cat_df(dfs)
    print('saving measure 1-33, 1 band')
    df_for_1_cat.to_excel(f'mixer-result-1-5_11-15_30-33-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')

    print('extracting measure 1-33, 2 band')
    df_for_2_cat = _build_1_33_measure_2_cat_df(dfs)
    print('saving measure 1-33, 2 band')
    df_for_2_cat.to_excel(f'mixer-result-6-10_26-29-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')

    print('extracting measure 34-66, 1 band')
    df_for_1_cat = _build_1_33_measure_1_cat_df(dfs)
    print('saving measure 34-66, 1 band')
    df_for_1_cat.to_excel(f'mixer-result-34-38_44-48_63-66-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')

    print('extracting measure 34-66, 2 band')
    df_for_2_cat = _build_1_33_measure_2_cat_df(dfs)
    print('saving measure 34-66, 2 band')
    df_for_2_cat.to_excel(f'mixer-result-39-43_59-62-{datetime.datetime.now().isoformat().replace(":", ".")}.xlsx')


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
