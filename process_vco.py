import os
import pandas as pd


def _fileter_xlsx(path):
    return sorted([os.path.normpath(f'{path}/{p}') for p in os.listdir(path) if p.endswith('.xlsx')])


def _is_file_list_valid(files):
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

    result.to_excel('out.xlsx', engine='openpyxl')


if __name__ == '__main__':
    # path = sys.argv
    path = os.path.normpath('ref/s2p_process/vco/xlsx')
    main(path)
