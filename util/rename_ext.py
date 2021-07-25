import os

# ----- параметры скрипты
source_ext = ['.aaa', '.bbb', '.ccc', '.s2p']
target_ext = '.s1p'
start_dir = 'rename'
# -----


def walk_dirs(source_dir):
    print('processing:', source_dir)

    dirs, files = filter_dir(source_dir)

    print('sub dirs:', dirs)
    print('files:', files)

    rename_files(source_dir, files)

    for d in dirs:
        walk_dirs(f'{source_dir}\\{d}')


def filter_dir(path):
    return \
        [f for f in os.listdir(path) if os.path.isdir(f'{path}\\{f}')], \
        [f for f in os.listdir(path) if os.path.isfile(f'{path}\\{f}') and f[-4:] in source_ext]


def rename_files(dir, files):
    for f in files:
        current_path = f'{dir}\\{f}'
        new_path = current_path[:-4] + target_ext
        os.rename(current_path, new_path)


if __name__ == '__main__':
    walk_dirs(start_dir)
