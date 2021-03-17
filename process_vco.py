import os
import sys


def _filter_s1p(path):
    return sorted([os.path.normpath(f'{path}/{p}') for p in os.listdir(path) if p.endswith('.s1p')])


def main(path):
    for file in _filter_s1p(path):
        print(file)


if __name__ == '__main__':
    # path = sys.argv
    path = os.path.normpath('ref/s2p_process')
    main(path)
