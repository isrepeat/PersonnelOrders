"""Читает проверенную расшифровку фото выбранной папки Sources и сверяет ШПО."""

import argparse
import json
import sys

from food_run import load_run


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--date', required=True, help='Имя созданной пользователем папки Sources/YYYY.MM.DD')
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    print(json.dumps(load_run(args.date), ensure_ascii=False))