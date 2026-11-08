"""Читает проверенную расшифровку фото выбранной папки Sources и сверяет ШПО."""

import argparse
import json
import sys

from food_run import load_run


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--date', required=True, help='Имя созданной пользователем папки Sources/YYYY.MM.DD')
    parser.add_argument('--output', help='Путь новой версии таблицы для определения первой версии записи')
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8')
    run = load_run(args.date)
    if args.output:
        from food_versions import apply_versions
        apply_versions(run, args.output)
    print(json.dumps(run, ensure_ascii=False))