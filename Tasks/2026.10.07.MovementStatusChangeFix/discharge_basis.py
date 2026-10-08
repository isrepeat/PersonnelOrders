import sys
import re
import json
import warnings
from datetime import date, timedelta
import openpyxl

warnings.filterwarnings('ignore')
sys.stdout.reconfigure(encoding='utf-8')
sheet = openpyxl.load_workbook(sys.argv[1], data_only=True).active
months = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня', 'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня']
updates = []
for row in range(11, sheet.max_row + 1, 4):
    if sheet.cell(row, 4).value != 'Стаціонарне лікування' or str(sheet.cell(row + 1, 4).value).lower() != 'відпустка для лікування':
        continue
    text = sheet.cell(row, 13).value
    clause = re.search(r'вваж[аи]ти таким', text, re.I)
    if not clause:
        continue
    interval = re.search(r'з\s+(\d{1,2})\s+(\w+)\s+по\s+(\d{1,2})\s+(\w+)\s+(\d{4})\s+року', text[clause.start():], re.I)
    discharge = re.search(r'виписк[аи]\s+(?:із|з)\s+медичної\s+карти\s+стаціонарного\s+хворого[^\n;]*?\bвід\s+(\d{2})\.(\d{2})\.(\d{4})', text, re.I)
    if not interval or not discharge or interval[2] not in months or interval[4] not in months:
        continue
    start = date(int(interval[5]), months.index(interval[2]) + 1, int(interval[1]))
    released = date(int(discharge[3]), int(discharge[2]), int(discharge[1]))
    if start != released:
        continue
    end = date(int(interval[5]), months.index(interval[4]) + 1, int(interval[3]))
    if end < start:
        end = end.replace(year=end.year + 1)
    start += timedelta(days=1)
    end += timedelta(days=1)
    replacement = f'з {start.day:02d} {months[start.month-1]} по {end.day:02d} {months[end.month-1]} {end.year} року'
    corrected = text.replace(interval[0], replacement)
    spans = []
    for found in re.finditer(re.escape(replacement), corrected):
        spans.extend([[found.start() + 3, 2], [found.start() + replacement.index('по ') + 4, 2]])
    reason = f'Ошибка в приказе: начало отпуска совпадает с датой выписки {released:%d.%m.%Y}, указанной в основании. По правилу выписки переход должен быть {start:%d.%m.%Y}; период отпуска — {start:%d.%m.%Y}–{end:%d.%m.%Y}. Прибуття после лечения и Вибуття в отпуск исправлены; длительность отпуска сохранена. Дата документа использована как дата выписки по принятому правилу; при расхождении сверить фактическую дату в самой выписке. Остальные даты сохранены.'
    updates.append({'row': row, 'name': sheet.cell(row, 2).value, 'start': start.isoformat(), 'text': corrected, 'spans': spans, 'reason': reason})
print(json.dumps(updates, ensure_ascii=False))