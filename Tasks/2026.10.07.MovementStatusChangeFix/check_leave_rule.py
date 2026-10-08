import sys
import re
import json
import pickle
from datetime import date, timedelta
import check_orders

months = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня', 'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня']
def document_dates(text, pattern):
    dates = []
    for match in re.finditer(pattern + r'[^;\n]*?\bвід\s+(\d{1,2})\.(\d{1,2})\.(\d{4})', text, re.I):
        try:
            dates.append(date(int(match[3]), int(match[2]), int(match[1])))
        except ValueError:
            pass
    return sorted(set(dates))

results = []
with (check_orders.TASK / 'orders_evidence.cache').open('rb') as cached_file:
    saved = pickle.load(cached_file)
if any((check_orders.ROOT / path).stat().st_mtime_ns != stamp for path, stamp in saved['source_mtimes'].items()):
    raise RuntimeError('Исходные приказы изменены: требуется повторное извлечение')
data = check_orders.enhance_context(check_orders.repair_unmatched(saved['data']))
for case in data['cases']:
    p = case['pair']
    if str(p[3]).lower() != 'відпустка для лікування':
        continue
    points = set()
    for m in case['matches']:
        heads = [re.match(r'^(\d+(?:\.\d+)*)\.\s', t) for t in m['context'] + [m['text']]]
        heads = [h[1] for h in heads if h]
        points.add(heads[-1] if heads else 'не установлен')
    if len(points) != 1 or 'не установлен' in points:
        continue
    matches = list({(m['path'], m['paragraph']): m for m in case['matches']}.values())
    text = '\n\n'.join(m['source'] + '\n' + m.get('basis', '') for m in matches)
    discharge = document_dates(text, r'виписк[аи]\s+(?:із|з)\s+медичної\s+карти\s+стаціонарного\s+хворого')
    vlk = document_dates(text, r'(?:довідк[аи]\s+(?:військово[-\s]лікарської\s+комісії|ВЛК))')
    intervals = []
    for m in matches:
        clause = re.search(r'(?:вваж[аи]ти\s+таким|продовжити\s+відпустку)', m['source'], re.I)
        tail = m['source'][clause.start():] if clause else m['source']
        found = re.search(r'з\s+(\d{1,2})\s+(\w+)\s+(?:по\s+\d{1,2}\s+\w+\s+)?(\d{4})\s+року', tail, re.I)
        if found and found[2] in months:
            try:
                intervals.append(date(int(found[3]), months.index(found[2]) + 1, int(found[1])))
            except ValueError:
                pass
    intervals = sorted(set(intervals))
    expected = max(discharge + vlk) + timedelta(days=1) if discharge and vlk else None
    actual = intervals[0] if len(intervals) == 1 else None
    if expected and actual:
        delta = (actual - expected).days
        verdict = 'Соответствует правилу' if delta == 0 else 'Несоответствие правилу'
        reason = f'Выписка: {max(discharge):%d.%m.%Y}; справка ВЛК: {max(vlk):%d.%m.%Y}. Более поздняя дата + 1 день: {expected:%d.%m.%Y}. В приказе: {actual:%d.%m.%Y}. Отклонение: {delta:+d} дн.'
        if delta:
            reason += ' Проверить фактические даты в первичных документах и исправить начало отпуска; конец пересчитать с сохранением назначенной длительности. Сверить Прибуття прежнего события и Вибуття в отпуск.'
    else:
        delta = None
        verdict = 'Требует проверки'
        missing = []
        if not discharge:
            missing.append('дата выписки')
        if not vlk:
            missing.append('дата справки ВЛК')
        if not actual:
            missing.append('однозначная дата начала отпуска')
        reason = 'Не установлены: ' + ', '.join(missing) + '. Проверить оригинал персонального пункта, медицинскую выписку и справку ВЛК; без этих данных полное правило подтвердить нельзя.'
    if len(discharge) > 1 or len(vlk) > 1:
        verdict = 'Требует проверки'
        reason += ' Указаны несколько документов одного вида: проверить, какой относится к этому отпуску; выбор наиболее позднего пока предварительный.'
    results.append({'id': case['id'], 'name': p[0], 'ipn': p[1], 'order': p[4], 'point': '; '.join(points), 'discharge': max(discharge).isoformat() if discharge else None, 'vlk': max(vlk).isoformat() if vlk else None, 'actual': actual.isoformat() if actual else None, 'expected': expected.isoformat() if expected else None, 'delta': delta, 'verdict': verdict, 'reason': reason, 'source': '\n'.join(sorted({m['path'].split('\\')[-1] for m in matches}))})
print(json.dumps(results, ensure_ascii=False))