import sys
import re
import json
import collections
from datetime import date, datetime

import check_orders as source

def detected_status(text):
    t = text.casefold()
    if 'вагітніст' in t and 'полог' in t:
        return "Відпустка у зв'язку з вагітністю та пологами"
    if re.search(r'відпуст\w*[^.;]{0,55}сімейн', t):
        return 'Відпустка за сімейними обставинами'
    if re.search(r'відпуст\w*[^.;]{0,55}(?:для лікування|лікув)', t):
        return 'Відпустка для лікування'
    if 'щорічн' in t:
        return 'Щорічна відпустка'
    if 'самовільн' in t:
        return 'Самовільне залишення частини'
    if 'амбулаторн' in t or 'медичному огляді' in t or 'військово-лікарської комісії' in t:
        return 'Амбулаторне ВЛК'
    if 'лікуван' in t or re.search(r'лікувальн|лікарн|госпітал|медичн|інститут|диспансер|клінічн', t):
        return 'Стаціонарне лікування'
    if 'відряджен' in t or re.search(r'^\d+(?:\.\d+)*\.\s*з військової частини', t):
        return 'Відрядження'
    return ''

def start_date(text, order_date, fallback=True):
    text = source.normalized(text)
    # Даты документов и продовольственного обеспечения не заменяют фактическое событие.
    text = re.split(r'(?:Зарахувати|Зняти|Видати|Підстава:)|(?:на продовольч[ео] забезпечення)', text, maxsplit=1, flags=re.I)[0]
    interval = re.search(r'з\s+(\d{1,2})\s+по\s+\d{1,2}\s+([а-яіїєґ]+)(?:\s+(20\d{2}))?', text, re.I)
    if interval and interval.group(2).lower() in source.MONTHS:
        return date(int(interval.group(3) or order_date.year),source.MONTHS[interval.group(2).lower()],int(interval.group(1))), 'Дата начала периода'
    matches = list(re.finditer(r'з\s+(\d{1,2})\s+([а-яіїєґ]+)(?:\s+(20\d{2}))?',text,re.I))
    matches = [m for m in matches if m.group(2).lower() in source.MONTHS]
    if matches:
        m = matches[0]
        return date(int(m.group(3) or order_date.year),source.MONTHS[m.group(2).lower()],int(m.group(1))), 'Прямая дата события'
    if fallback:
        return order_date, 'Дата приказа: отдельная дата события не указана'
    return None, ''

def evidence_record(match, status, direction, event_date, date_rule, kind):
    return {'status':status,'direction':direction,'date':event_date.isoformat() if event_date else '', 'rule':date_rule,'kind':kind,'match':match}

def events_from_match(match):
    if not match['header_verified'] or match['identity'].startswith('Возможное'):
        return []
    order_date = date.fromisoformat(match['order_date'])
    text = match['text']
    headings = match['context']
    context = ' '.join(h for h in headings if h != text)
    own = re.sub(r'^\d+(?:\.\d+)*\.\s*','',text)
    body = source.normalized(' '.join([text,*match['continuation']]))
    records = []
    awol=re.search(r'(\d{1,2})\s+([а-яіїєґ]+)\s+(20\d{2})\s+року\s+самовільно залишив',own,re.I)
    if awol:
        d=date(int(awol.group(3)),source.MONTHS[awol.group(2).lower()],int(awol.group(1)))
        records.append(evidence_record(match,'Самовільне залишення частини','Выбытие',d,'Прямая дата самовольного оставления','Отдельное событие'))
        if 'не повернувся' in own.casefold() and detected_status(own[awol.end():])=='Стаціонарне лікування':
            records.append(evidence_record(match,'Стаціонарне лікування','Прибытие',d,'Учётное закрытие лечения датой начала СЗЧ; физическое возвращение не заявлено','Учётная смена статуса'))
        return records
    if re.search(r'прибув до військової частини',own,re.I) and re.search(r'Поновити (?:всі|на всіх) вид',body,re.I):
        d,rule=start_date(body,order_date)
        records.append(evidence_record(match,'Самовільне залишення частини','Прибытие',d,'Возвращение в другую часть и восстановление обеспечения','Отдельное событие'))
        return records
    closing_command=re.search(r'припинити[^.;]*?відряджен[^.;]*',body,re.I)
    if closing_command:
        close_date,close_rule=start_date(closing_command.group(),order_date,fallback=False)
        if close_date:
            records.append(evidence_record(match,'Відрядження','Прибытие',close_date,close_rule,'Отдельная дата закрытия'))
    # Завершение лечения и отпуск могут иметь две прямо названные даты в одном пункте.
    completion=re.search(r'(?:\d{1,2}\s+[а-яіїєґ]+\s+20\d{2}\s+року\s+)?(?:завершив (?:лікування|перебування)|був виписаний)',own,re.I)
    if completion:
        absolute=re.search(r'(\d{1,2})\s+([а-яіїєґ]+)\s+(20\d{2})',completion.group(),re.I)
        numeric=re.search(r'був виписаний\s+(\d{1,2})\.(\d{1,2})\.(20\d{2})',own,re.I)
        close_date=date(int(absolute.group(3)),source.MONTHS[absolute.group(2).lower()],int(absolute.group(1))) if absolute else (date(int(numeric.group(3)),int(numeric.group(2)),int(numeric.group(1))) if numeric else order_date)
        records.append(evidence_record(match,'Стаціонарне лікування','Прибытие',close_date,'Прямо указанное завершение лечения' if absolute or numeric else 'Завершение лечения в дату приказа','Отдельная дата закрытия'))
        new_action=re.search(r'(?:вибув|подовжити|продовжити)[^;]*',own[completion.end():],re.I)
        if new_action:
            new_status=detected_status(new_action.group())
            if new_status:
                d,rule=start_date(new_action.group(),order_date)
                records.append(evidence_record(match,new_status,'Выбытие',d,rule,'Отдельное событие'))
        return records
    # Переход прямо в пункте: прежний статус, затем новый статус с собственной датой.
    transition = re.search(r'(?:вваж(?:ати|ити)\s+так(?:им|ою|ими)(?:\s*,?\s*що)?|продовжити|подовжити|надати)',own,re.I)
    if transition:
        old_part = own[:transition.start()]
        new_part = own[transition.end():]
        history=list(re.finditer(r'перебував|перебувала|проходив|проходила|був госпіталізований|була госпіталізована',old_part,re.I))
        previous_status=detected_status(old_part[history[-1].start():]) if history else ''
        new_status = detected_status(new_part)
        direct_change = bool(previous_status and new_status and previous_status != new_status)
        if previous_status and new_status and previous_status==new_status:
            d,rule=start_date(new_part,order_date)
            records.append(evidence_record(match,new_status,'Выбытие',d,rule,'Отдельное событие'))
            return records
        if direct_change:
            d,rule = start_date(new_part,order_date)
            records.append(evidence_record(match,new_status,'Выбытие',d,rule,'Прямая смена статуса'))
            closing_text = re.search(r'(?:Припинити|Закінчити|Припинено)[^.;]*?(?:відряджен|відпуст)[^.;]*',body,re.I)
            explicit_close = None
            if closing_text:
                explicit_close,close_rule = start_date(closing_text.group(),order_date,fallback=False)
            if explicit_close:
                closing_status = detected_status(closing_text.group()) or previous_status
                records.append(evidence_record(match,closing_status,'Прибытие',explicit_close,close_rule,'Отдельная дата закрытия'))
            else:
                records.append(evidence_record(match,previous_status,'Прибытие',d,'Учётное закрытие прежнего статуса датой начала нового; вывод из прямого перехода','Учётная смена статуса'))
            return records
    # Возвращение из СЗЧ с отдельным направлением в новую командировку.
    if 'самовільно залишив' in own.casefold() and 'повернувся' in own.casefold():
        returned = re.split(r'вваж(?:ати|ити)\s+так(?:им|ою)[^.;]{0,25}що',own,maxsplit=1,flags=re.I)[-1]
        d,rule=start_date(returned,order_date)
        records.append(evidence_record(match,'Самовільне залишення частини','Прибытие',d,rule,'Отдельное событие'))
        command=re.search(r'Направити у відрядження[^.]*?(?:з\s+\d{1,2}\s+[а-яіїєґ]+\s+20\d{2})',body,re.I)
        if command:
            departure,departure_rule=start_date(command.group(),order_date)
            records.append(evidence_record(match,'Відрядження','Выбытие',departure,departure_rule,'Отдельное событие'))
        return records
    # Отдельные пункты о возвращении и новом выбытии сохраняют собственные даты.
    returning = bool(re.search(r'прибул[иоа]|повернул|повернув',context,re.I))
    departing = bool(re.search(r'вибул[иоа]|направити',context,re.I))
    if not departing and any(re.match(r'^\d+(?:\.\d+)*\.\s*(?:З|Із)\s+(?!\d)',h,re.I) for h in headings if h!=text):
        returning=True
    subheading = headings[-1] if headings and headings[-1] != text else context
    status = detected_status(subheading)
    if not status:
        for h in reversed(headings):
            if h==text:
                continue
            status=detected_status(h)
            if status:
                subheading=h
                break
    if not status and returning and re.search(r'стаціонарн|виписка.*медичн',match['basis'],re.I):
        status='Стаціонарне лікування'
    # Отпуск в собственном тексте после возвращения — самостоятельное открытие.
    new_action=re.search(r'(?:(?:вибув|направити)\s+(?:у|в|на)\s+|продовжити|подовжити)[^;]*',own,re.I)
    if new_action:
        new_status=detected_status(new_action.group())
        if new_status:
            d,rule=start_date(new_action.group(),order_date)
            records.append(evidence_record(match,new_status,'Выбытие',d,rule,'Отдельное событие'))
    # Возвращение из военной части само по себе не доказывает командировку:
    # лечебные учреждения также имеют номер военной части.
    if status=='Відрядження' and 'відряджен' not in subheading.casefold():
        status = 'Стаціонарне лікування' if re.search(r'стаціонарн|виписка.*медичн',match['basis'],re.I) else ('Відрядження' if 'посвідчення про відрядження' in match['basis'].casefold() else '')
    if returning or departing:
        if not status:
            return records
        factual_own=own[:new_action.start()] if returning and new_action else own
        explicit,rule = start_date(factual_own,order_date,fallback=False)
        if explicit is None:
            for h in reversed(headings):
                if h == text:
                    continue
                explicit,rule = start_date(h,order_date,fallback=False)
                if explicit is not None:
                    break
        if explicit is None:
            explicit,rule = order_date,'Дата приказа: отдельная дата события не указана'
        records.append(evidence_record(match,status,'Прибытие' if returning else 'Выбытие',explicit,rule,'Отдельное событие'))
    return records

def main():
    data = source.main()
    decisions = []
    for case in data['cases']:
        p = case['pair']
        records = [r for m in case['matches'] for r in events_from_match(m)]
        sides = []
        for status,direction,original_date,rownum in [(p[2],'Прибытие',p[5][:10],p[8]),(p[3],'Выбытие',p[6][:10],p[9])]:
            candidates = [r for r in records if r['direction']==direction and r['status'].casefold()==status.casefold()]
            explicit = [r for r in candidates if r['kind']!='Учётная смена статуса']
            if explicit:
                candidates = explicit
            signatures = {(r['status'].casefold(),r['date']) for r in candidates}
            if len(signatures)==1:
                chosen = candidates[0]
                corrected_date = chosen['date']
                corrected_status = status
                changed = corrected_date != original_date
                verdict = ('Предлагается учётная дата' if chosen['kind']=='Учётная смена статуса' else 'Исправить дату') if changed else 'Без изменений'
                reason = chosen['rule']
                if changed and chosen['rule'].startswith('Дата приказа:'):
                    verdict='Предлагается дата приказа'
                    reason='Отдельная фактическая дата отсутствует. По общему правилу используется дата приказа; продовольственная дата не заменяет фактическую'
                if abs((date.fromisoformat(corrected_date)-date.fromisoformat(original_date)).days)>3 and case['id']!=319:
                    verdict='Проверить дату в приказе'
                    reason='Дата в пункте приказа отличается от исходной более чем на 3 дня: '+corrected_date+'. Возможна ошибка в приказе или неверная ссылка на приказ'
                    corrected_date=corrected_status=''
            elif len(signatures)>1:
                chosen = candidates[0]
                corrected_date = corrected_status = ''
                verdict = 'Проверить противоречие'
                reason = 'В источниках найдено несколько разных дат: '+', '.join(sorted({r['date'] for r in candidates}))
            else:
                alternative = [r for r in records if r['direction']==direction and r['kind'] in ('Отдельное событие','Учётная смена статуса')]
                signatures = {(r['status'],r['date']) for r in alternative}
                if len(signatures)==1:
                    chosen = alternative[0]
                    corrected_status,corrected_date = chosen['status'],chosen['date']
                    verdict = 'Уточнить тип медицинского события' if {status,corrected_status}=={'Амбулаторне ВЛК','Стаціонарне лікування'} else ('Исправить статус и дату' if corrected_date!=original_date else 'Исправить статус')
                    reason = 'В пункте приказа указан другой статус. '+chosen['rule']
                    if abs((date.fromisoformat(corrected_date)-date.fromisoformat(original_date)).days)>3:
                        verdict='Проверить дату в приказе'
                        reason='Найденный пункт относится к другой дате: '+corrected_date+'. Нужна проверка номера/года приказа и события'
                        corrected_date=corrected_status=''
                    if verdict=='Уточнить тип медицинского события':
                        reason='В приказе указан медосмотр/лечебное учреждение; амбулаторный или стационарный характер нужно подтвердить медицинским документом. '+chosen['rule']
                        corrected_date=corrected_status=''
                else:
                    chosen = None
                    corrected_date = corrected_status = ''
                    verdict = 'Требуется проверка'
                    reason = 'Соответствующее событие не найдено однозначно в тексте приказа'
            sides.append({'original_status':status,'original_date':original_date,'direction':direction,'source_row':rownum,'corrected_status':corrected_status,'corrected_date':corrected_date,'verdict':verdict,'reason':reason,'chosen':chosen,'candidates':candidates})
        overrides={
            27:'В таблице Євген Євгенович, в приказе Євген Євгенійович. Без ІПН в пункте совпадение человека не подтверждено',
            67:"В таблице ВОРОБЕЙ, в приказе ВОРОБ'Я. Нужно подтвердить вариант фамилии по ІПН или первичному документу",
            89:'Пункт 70 содержит срок и место, но действие и прежний статус пропущены. Подтверждать смену статуса по этому тексту нельзя',
            191:'№369 от 19.12.2025 описывает отпуск с 20.12.2025; исходная пара относится к 14–15.09.2025. Проверить ссылку на приказ',
            210:'В таблице ЛОМАКО, в приказе ЛОМАЦІ. Нужно подтвердить форму фамилии и человека по ІПН или первичному документу',
            272:'В имени файла №162 от 03.06.2026, в заголовке документа №161 от 02.06.2026. Номер и дата источника противоречат ссылке',
            334:'В пункте дата выписки 22.08.2035, в основании 22.08.2025; заголовок «отпуск для лечения», персональная формулировка «щорічна відпустка». Источник противоречив',
            440:'В приказе прежний статус назван отпуском для лечения, в таблице — стационарным лечением, а основание содержит выписку из стационара. Уточнить прежний статус',
        }
        if case['id'] in overrides:
            affected=range(2) if case['id'] not in (440,) else [0]
            for side_index in affected:
                sides[side_index].update({'verdict':'Требуется проверка','corrected_status':'','corrected_date':'','reason':overrides[case['id']]})
        if case['id']==265:
            sides[1].update({'verdict':'Уточнить статус по приказу','corrected_status':'','corrected_date':'','reason':'Пункт направляет в медицинскую роту, но не называет стационарное лечение или отпуск; основание содержит отпускной билет. Тип отсутствия неоднозначен'})
        additions=[]
        if p[2]=='Самовільне залишення частини' and p[3]=='Відрядження' and all(s['chosen'] and s['verdict']=='Без изменений' for s in sides):
            arrival=date.fromisoformat(sides[0]['corrected_date'])
            departure=date.fromisoformat(sides[1]['corrected_date'])
            if arrival<departure:
                additions.append({'status':'Повернення до іншої частини','from':arrival.isoformat(),'to':departure.isoformat(),'reason':'Возвращение после СЗЧ в другую часть раньше начала командировки. По схеме обработки РУХ требуется промежуточное событие','match':sides[0]['chosen']['match']})
        decision = {'id':case['id'],'pair':p,'sides':sides,'case':case,'additions':additions}
        decisions.append(decision)
    return {'decisions':decisions,'orders':data['order_keys'],'files':data['files']}

if __name__=='__main__':
    data=main()
    if '--summary' in sys.argv:
        counts=collections.Counter(s['verdict'] for d in data['decisions'] for s in d['sides'])
        print(json.dumps({'counts':counts,'changed':[(d['id'],d['pair'][0],[(s['direction'],s['original_date'],s['corrected_date'],s['verdict']) for s in d['sides'] if s['verdict'].startswith('Исправить')]) for d in data['decisions'] if any(s['verdict'].startswith('Исправить') for s in d['sides'])], 'pending':[(d['id'],d['pair'][0],[(s['direction'],s['original_status'],s['original_date'],s['reason']) for s in d['sides'] if s['verdict'] not in ('Без изменений','Исправить дату','Исправить статус','Исправить статус и дату')]) for d in data['decisions'] if any(s['verdict'] not in ('Без изменений','Исправить дату','Исправить статус','Исправить статус и дату') for s in d['sides'])]},ensure_ascii=False))
    elif '--details' in sys.argv:
        ids={int(x) for x in sys.argv[sys.argv.index('--details')+1].split(',')}
        print(json.dumps([d for d in data['decisions'] if d['id'] in ids],ensure_ascii=False))
    else:
        print(json.dumps(data,ensure_ascii=False))