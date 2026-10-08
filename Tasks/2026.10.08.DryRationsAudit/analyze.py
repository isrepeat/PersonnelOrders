import re
import sys
import json
import gzip
import datetime as dt
import collections
from pathlib import Path

import openpyxl
from docx import Document

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'outputs' / 'dry-rations-audit-20261008'
OUT.mkdir(parents=True, exist_ok=True)
MONTHS = {'серпня': 8, 'вересня': 9, 'жовтня': 10, 'липня': 7, 'листопада': 11}
MONTH = '|'.join(MONTHS)
RANKS = {'молодшого сержанта':'молодший сержант','старшого сержанта':'старший сержант','головного сержанта':'головний сержант','штаб-сержанта':'штаб-сержант','майстер-сержанта':'майстер-сержант','старшого солдата':'старший солдат','старший солдата':'старший солдат','головний сержанта':'головний сержант','солдата':'солдат','солдат':'солдат','капітана':'капітан','майора':'майор','сержанта':'сержант','лейтенанта':'лейтенант','лейтенант':'лейтенант','старшого лейтенанта':'старший лейтенант','матроса':'матрос','рядового':'рядовий'}
PERSON = re.compile(r'^(?:\d+(?:\.\d+)*\.\s*)?('+'|'.join(sorted(map(re.escape,RANKS),key=len,reverse=True))+r')\s+([А-ЯІЇЄҐA-Z][А-ЯІЇЄҐA-Z’\x27\-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ’\x27\-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ’\x27\-]+)',re.I)

RANKS.update({'молодшого лейтенанта':'молодший лейтенант','старшого лейтенант':'старший лейтенант','старшого сержант':'старший сержант','молодшого сержант':'молодший сержант','рекрута':'рекрут'})
PERSON = re.compile(r'^(?:\d+(?:\.\d+)*\.\s*)?('+'|'.join(sorted(map(re.escape,RANKS),key=len,reverse=True))+r')\s+([А-ЯІЇЄҐA-Z][А-ЯІЇЄҐA-Z’\x27\-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ’\x27\-]+)\s+([А-ЯІЇЄҐ][а-яіїєґ’\x27\-]+)',re.I)

def clean(text):
    text=re.sub(r'солдат\s+а\b','солдата',text)
    return re.sub(r'\s+',' ',text.replace('\u00a0',' ')).strip()

def norm(text):
    return ''.join(c for c in str(text).casefold().replace('a','а') if c.isalpha())

def period(text):
    match=re.search(r'з\s*(\d{1,2})\s*(?:('+MONTH+r'))?\s*по\s*(\d{1,2})\s*('+MONTH+r')\s*(2026)',text,re.I)
    if match:
        day, month1, end, month2, year=match.groups()
        start=dt.date(int(year),MONTHS[month1 or month2],int(day))
        stop=dt.date(int(year),MONTHS[month2],int(end))
        return start, (stop-start).days+1
    match=re.search(r'(?:на|з)\s+(\d{1,2})\s+('+MONTH+r')\s+(2026)',text,re.I)
    if match:
        return dt.date(int(match[3]),MONTHS[match[2]],int(match[1])),None
    match=re.search(r'(\d{1,2})\s+('+MONTH+r')\s+(2026)',text,re.I)
    if match: return dt.date(int(match[3]),MONTHS[match[2]],int(match[1])),None
    return None,None

book_path=Path('C:/Users/isrepeat/OneDrive/Рабочий стол/РУХ_last.xlsx')
wb=openpyxl.load_workbook(Path(sys.argv[1]) if len(sys.argv)>1 else book_path,read_only=True,data_only=True)
existing=[]
canonical=set()
for rn,row in enumerate(wb['Сухпрод'].iter_rows(max_col=10,values_only=True),1):
    if rn>2 and row[2]:
        vals=list(row[1:10])
        vals=[v.date().isoformat() if isinstance(v,dt.datetime) else v for v in vals]
        existing.append({'row':rn,'values':vals})
        canonical.add(str(row[2]).strip())
for row in wb['Список'].iter_rows(min_row=6,max_col=3,values_only=True):
    if row[1]: canonical.add(str(row[1]).strip())
wb.close()
shpo=openpyxl.load_workbook(ROOT/'Tools'/'ШПО.xlsx',read_only=True,data_only=True)
authoritative={}
for row in shpo['АЛФ'].iter_rows(min_row=2,max_col=2,values_only=True):
    if row[1]:
        canonical.add(str(row[1]).strip())
        authoritative[norm(row[1])]=str(row[1]).strip()
shpo.close()

def forms(word,index):
    word=word.casefold().replace('’',"'")
    result={norm(word),norm(word+'а'),norm(word+'у'),norm(word+'я'),norm(word+'ю')}
    if word.endswith('й'): result.update(map(norm,[word[:-1]+'я',word[:-1]+'ю']))
    if index==0 and word.endswith(('ий','ій')): result.update(map(norm,[word[:-2]+'ого',word[:-2]+'ього']))
    if word.endswith('о'): result.update(map(norm,[word[:-1]+'а',word[:-1]+'у']))
    if word.endswith('а'): result.update(map(norm,[word[:-1]+'и',word[:-1]+'і',word[:-1]+'у']))
    if word.endswith('я'): result.update(map(norm,[word[:-1]+'і',word[:-1]+'ю']))
    if word.endswith('я'): result.add(norm(word[:-1]+'ї'))
    if word.endswith('ь'): result.update(map(norm,[word[:-1]+'я',word[:-1]+'ю']))
    if index==0 and word.endswith('ець'): result.update(map(norm,[word[:-3]+'ця',word[:-3]+'цю',word[:-3]+'ьця']))
    if index==0 and word.endswith('єць'): result.update(map(norm,[word[:-3]+'йця',word[:-3]+'йцю']))
    if index==0 and word.endswith('ок'): result.update(map(norm,[word[:-2]+'ка',word[:-2]+'ку']))
    if index==0 and word.endswith('ів'): result.update(map(norm,[word[:-2]+'ова',word[:-2]+'ову']))
    if index==0 and word.endswith('ой'): result.add(norm(word[:-2]+'ого'))
    if index==0 and word=='соловей': result.update(map(norm,["солов'я",'соловья']))
    if index==0 and word=='моругий': result.add(norm('моругу'))
    if index==1:
        irregular={'павло':'павла','лев':'лева','ігор':'ігоря','олесь':'олеся','любов':'любові'}
        if word in irregular: result.add(norm(irregular[word]))
    return result

indexes=[collections.defaultdict(set) for _ in range(3)]
by_norm={}
for name in sorted(canonical):
    parts=name.split()
    if len(parts)!=3: continue
    key=norm(name)
    by_norm.setdefault(key,name)
    for i,word in enumerate(parts):
        for form in forms(word,i): indexes[i][form].add(key)

def resolve(parts):
    manual={'КЕРНОСА Олександра Івановича':'КІРНОС Олександр Іванович','ШЕПЕТОВСЬКОГО Тимура Євгенійовича':'ШЕПЕТОВСЬКИЙ Тимур Євгенович','КОСТИРКА Владислава Дмитровича':'КОСТИРКО Владислав Дмитрович'}
    original=' '.join(parts)
    if original in manual: return manual[original],'Вариант написания в приказе: '+original
    keys=set.intersection(*(indexes[i].get(norm(word),set()) for i,word in enumerate(parts)))
    if len(keys)==1: return by_norm[next(iter(keys))],''
    preferred=keys & set(authoritative)
    if len(preferred)==1: return authoritative[next(iter(preferred))],''
    return ' '.join(parts), 'Не найдено однозначное соответствие ФИО: '+' / '.join(by_norm[k] for k in sorted(keys))

groups=[]
warnings=[]
files=sorted(p for p in (ROOT/'Documents/Orders/А7383/2026').glob('*.docx') if re.match(r'2026-(08|09|10)-',p.name))
for path in files:
    paragraphs=[clean(p.text) for p in Document(path).paragraphs]
    order=int(re.search(r'№\s*(\d+)',path.name)[1])
    for hit,text in enumerate(paragraphs):
        if not ('сух' in text.lower() and 'пай' in text.lower()): continue
        if 'нижчепойменованих' not in text.lower():
            # Для индивидуальной выдачи дата берётся из ближайшего абзаца о выбытии.
            lower=next((j for j in range(hit-1,-1,-1) if paragraphs[j].startswith('Підстава:')),0)
            date_text=next((paragraphs[j] for j in range(lower+1,hit) if period(paragraphs[j])[0]),'')
            start,_=period(date_text)
            duration=int(re.search(r'на\s+(\d+)',text)[1])
            basis=next((paragraphs[j] for j in range(hit+1,min(len(paragraphs),hit+6)) if paragraphs[j].startswith('Підстава:')),'')
            persons=[]
            for j in range(lower+1,hit):
                m=PERSON.match(paragraphs[j])
                if m:
                    name,note=resolve(m.groups()[1:])
                    persons.append({'rank':RANKS[m[1].lower()],'name':name,'original':' '.join(m.groups()[1:]),'paragraph':j+1,'note':note})
            if not persons or not start or not basis: warnings.append([path.name,hit+1,'Индивидуальная выдача не распознана']); continue
            groups.append({'file':path.name,'order':order,'paragraph':hit+1,'heading':date_text+' '+text,'start':start.isoformat(),'days':duration,'unit':'А7383','basis':basis,'people':persons})
            continue
        end=next((j for j in range(hit+1,len(paragraphs)) if re.match(r'^\d+\.\s',paragraphs[j])),len(paragraphs))
        main=text
        start,days=period(main)
        dur=re.search(r'на\s+(\d+)\s*(?:\([^)]*\)\s*)?(?:доб|діб)',main,re.I)
        if dur: days=int(dur[1])
        unit_match=re.search(r'військової частини\s+([АA]\d{4})',main,re.I)
        base_unit=unit_match[1].replace('A','А') if unit_match else 'А7383'
        unit=base_unit
        people=[]
        heading=main
        declared=None
        for j in range(hit+1,end):
            p=paragraphs[j]
            if not p: continue
            m=PERSON.match(p)
            if m:
                name,note=resolve(m.groups()[1:])
                people.append({'rank':RANKS[m[1].lower()],'name':name,'original':' '.join(m.groups()[1:]),'paragraph':j+1,'note':note})
                continue
            initials=re.match(r'головного сержанта ЛЕБЕДЄВА\s+М\.\s*М\.?',p)
            if initials:
                candidates=[n for n in canonical if n.startswith('ЛЕБЕДЄВ М') and len(n.split())==3 and n.split()[2].startswith('М')]
                assert len(candidates)==1,candidates
                people.append({'rank':'головний сержант','name':candidates[0],'original':p,'paragraph':j+1,'note':'Инициалы М.М. раскрыты по единственному соответствию в справочниках'})
                continue
            if p.startswith('Підстава:'):
                if people:
                    if not start or not days: warnings.append([path.name,j+1,'Нет периода: '+heading])
                    else:
                        groups.append({'file':path.name,'order':order,'paragraph':hit+1,'heading':heading,'start':start.isoformat(),'days':days,'unit':unit,'basis':p,'people':people,'declared':declared})
                        if declared is not None and declared!=len(people): warnings.append([path.name,j+1,f'Число в приказе {declared}, извлечено {len(people)}'])
                    people=[]
                    declared=None
                continue
            count=re.fullmatch(r'\(?\s*(\d+)\s*(?:всл)?\s*\)?',p,re.I)
            if count: declared=int(count[1]); continue
            new_start,new_days=period(p)
            if new_start:
                start=new_start
                if new_days: days=new_days
                duration=re.search(r'на\s+(\d+)\s*(?:\([^)]*\)\s*)?(?:доб|діб)',p,re.I)
                if duration: days=int(duration[1])
                heading=main+' '+p
            code=re.search(r'військової частини\s+([АA]\d{4})',p,re.I)
            if code: unit=code[1].replace('A','А'); heading=main+' '+p
            elif re.search(r'(?:батальйону|роти)',p,re.I) and not people:
                unit=base_unit
                heading=main+' '+p
            if re.search(r'(?:солдат|сержант|капітан|майор|лейтенант)',p,re.I) and not code and not new_start:
                warnings.append([path.name,j+1,'Не распознан персональный абзац: '+p])
        if people: warnings.append([path.name,end,'Группа без основания',len(people)])

payload={'source':str(book_path),'files':[p.name for p in files],'existing':existing,'groups':groups,'warnings':warnings}
with gzip.open(OUT/'analysis.json.gz','wt',encoding='utf-8') as f: json.dump(payload,f,ensure_ascii=False)
unresolved=collections.Counter(p['original'] for g in groups for p in g['people'] if p['note'])
print(json.dumps({'files':len(files),'groups':len(groups),'people':sum(len(g['people']) for g in groups),'existing':len(existing),'warnings':warnings,'unresolved':list(unresolved.items()),'groups_summary':[(g['order'],g['start'],g['days'],g['unit'],len(g['people'])) for g in groups]},ensure_ascii=False,indent=2))