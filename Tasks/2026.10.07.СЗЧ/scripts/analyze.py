import argparse, json, datetime, pathlib, re, sys
import openpyxl

parser = argparse.ArgumentParser(description='Сверка сентябрьских зачислений с источниками СЗЧ и историей ТП')
parser.add_argument('--sources-dir', default=str(pathlib.Path(__file__).resolve().parent.parent))
parser.add_argument('--movement', required=True)
parser.add_argument('--month', default='2026-09')
parser.add_argument('--full', action='store_true')
args = parser.parse_args()
if not re.fullmatch(r'\d{4}-\d{2}', args.month):
    raise ValueError('Месяц должен иметь формат YYYY-MM')
root = pathlib.Path(args.sources_dir)
paths = [root / name for name in [
    '0000. СЗЧ по ухвалі суду 05_10_2026.xlsx',
    '00000. СЗЧ_через_БРЕЗ_новий_алгоритм_на 05.10.2026.xlsx',
    'Поіменний список рапортів СЗЧ (1).xlsm'
]] + [pathlib.Path(args.movement)]
for path in paths:
    if not path.is_file():
        raise FileNotFoundError(path)
data = []
for path in paths:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    sheets = []
    for ws in wb:
        if path == paths[-1] and ws.title not in ['Список', 'Прибулі']:
            continue
        rows = [[v.isoformat() if isinstance(v,(datetime.datetime,datetime.date)) else v for v in row] for row in ws.values]
        if path != paths[-1] or ws.title in ['Список','Прибулі']:
            sheets.append({'name':ws.title,'rows':rows,'tables':{t.name:t.ref for t in ws.tables.values()}})
    data.append({'file':str(path),'sheets':sheets})
import difflib
def norm(v):
    return re.sub(r'[^А-ЯІЇЄҐA-Z0-9]', '', str(v or '').upper().replace('Ґ','Г').replace('Ё','Е'))
def ident(v):
    return str(v or '').strip().removesuffix('.0')
sources=[]
for idx,d in enumerate(data[:3]):
    s=d['sheets'][0]
    nc,ic,start = (1,2,1) if idx==0 else ((2,10,1) if idx==1 else (2,3,2))
    for rn,r in enumerate(s['rows'][start:],start+1):
        if isinstance(r[nc],str) and len(r[nc].split())>=3 and not re.search(r'\d',r[nc]):
            sources.append(dict(name=r[nc].strip(),ipn=ident(r[ic]),source=pathlib.Path(d['file']).name,sheet=s['name'],row=rn,key=norm(r[nc])))
movement=[next(s for s in data[3]['sheets'] if s['name']==name) for name in ['Список','Прибулі']]
people=[dict(name=r[1],ipn=ident(r[2]),enroll=r[3],row=i) for i,r in enumerate(movement[0]['rows'],1) if isinstance(r[3],str) and r[3].startswith(args.month)]
tp=[dict(name=r[2],arrival=r[7],departure=r[10],note=r[18],row=i,basis=r[13],depart_basis=r[15]) for i,r in enumerate(movement[1]['rows'][5:],6) if r[2]]
for p in people:
    exact=[s for s in sources if s['key']==norm(p['name'])]
    byid=[s for s in sources if p['ipn'] and s['ipn']==p['ipn']]
    best=sorted(sources,key=lambda s:difflib.SequenceMatcher(None,norm(p['name']),s['key']).ratio(),reverse=True)[:3]
    p['matches']=exact or byid
    p['match_type']='ФИО' if exact else ('ИПН' if byid else 'Нет')
    p['candidates']=[dict(s,score=round(difflib.SequenceMatcher(None,norm(p['name']),s['key']).ratio(),3)) for s in best]
    p['events']=[t for t in tp if norm(t['name'])==norm(p['name'])]
    if not p['events'] and p['matches']:
        p['tp_candidates']=sorted([(round(difflib.SequenceMatcher(None,norm(p['name']),norm(t['name'])).ratio(),3),t['name'],t['row']) for t in tp],reverse=True)[:3]
    events=sorted([t for t in p['events'] if t['arrival'] and t['arrival']<=p['enroll']],key=lambda t:(t['arrival'],t['row']))
    selected=events[-1] if events else None
    p['latest']=selected
    while selected:
        prev=[t for t in events if t['row']!=selected['row'] and (t['arrival'],t['row'])<(selected['arrival'],selected['row']) and t['departure']==selected['arrival'] and 'ЗМІНАСТАТУС' in norm(t['note'])]
        if not prev: break
        selected=prev[-1]
    p['selected']=selected
    p['final']=bool(p['matches'] and selected and selected['arrival'].startswith(args.month))
result={'people':people,'source_counts':[sum(s['source']==pathlib.Path(d['file']).name for s in sources) for d in data[:3]]}
if args.full:
    print(json.dumps(result,ensure_ascii=False))
else:
    print('COUNTS',len(people),sum(bool(p['matches']) for p in people),result['source_counts'])
    print('FINAL',sum(p['final'] for p in people))
    from collections import Counter
    print('DUPLICATES',[(k,v) for k,v in Counter((norm(p['name']),p['ipn']) for p in people).items() if v>1])
    for p in people:
        if not p['matches'] and p['candidates'][0]['score']>=.89:
            print('CANDIDATE',p['name'],p['ipn'],[(s['name'],s['ipn'],s['score']) for s in p['candidates'][:2]])
        if p['matches'] and (not p['events'] or p['match_type']!='ФИО'):
            print('SPECIAL',p['name'],p['match_type'],p.get('tp_candidates'),[(s['name'],s['ipn']) for s in p['matches']])
        if p['final']:
            print('FINAL_PERSON',p['name'],p['selected']['arrival'],p['selected']['row'])
        if any(s['ipn'] and p['ipn'] and s['ipn']!=p['ipn'] for s in p['matches']):
            print('ID_CONFLICT',p['name'],p['ipn'],[(s['name'],s['ipn']) for s in p['matches']])