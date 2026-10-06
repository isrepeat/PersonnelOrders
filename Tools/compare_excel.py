import argparse
import collections
import datetime
import html
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def value(v):
    if v is None or v == '':
        return ''
    if isinstance(v, datetime.datetime):
        return v.isoformat(sep=' ') if v.time() != datetime.time() else v.strftime('%d.%m.%Y')
    if hasattr(v, 'text'):
        return v.text or ''
    return str(v)


def read(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    result = {}
    with zipfile.ZipFile(path) as z:
        for index, sheet in enumerate(wb):
            tree = ET.fromstring(z.read(f'xl/worksheets/sheet{index + 1}.xml'))
            cols = [column_index_from_string(''.join(filter(str.isalpha, c.attrib['r']))) for c in tree.findall('.//m:c', NS) if c.find('m:v', NS) is not None or c.find('m:f', NS) is not None or c.find('m:is', NS) is not None]
            width = max(cols, default=1)
            rows = [(i, tuple(value(c) for c in r)) for i, r in enumerate(sheet.iter_rows(max_col=width, values_only=True), 1)]
            result[sheet.title] = [(i, r) for i, r in rows if any(r)]
    wb.close()
    return result


def escape(v):
    return html.escape(str(v))


def content(text):
    if text.startswith('=') or len(text) > 350 or text.count('\n') > 3:
        short = text.split('\n')[0][:100]
        return f'<details class="formula"><summary>{escape(short)}…</summary><pre>{escape(text)}</pre></details>'
    return f'<pre>{escape(text)}</pre>'


def generate(args):
    old, new = map(read, [args.old, args.new])
    sections = []
    totals = collections.Counter()
    summary = []
    for sheet in dict.fromkeys([*old, *new]):
        aa, bb = old.get(sheet, []), new.get(sheet, [])
        width = max([len(r) for _, r in aa + bb], default=1)
        aa = [(i, r + ('',) * (width - len(r))) for i, r in aa]
        bb = [(i, r + ('',) * (width - len(r))) for i, r in bb]
        ha = next(((i, r) for i, r in aa if args.key in r), None)
        hb = next(((i, r) for i, r in bb if args.key in r), None)
        labels = list((hb or ha or (0, tuple(get_column_letter(c + 1) for c in range(width))))[1])
        labels = [label or get_column_letter(c + 1) for c, label in enumerate(labels)]
        table_header = hb or ha
        table_columns = [c for c in range(width) if table_header[1][c]] if table_header else list(range(width))
        columns = []
        if ha and hb:
            for c, (x, y) in enumerate(zip(ha[1], hb[1])):
                if x != y:
                    columns.append((get_column_letter(c + 1), x, y))
            if columns:
                raise RuntimeError('Изменены заголовки: требуется сопоставление столбцов по именам')
        # Сначала снимаем точные совпадения, включая переставленные и повторяющиеся строки.
        pool = collections.defaultdict(collections.deque)
        for j, (_, row) in enumerate(bb):
            pool[row].append(j)
        used_a, used_b = set(), set()
        for j, (_, row) in enumerate(aa):
            if pool[row]:
                used_a.add(j)
                used_b.add(pool[row].popleft())
        changes = []
        if ha and hb:
            name_col = ha[1].index(args.key)
            candidates = collections.defaultdict(list)
            for j, (rn, row) in enumerate(bb):
                if j not in used_b and rn > hb[0]:
                    candidates[row[name_col]].append(j)
            for j, (rn, row) in enumerate(aa):
                if j in used_a or rn <= ha[0] or not row[name_col]:
                    continue
                ranked = []
                for k in candidates[row[name_col]]:
                    if k in used_b:
                        continue
                    target = bb[k][1]
                    active = [c for c in range(width) if row[c] or target[c]]
                    score = sum(row[c] == target[c] for c in active) / max(1, len(active))
                    ranked.append((score, -abs(rn - bb[k][0]), k))
                ranked.sort(reverse=True)
                if ranked and ranked[0][0] >= args.threshold:
                    k = ranked[0][2]
                    ambiguous = len(ranked) > 1 and ranked[0][0] == ranked[1][0]
                    changes.append(('change', aa[j], bb[k], ambiguous))
                    used_a.add(j)
                    used_b.add(k)
        # Вне таблиц сравниваем оставшиеся строки по исходному номеру.
        remaining_b = {rn: j for j, (rn, _) in enumerate(bb) if j not in used_b}
        for j, (rn, row) in enumerate(aa):
            if j in used_a or (ha and rn > ha[0]):
                continue
            k = remaining_b.get(rn)
            if k is not None:
                changes.append(('change', aa[j], bb[k], False))
                used_a.add(j)
                used_b.add(k)
        changes += [('delete', aa[j], None, False) for j in range(len(aa)) if j not in used_a]
        changes += [('add', None, bb[j], False) for j in range(len(bb)) if j not in used_b]
        counts = collections.Counter(t for t, _, _, _ in changes)
        assert len(aa) + counts['add'] - counts['delete'] == len(bb)
        totals.update(counts)
        summary.append({'sheet': sheet, **{key: counts[key] for key in ['change', 'add', 'delete']}, 'old_nonempty': len(aa), 'new_nonempty': len(bb)})
        cards = []
        for kind, before, after, ambiguous in sorted(changes, key=lambda x: (x[2] or x[1])[0]):
            row = (after or before)[1]
            identity = row[ha[1].index(args.key)] if ha and args.key in ha[1] else sheet
            tag = {'change': 'Изменена строка', 'add': 'Добавлена строка', 'delete': 'Удалена строка'}[kind]
            location = f"{before[0] if before else '—'} → {after[0] if after else '—'}"
            lines = []
            for c in table_columns:
                x, y = before[1][c] if before else '', after[1][c] if after else ''
                unchanged = (kind == 'change' and x == y) or (not x and not y)
                status = 'unchanged' if unchanged else 'replacement' if x and y else 'insertion' if y else 'removal'
                cell_tag = {'unchanged': 'Без изменений', 'replacement': 'Заменено', 'insertion': 'Добавлено', 'removal': 'Удалено'}[status]
                addr = f'{get_column_letter(c + 1)}{(after or before)[0]}'
                history = f'<div class="previous"><span>Было · {get_column_letter(c + 1)}{before[0]}</span>{content(x)}</div>' if status == 'replacement' else ''
                current = '<pre>-- // --</pre>' if unchanged else content(y or x)
                lines.append(f'<div class="cell {status}"><div class="cell-heading"><b>{escape(labels[c])}</b><span>{addr}</span></div><span class="cell-tag">{cell_tag}</span>{history}<div class="current">{current}</div></div>')
            context = []
            for c, label in enumerate(labels):
                if label == 'Подія' and row[c]:
                    context.append(escape(row[c]))
            note = '<p class="warning">Несколько записей имеют одинаковую близость: проверьте сопоставление.</p>' if ambiguous else ''
            cards.append(f'<article data-kind="{kind}"><header><b>{escape(identity)}</b><span class="row-context">{" · ".join(context)}</span></header>{note}<div class="row-scroll"><div class="row-cells">{"".join(lines)}</div></div></article>')
        sections.append(f'<section id="s{len(sections)}"><h2>{escape(sheet)} <small>~{counts["change"]} / +{counts["add"]} / −{counts["delete"]}</small></h2>{"".join(cards) or "<p class=muted>Изменений данных и формул нет.</p>"}</section>')

    css = ".unchanged{background:#2b2520;border-color:#504237;color:#a89785}.unchanged .current pre{font-weight:400}body{margin:0;background:#f5f7fa;color:#202936;font:14px Segoe UI,Arial}main{max-width:1450px;margin:auto;padding:30px}h1{font-size:27px;margin:0 0 12px}p{line-height:1.6}.muted,small{color:#697586}nav{position:sticky;top:0;background:#f5f7faf5;padding:14px 0;display:flex;flex-wrap:wrap;gap:8px;z-index:1;border-bottom:1px solid #d7dfe8}a,button,select,input{font:inherit;border:1px solid #ccd5df;border-radius:6px;padding:8px 12px;background:white;color:#253858}a{text-decoration:none}input{min-width:240px}h2{margin:32px 0 14px}h2 small{font-size:14px;font-weight:400}article{border:1px solid #d4dce5;border-radius:8px;overflow:hidden;margin:16px 0;background:white}header{display:flex;justify-content:space-between;gap:16px;padding:14px 16px;background:#eef2f7}header span{color:#5a6778}.context{padding:0 16px;color:#5a6778;font-size:12px}.line{display:grid;grid-template-columns:26px 65px minmax(150px,240px) minmax(0,1fr);padding:8px 12px;gap:8px;border-top:1px solid #ffffff80}.del{background:#ffeef0}.add{background:#e6ffed}.sign{font:bold 16px Consolas}.del .sign{color:#b42332}.add .sign{color:#087b35}.addr{font:12px Consolas;color:#667085}.label{font-size:12px;font-weight:600}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:13px Consolas,monospace;margin:0;line-height:1.5}.stats{display:flex;gap:12px;flex-wrap:wrap;margin:20px 0}.stats b{background:white;border:1px solid #d4dce5;padding:14px 20px;border-radius:8px;font-size:18px}.warning{color:#935b00;padding:0 16px}details{background:white;border:1px solid #d4dce5;padding:12px;border-radius:8px}summary{cursor:pointer}table{border-collapse:collapse}td,th{text-align:left;padding:8px 20px 8px 0}@media(max-width:700px){main{padding:15px}.line{grid-template-columns:20px 50px 1fr}.line pre{grid-column:3}header{flex-direction:column}}@media print{nav{display:none}article{break-inside:avoid}}body{background:#211c19;color:#eee5dc;color-scheme:dark}.muted,small,h2 small{color:#b9a99a}nav{background:#211c19f5;border-bottom-color:#514236}a,button,select,input{background:#302721;color:#eadbcb;border-color:#5b4839}a:hover,button:hover{background:#443329;border-color:#b58961}a:focus-visible,button:focus-visible,select:focus-visible,input:focus-visible{outline:2px solid #d3a574;outline-offset:2px}input::placeholder{color:#b6a390}article{background:#29231f;border-color:#514236}header{background:#352b24}header span,.context{color:#c2b09e}.line{border-top-color:#ffffff0a}.del{background:#412725;color:#f2d2cd}.add{background:#26372a;color:#d5ead4}.del .sign{color:#f19a90}.add .sign{color:#9bd39b}.addr{color:#bba996}.label{color:#d5bda6}.stats b,details{background:#2d251f;border-color:#514236}.stats b{color:#e8c39f}.warning{color:#efbf77}summary{color:#e3c3a3}section{scroll-margin-top:130px}::selection{background:#8b6040;color:#fff4e6}.row-scroll{overflow-x:auto;padding:0 14px 14px;scrollbar-color:#81624a #29231f}.row-cells{display:flex;align-items:flex-start;gap:10px;min-width:min-content}.cell{flex:0 0 240px;width:240px;box-sizing:border-box;padding:12px;border:1px solid;border-radius:6px;overflow-wrap:anywhere}.cell-heading{display:flex;justify-content:space-between;gap:10px;font-size:12px;margin-bottom:8px}.cell-heading span{font:11px Consolas;color:#c9b9a8}.cell-tag{display:inline-block;font-size:10px;letter-spacing:.03em;margin-bottom:10px;color:#d4c5b8}.insertion{background:#263c2d;border-color:#41694c;color:#daf0dc}.replacement{background:#3b2d49;border-color:#695079;color:#efdefa}.removal{background:#452b28;border-color:#7c4941;color:#f2d2cd}.previous{padding-bottom:10px;margin-bottom:10px;border-bottom:1px solid #ffffff20;color:#beabc9}.previous>span{display:block;font-size:10px;margin-bottom:6px}.previous pre{font-size:12px}.current pre{font-weight:600}.formula{background:#00000012!important;border:1px solid #ffffff20!important;padding:7px!important;min-width:0}.formula summary{font:12px Consolas;color:inherit;overflow-wrap:anywhere}.formula pre{margin-top:10px;font-weight:400;max-height:320px;overflow:auto}.legend{display:flex;gap:18px;flex-wrap:wrap;font-size:12px;color:#c9b9a8;margin:14px 0}.legend span:before{content:\"\";display:inline-block;width:10px;height:10px;margin-right:7px;border-radius:2px}.legend .green:before{background:#598565}.legend .purple:before{background:#946caf}.legend .red:before{background:#b46f61}@media print{.row-scroll{overflow:visible}.row-cells{flex-wrap:wrap;min-width:0}}main{max-width:none;width:100%;box-sizing:border-box;margin:0;padding:16px}h1{font-size:22px;margin-bottom:6px}p{line-height:1.4}.stats{gap:8px;margin:12px 0}.stats b{padding:8px 12px;font-size:15px}.legend{gap:14px;margin:10px 0}details{padding:8px}nav{padding:9px 0;gap:6px}a,button,select,input{padding:6px 9px;font-size:12px}h2{margin:20px 0 10px;font-size:19px}article{margin:10px 0;border-radius:6px}header{padding:9px 11px;font-size:12px}.context{margin:7px 0;padding:0 11px;font-size:11px}.row-scroll{padding:0 9px 9px}.row-cells{gap:6px}.cell{flex-basis:170px;width:170px;padding:8px;border-radius:4px}.cell-heading{font-size:11px;gap:6px;margin-bottom:5px}.cell-heading span{font-size:10px}.cell-tag{font-size:9px;margin-bottom:6px}.unchanged .cell-tag{display:none}pre{font-size:12px;line-height:1.35}.previous{padding-bottom:6px;margin-bottom:6px}.previous>span{margin-bottom:4px}.previous pre{font-size:11px}.formula{padding:5px!important}.formula summary{font-size:11px}section{scroll-margin-top:95px}@media(max-width:700px){main{padding:10px}header{gap:5px}.cell{flex-basis:150px;width:150px}}.row-cells{align-items:stretch}.cell{height:96px;min-height:96px;max-height:96px;overflow:hidden;cursor:pointer;position:relative;flex:0 0 170px;width:170px}.cell-heading{height:28px;overflow:hidden;margin-bottom:4px}.cell-heading b{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}.cell-heading span{flex-shrink:0;white-space:nowrap}.cell>.previous{display:none}.cell>.current{max-height:30px;overflow:hidden}.cell>.current pre,.cell>.current .formula summary{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}.cell .formula{pointer-events:none}.cell:hover{filter:brightness(1.14)}.cell:focus-visible{outline:2px solid #e0b080;outline-offset:-3px}.cell-dialog{width:min(850px,calc(100vw - 40px));max-height:80vh;box-sizing:border-box;background:#2d251f;color:#eee5dc;border:1px solid #81624a;border-radius:10px;padding:20px;overflow:auto}.cell-dialog::backdrop{background:#000a;backdrop-filter:blur(2px)}.dialog-top{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;margin-bottom:12px}.dialog-top h3{margin:0;font-size:18px}.dialog-close{cursor:pointer;font-size:20px;padding:2px 10px}.dialog-context{font-size:12px;color:#bca995;margin:0 0 16px}.dialog-body .cell-heading{height:auto;font-size:14px;overflow:visible;margin:0 0 10px}.dialog-body .cell-heading b{display:block}.dialog-body .previous{display:block;color:#ccb7d9}.dialog-body .current{padding:12px;background:#211c19;border-radius:5px}.dialog-body pre{display:block;white-space:pre-wrap;overflow-wrap:anywhere;font-size:14px;line-height:1.5;font-weight:400;max-height:none;overflow:visible}.dialog-body .formula{pointer-events:auto}.dialog-body .formula pre{max-height:none}.dialog-body .cell-tag{font-size:12px;margin-bottom:12px}@media(max-width:700px){.cell{flex-basis:170px;width:170px}.cell-dialog{padding:14px}}html{overflow-x:auto}body{overflow-x:visible}main{width:max-content;min-width:100%;box-sizing:border-box}.row-scroll{overflow:visible;scrollbar-width:none}.row-cells{width:max-content;min-width:0}article{overflow:visible}.cell-dialog{position:fixed}nav{max-width:calc(100vw - 32px)}@media(max-width:700px){nav{max-width:calc(100vw - 20px)}}article>header{justify-content:flex-start;align-items:center;flex-wrap:nowrap;padding:6px 9px;gap:14px;white-space:nowrap;font-size:11px}article>header>b{flex-shrink:0}article>header>.row-context{font-size:10px;color:#bca995;flex:1;overflow:hidden;text-overflow:ellipsis}article>header>.row-location{flex-shrink:0;font-size:10px}article>.row-scroll{padding-top:6px}article{margin:7px 0}@media(max-width:700px){article>header{flex-direction:row;gap:8px}}body{padding-bottom:54px}.page-footer{position:fixed;bottom:0;left:0;right:0;height:44px;box-sizing:border-box;background:#302720;border-top:1px solid #68503d;display:flex;align-items:stretch;justify-content:center;gap:0;z-index:10;box-shadow:0 -3px 12px #0004}.page-footer button{min-width:160px;border:0;border-right:1px solid #504032;border-radius:0;background:transparent;color:#c7b5a4;padding:0 20px;cursor:pointer;font-size:13px}.page-footer button.active{background:#4b3729;color:#ffdfbc;box-shadow:inset 0 3px #c29668}.page-footer button:hover{background:#443329}nav>a{display:none}section[hidden]{display:none!important}@media(max-width:700px){.page-footer button{min-width:0;flex:1;padding:0 10px}}.cell>.cell-heading>span{display:none}article>header>.row-context{flex:0 1 auto}.dialog-body .cell-heading>span{display:inline}article>header{gap:0}article>header>.row-context{font:inherit;font-weight:700;color:inherit;display:flex;align-items:center}article>header>.row-context:not(:empty)::before{content:'—';display:inline-block;margin:0 12px;font:inherit;font-weight:700;flex-shrink:0}article>header>.row-context{color:#bca995}.column-settings{max-width:calc(100vw - 32px);box-sizing:border-box;margin:8px 0}.column-controls{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:8px;padding:12px 0;max-height:270px;overflow:auto}.column-controls label{display:grid;grid-template-columns:minmax(80px,1fr) 100px 65px;gap:8px;align-items:center;font-size:11px}.column-controls input{min-width:0;width:100%;box-sizing:border-box;padding:4px}.column-controls input[type=range]{padding:0;accent-color:#c29668}.column-settings[data-save-error]:after{content:attr(data-save-error);display:block;color:#efbf77;margin-top:8px}.reset-widths{cursor:pointer}.column-controls label>span{text-align:right}.page-footer{height:auto;flex-direction:column;align-items:stretch}.footer-tabs{display:flex;justify-content:center;height:44px;flex-shrink:0}.page-footer>.column-settings{width:100%;max-width:none;margin:0;padding:7px 16px;box-sizing:border-box;border-top:1px solid #504032}.page-footer>.column-settings>summary{cursor:pointer;font-size:12px}.page-footer .column-controls{max-height:min(230px,35vh)}.page-footer .reset-widths{height:30px;min-width:0;border:1px solid #68503d;border-radius:4px;font-size:11px}.page-footer .column-controls input{background:#302721}.page-footer .column-controls label>span{text-align:right}.page-footer button:focus,.page-footer button:focus-visible{outline:none;outline-offset:0}.settings-actions{display:flex;align-items:center;gap:8px;margin-top:8px}.page-footer .settings-actions button{height:30px;min-width:0;border:1px solid #68503d;border-radius:4px;font-size:11px;padding:0 12px}.settings-actions span{font-size:11px;color:#c7b5a4}"
    nav = ''.join(f'<a href="#s{i}">{escape(s["sheet"])}</a>' for i, s in enumerate(summary))
    table = '<table><tr><th>Лист</th><th>Изменено</th><th>Добавлено</th><th>Удалено</th></tr>' + ''.join(f'<tr><td>{escape(s["sheet"])}</td><td>{s["change"]}</td><td>{s["add"]}</td><td>{s["delete"]}</td></tr>' for s in summary) + '</table>'
    script = '''const q=document.querySelector('input'),f=document.querySelector('select');function filter(){for(const a of document.querySelectorAll('article'))a.hidden=!(a.textContent.toLowerCase().includes(q.value.toLowerCase())&&(!f.value||a.dataset.kind===f.value));}q.oninput=filter;f.onchange=filter;
const dialog=document.createElement('dialog');
dialog.className='cell-dialog';
dialog.innerHTML='<div class="dialog-top"><h3 id="cell-dialog-title"></h3><button type="button" class="dialog-close" aria-label="Закрыть">×</button></div><p class="dialog-context"></p><div class="dialog-body"></div>';
dialog.setAttribute('aria-labelledby','cell-dialog-title');
document.body.append(dialog);
dialog.querySelector('.dialog-close').onclick=()=>dialog.close();
dialog.addEventListener('click',e=>{const r=dialog.getBoundingClientRect();if(e.target===dialog&&(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom))dialog.close();});
function openCell(cell){
 const article=cell.closest('article');
 dialog.querySelector('h3').textContent=cell.querySelector('.cell-heading b').textContent;
 dialog.querySelector('.dialog-context').textContent=article.querySelector('header').textContent+' · '+article.closest('section').querySelector('h2').textContent;
 const body=dialog.querySelector('.dialog-body');
 body.replaceChildren();
 for(const child of cell.children)body.append(child.cloneNode(true));
 for(const detail of body.querySelectorAll('details'))detail.open=true;
 dialog.showModal();
}
for(const cell of document.querySelectorAll('.cell')){
 cell.setAttribute('role','button');cell.tabIndex=0;
 cell.setAttribute('aria-haspopup','dialog');
 cell.setAttribute('aria-label',cell.querySelector('.cell-heading b').textContent+' — открыть подробности');
 cell.addEventListener('click',e=>{e.preventDefault();openCell(cell);});
 cell.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();openCell(cell);}});
}
'''
    script += "\nconst pageFooter=document.createElement('footer');\npageFooter.className='page-footer';\npageFooter.setAttribute('aria-label','Листы отчёта');\nconst pageSections=[...document.querySelectorAll('main>section')];\nconst pageNames=[['Список','Список'],['Відсутні','Отсутствующие'],['Прибулі','Прибывшие']];\nfunction selectPage(section,button){\n for(const item of pageSections)item.hidden=item!==section;\n for(const tab of pageFooter.querySelectorAll('button')){const active=tab===button;tab.classList.toggle('active',active);tab.setAttribute('aria-pressed',String(active));}\n history.replaceState(null,'','#'+section.id);\n window.scrollTo(0,0);\n}\nfor(const [sheet,label] of pageNames){\n const section=pageSections.find(s=>s.querySelector('h2').firstChild.textContent.trim()===sheet);\n if(!section)continue;\n const button=document.createElement('button');button.type='button';button.textContent=label;\n button.onclick=()=>selectPage(section,button);\n pageFooter.append(button);\n}\ndocument.body.append(pageFooter);\nconst initial=pageSections.find(s=>'#'+s.id===location.hash&&pageNames.some(([name])=>s.querySelector('h2').firstChild.textContent.trim()===name))||pageSections.find(s=>s.querySelector('h2').firstChild.textContent.trim()==='Список');\nif(initial){const index=pageNames.findIndex(([name])=>initial.querySelector('h2').firstChild.textContent.trim()===name);selectPage(initial,[...pageFooter.children][index]);}\n"
    script += "\nconst sizing=document.createElement('details');sizing.className='column-settings';sizing.innerHTML='<summary>Ширина столбцов</summary><div class=\"column-controls\"></div><button type=\"button\" class=\"reset-widths\">Сбросить размеры листа</button>';\ndocument.querySelector('nav').before(sizing);\nconst widthStoreKey='excel-diff-column-widths-v1';let savedWidths={};\ntry{savedWidths=JSON.parse(localStorage.getItem(widthStoreKey)||'{}');if(!savedWidths||typeof savedWidths!=='object')savedWidths={};}catch(e){savedWidths={};}\nfunction sheetName(section){return section.querySelector('h2').firstChild.textContent.trim();}\nfunction columnNames(section){return [...(section.querySelector('article .row-cells')?.children||[])].map(c=>c.querySelector('.cell-heading b').textContent);}\nfunction applyWidth(section,index,width){\n for(const article of section.querySelectorAll('article')){const cell=article.querySelector('.row-cells')?.children[index];if(cell){cell.style.width=width+'px';cell.style.flexBasis=width+'px';}}\n}\nfor(const section of pageSections){const values=savedWidths[sheetName(section)]||{};columnNames(section).forEach((name,index)=>{const value=Number(values[name]);if(Number.isFinite(value)&&value>=60&&value<=600)applyWidth(section,index,value);});}\nfunction currentSection(){return pageSections.find(s=>!s.hidden);}\nfunction saveWidths(){try{localStorage.setItem(widthStoreKey,JSON.stringify(savedWidths));sizing.removeAttribute('data-save-error');}catch(e){sizing.setAttribute('data-save-error','Браузер запретил сохранение размеров');}}\nfunction renderSizing(){\n const section=currentSection();const controls=sizing.querySelector('.column-controls');controls.replaceChildren();if(!section)return;\n const sheet=sheetName(section);const values=savedWidths[sheet]||{};\n columnNames(section).forEach((name,index)=>{\n const label=document.createElement('label');const title=document.createElement('span');title.textContent=name;\n const slider=document.createElement('input');slider.type='range';slider.min='60';slider.max='600';slider.step='1';slider.value=values[name]||170;slider.setAttribute('aria-label','Ширина '+name);\n const number=document.createElement('input');number.type='number';number.min='60';number.max='600';number.step='1';number.value=slider.value;number.setAttribute('aria-label','Ширина '+name+' в пикселях');\n function update(value){const width=Math.min(600,Math.max(60,Number(value)));if(!Number.isFinite(width))return;slider.value=width;number.value=width;applyWidth(section,index,width);savedWidths[sheet]||={};savedWidths[sheet][name]=width;saveWidths();}\n slider.oninput=()=>update(slider.value);number.oninput=()=>{if(number.value!==''&&Number(number.value)>=60&&Number(number.value)<=600)update(number.value);};number.onchange=()=>update(number.value);\n label.append(title,slider,number);controls.append(label);\n });\n}\nsizing.querySelector('.reset-widths').onclick=()=>{const section=currentSection();if(!section)return;delete savedWidths[sheetName(section)];columnNames(section).forEach((name,index)=>applyWidth(section,index,170));saveWidths();renderSizing();};\npageFooter.addEventListener('click',e=>{if(e.target.closest('.footer-tabs button'))renderSizing();});renderSizing();\n"
    script += "\nconst footerTabs=document.createElement('div');footerTabs.className='footer-tabs';\nfor(const button of [...pageFooter.children]){if(button.tagName==='BUTTON')footerTabs.append(button);}\npageFooter.prepend(footerTabs);\nconst footerSizing=sizing;pageFooter.append(footerSizing);\nfunction updateFooterSpace(){document.body.style.paddingBottom=(pageFooter.offsetHeight+10)+'px';}\nfooterSizing.addEventListener('toggle',updateFooterSpace);\nnew ResizeObserver(updateFooterSpace).observe(pageFooter);updateFooterSpace();\n"
    script += "\nconst settingsActions=document.createElement('div');settingsActions.className='settings-actions';\nconst exportWidths=document.createElement('button');exportWidths.type='button';exportWidths.textContent='Экспорт JSON';\nconst importWidths=document.createElement('button');importWidths.type='button';importWidths.textContent='Импорт JSON';\nconst settingsMessage=document.createElement('span');settingsMessage.setAttribute('role','status');\nconst settingsFile=document.createElement('input');settingsFile.type='file';settingsFile.accept='.json,application/json';settingsFile.hidden=true;\nsettingsActions.append(exportWidths,importWidths,settingsMessage,settingsFile);footerSizing.append(settingsActions);\nexportWidths.onclick=()=>{\n const allWidths=Object.create(null);\n for(const section of pageSections){\n const sheet=sheetName(section);allWidths[sheet]=Object.create(null);\n const cells=section.querySelector('article .row-cells')?.children||[];\n columnNames(section).forEach((name,index)=>{const width=Number.parseFloat(cells[index]?.style.width)||Number(savedWidths[sheet]?.[name])||170;allWidths[sheet][name]=width;});\n }\n const payload={version:1,columnWidths:allWidths};\n const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});\n const url=URL.createObjectURL(blob);const link=document.createElement('a');\n link.href=url;link.download='excel-column-widths.json';document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);\n settingsMessage.textContent='Настройки экспортированы';\n};\nimportWidths.onclick=()=>settingsFile.click();\nsettingsFile.onchange=async()=>{\n const file=settingsFile.files[0];if(!file)return;\n try{\n if(file.size>1048576)throw new Error('Файл настроек слишком большой');\n const payload=JSON.parse(await file.text());\n if(payload.version!==1||!payload.columnWidths||typeof payload.columnWidths!=='object'||Array.isArray(payload.columnWidths))throw new Error('Неверный формат настроек');\n const validated=Object.create(null);\n for(const [sheet,columns] of Object.entries(payload.columnWidths)){\n if(!columns||typeof columns!=='object'||Array.isArray(columns))throw new Error('Неверные настройки листа');\n validated[sheet]=Object.create(null);\n for(const [name,width] of Object.entries(columns)){\n if(typeof width!=='number'||!Number.isFinite(width)||width<60||width>600)throw new Error('Ширина должна быть числом от 60 до 600');\n validated[sheet][name]=width;\n }\n }\n savedWidths=validated;\n for(const section of pageSections){const values=savedWidths[sheetName(section)]||{};columnNames(section).forEach((name,index)=>applyWidth(section,index,values[name]??170));}\n saveWidths();renderSizing();updateFooterSpace();\n settingsMessage.textContent='Настройки импортированы';\n }catch(error){settingsMessage.textContent='Ошибка импорта: '+error.message;}\n finally{settingsFile.value='';}\n};\n"
    document = f'<!doctype html><html lang="ru"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>РУХ · сравнение Excel</title><style>{css}</style><main><h1>РУХ · изменения между версиями</h1><p class="muted">{escape(args.old.name)} → {escape(args.new.name)} · − старая версия / + новая версия</p><div class="stats"><b>~ {totals["change"]} изменено</b><b>+ {totals["add"]} добавлено</b><b>− {totals["delete"]} удалено</b><b>0 добавленных / удалённых столбцов</b></div><div class="legend"><span class="green">Добавлено значение</span><span class="purple">Заменено значение · внутри показано «Было»</span><span class="red">Удалено значение</span></div><details><summary>Сводка и способ сравнения</summary>{table}<p>Сравниваются сохранённые значения и тексты формул на всех листах. Оформление и кэшированные результаты вычислений не сравниваются. Пустые строки и пустые оформленные столбцы не считаются записями. Заголовки таблиц совпадают.</p><p>Сначала сопоставлены полностью одинаковые строки независимо от позиции с учётом повторов. Остальные записи сопоставлены по полю {escape(args.key)} и наибольшей доле совпадающих заполненных полей (не менее {args.threshold:.0%}). Это эвристика: одинаковые по близости кандидаты отмечены для проверки. Строки вне таблиц сопоставлены по номеру. Перестановки одинаковых записей не показаны как изменения.</p></details><nav>{nav}<input placeholder="Поиск по ФИО, полю или значению"><select><option value="">Все изменения</option><option value="change">Изменённые</option><option value="add">Добавленные</option><option value="delete">Удалённые</option></select></nav>{"".join(sections)}</main><script>{script}</script></html>'
    return document, totals


def main():
    parser = argparse.ArgumentParser(description='Офлайн-сравнение Excel с горизонтальным HTML-отчётом.')
    parser.add_argument('old', type=Path, help='Предыдущая версия')
    parser.add_argument('new', type=Path, help='Новая версия')
    parser.add_argument('-o', '--output', type=Path, help='Путь отчёта; по умолчанию Results/ExcelComparing/<имя новой книги>.diff.html')
    parser.add_argument('--key', default='ПІБ', help='Название поля для сопоставления строк')
    parser.add_argument('--threshold', type=float, default=0.65, help='Минимальная доля совпадений, от 0 до 1')
    parser.add_argument('--stdout', action='store_true', help='Вывести HTML вместо записи файла')
    args = parser.parse_args()
    if args.output is None:
        args.output = Path(__file__).resolve().parent.parent / 'Results' / 'ExcelComparing' / f'{args.new.stem}.diff.html'
    if not 0 < args.threshold <= 1:
        parser.error('--threshold должен быть больше 0 и не больше 1')
    if args.output.resolve() in [args.old.resolve(), args.new.resolve()]:
        parser.error('Отчёт не может перезаписывать исходную книгу')
    try:
        document, totals = generate(args)
        if args.stdout:
            sys.stdout.reconfigure(encoding='utf-8')
            print(document, end='')
        else:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(document, encoding='utf-8')
            print(f'Отчёт: {args.output.resolve()}')
            print(f'Изменено: {totals["change"]}; добавлено: {totals["add"]}; удалено: {totals["delete"]}')
    except PermissionError:
        parser.exit(1, 'Файл заблокирован. Сохраните и закройте книгу в Excel.\n')
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as error:
        parser.exit(1, f'Ошибка: {error}\n')


if __name__ == '__main__':
    main()