"""Сверяет пять полных ПІБ в архиве приказов 2026 года, включая майские основания учёта."""

import gzip
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree as ET

sys.stdout.reconfigure(encoding='utf-8')
root = Path(__file__).resolve().parents[1]
targets = {
    'ШКЛЯРЕНКО Микола Миколайович': ('ШКЛЯРЕН', 'МИКОЛ', 'МИКОЛАЙОВ'),
    'ЯРУШАК Михайло Осипович': ('ЯРУША', 'МИХАЙЛ', 'ОСИПОВ'),
    'ШИШКОВ Володимир Володимирович': ('ШИШКОВ', 'ВОЛОДИМИР', 'ВОЛОДИМИРОВ'),
    'ПАВЛОВ Максим Сергійович': ('ПАВЛОВ', 'МАКСИМ', 'СЕРГІЙОВ'),
    'НАГОРНЯК В’ячеслав Олександрович': ('НАГОРНЯК', 'ЯЧЕСЛАВ', 'ОЛЕКСАНДРОВ'),
}
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
hits, files, errors = [], [], []
for path in sorted((root / 'Documents/Orders').rglob('*.docx')):
    match = re.match(r'(2026-\d{2}-\d{2})', path.name)
    if not match or match[1] < '2026-01-01':
        continue
    files.append(str(path.relative_to(root)))
    try:
        with zipfile.ZipFile(path) as archive:
            tree = ET.fromstring(archive.read('word/document.xml'))
        paragraphs = [''.join(p.xpath('.//w:t/text()', namespaces=ns)).strip() for p in tree.findall('.//w:p', ns)]
        headings = []
        for index, paragraph in enumerate(paragraphs):
            normalized = paragraph.upper().replace('Ё', 'Е')
            if re.search(r'продовольч|котлов|сух\w*\s+пай|зарахува|зняти|виключити|Вважати|з\s+\d{1,2}\s+\w+\s+2026', paragraph, re.I):
                headings.append(index)
            for name, stems in targets.items():
                pattern = re.escape(stems[0]) + r'\S*\s+\S*' + re.escape(stems[1]) + r'\S*\s+' + re.escape(stems[2]) + r'\S*'
                if not re.search(pattern, normalized):
                    continue
                preceding = headings[-7:]
                context = [{'paragraph': i + 1, 'text': paragraphs[i]} for i in sorted(set(preceding + list(range(max(0, index - 3), min(len(paragraphs), index + 4))))) if paragraphs[i]]
                hit = dict(name=name, date=match[1], file=str(path.relative_to(root)), paragraph=index + 1, text=paragraph, context=context)
                hits.append(hit)
                print(match[1], name, path.name, index + 1)
    except Exception as error:
        errors.append(dict(file=str(path), error=str(error)))
out = root / 'Tasks/01.Food/Work/2026.10.14/order_audit.json.gz'
with gzip.open(out, 'wt', encoding='utf-8') as stream:
    json.dump(dict(files=files, hits=hits, errors=errors), stream, ensure_ascii=False)
print(json.dumps(dict(scanned=len(files), hits=len(hits), errors=errors), ensure_ascii=False))