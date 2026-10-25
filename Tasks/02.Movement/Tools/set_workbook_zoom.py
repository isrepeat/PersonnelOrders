"""Сохранение масштаба листов и метаданных формулы прекращения сухпайка."""
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


def set_zoom(path):
    path = Path(path)
    with ZipFile(path) as source:
        entries = [(entry, source.read(entry.filename)) for entry in source.infolist()]
    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    formula_text = 'DryRationEvents[[#This Row],[Початок]]+DryRationEvents[[#This Row],[Тривалість]]'
    has_dry_formulas = any(
        'Початок' in (formula.text or '') and 'Тривалість' in (formula.text or '')
        for entry, data in entries
        if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', entry.filename)
        for formula in ET.fromstring(data).iter(f'{{{namespace}}}f')
    )
    temporary = path.with_suffix('.zoom.tmp')
    with ZipFile(temporary, 'w') as target:
        for entry, data in entries:
            if entry.filename == 'xl/styles.xml':
                ET.register_namespace('', namespace)
                root = ET.fromstring(data)
                fills = root.find(f'{{{namespace}}}fills')
                fill = ET.SubElement(fills, f'{{{namespace}}}fill')
                pattern = ET.SubElement(fill, f'{{{namespace}}}patternFill', patternType='solid')
                ET.SubElement(pattern, f'{{{namespace}}}fgColor', rgb='FF28251E')
                ET.SubElement(pattern, f'{{{namespace}}}bgColor', indexed='64')
                fills.set('count', str(len(fills)))
                # Стиль по умолчанию покрывает весь лист, включая неиспользованные ячейки.
                default = root.find(f'{{{namespace}}}cellXfs')[0]
                default.set('fillId', str(len(fills) - 1))
                default.set('applyFill', '1')
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            elif re.fullmatch(r'xl/worksheets/sheet\d+\.xml', entry.filename):
                text = data.decode('utf-8')
                def update(match):
                    tag = re.sub(r'\szoomScale(?:Normal)?="[^"]*"', '', match.group())
                    end = '/>' if tag.endswith('/>') else '>'
                    return tag[:-len(end)] + ' zoomScale="80" zoomScaleNormal="80"' + end
                text, count = re.subn(r'<(?:\w+:)?sheetView\b[^>]*>', update, text)
                # Excel хранит ссылки текущей строки в полном формате, без сокращения @.
                def normalize_dry_formula(match):
                    if 'Початок' in match[2] and 'Тривалість' in match[2]:
                        return match[1] + formula_text + match[3]
                    return match[0]
                text = re.sub(r'(<(?:\w+:)?f\b[^>]*>)(.*?)(</(?:\w+:)?f>)', normalize_dry_formula, text, flags=re.S)
                if not count:
                    root = re.search(r'<((?:\w+:)?worksheet)\b[^>]*>', text)
                    if not root:
                        raise ValueError(f'Не найден корень листа: {entry.filename}')
                    prefix = root.group(1).removesuffix('worksheet')
                    views = f'<{prefix}sheetViews><{prefix}sheetView workbookViewId="0" zoomScale="80" zoomScaleNormal="80"/></{prefix}sheetViews>'
                    anchor = re.search(r'<(?:\w+:)?(?:sheetFormatPr|cols|sheetData)\b', text)
                    position = anchor.start() if anchor else root.end()
                    text = text[:position] + views + text[position:]
                data = text.encode('utf-8')
            elif re.fullmatch(r'xl/tables/[^/]+\.xml', entry.filename):
                root = ET.fromstring(data)
                style = root.find(f'{{{namespace}}}tableStyleInfo')
                if style is not None:
                    style.attrib.pop('name', None)
                    style.set('showRowStripes', '0')
                    style.set('showColumnStripes', '0')
                ET.register_namespace('', namespace)
                if root.get('name') == 'DryRationEvents':
                    ET.register_namespace('', namespace)
                    for column in root.findall(f'{{{namespace}}}tableColumns/{{{namespace}}}tableColumn'):
                        if column.get('name') == 'Припинення':
                            formula = column.find(f'{{{namespace}}}calculatedColumnFormula')
                            if not has_dry_formulas:
                                if formula is not None:
                                    column.remove(formula)
                                continue
                            if formula is None:
                                formula = ET.SubElement(column, f'{{{namespace}}}calculatedColumnFormula')
                            formula.text = formula_text
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            target.writestr(entry, data)
    temporary.replace(path)


if __name__ == '__main__':
    set_zoom(sys.argv[1])