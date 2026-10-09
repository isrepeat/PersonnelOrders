import copy
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement

root = Path(__file__).resolve().parents[1]
task = root / 'Tasks/01.Food'
source = next((task / 'Results/2026.10.09/Extracts').glob('*МІЗЯК*.docx'))
out = task / 'Results/2026.10.09/SignatureTests'
out.mkdir(exist_ok=True)

def bind_signature(doc):
    # Связываем абзацы всех строк таблицы, кроме последнего абзаца последней строки.
    table = doc.tables[-1]
    for index, row in enumerate(table.rows):
        properties = row._tr.get_or_add_trPr()
        if properties.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cantSplit') is None:
            properties.append(OxmlElement('w:cantSplit'))
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.keep_with_next = index < len(table.rows) - 1
                paragraph.paragraph_format.keep_together = True

for number in [1, 2, 3]:
    doc = Document(source)
    basis = next(i for i, paragraph in enumerate(doc.paragraphs) if paragraph.text.startswith('Підстава:'))
    bind_signature(doc)
    first = basis
    if number == 2:
        # Два последних пункта списка вместе с промежуточным пустым абзацем.
        first = basis - 3
    for paragraph in doc.paragraphs[first:]:
        if paragraph._p.getprevious() is not None and paragraph._p.getprevious().tag.endswith('}tbl'):
            break
        paragraph.paragraph_format.keep_with_next = True
        paragraph.paragraph_format.keep_together = True
    if number == 3:
        doc.paragraphs[basis].paragraph_format.page_break_before = True
    names = {
        1: '01 Основание и подписи вместе.docx',
        2: '02 Два пункта списка, основание и подписи вместе.docx',
        3: '03 Новая страница перед основанием.docx',
    }
    doc.save(out / names[number])

# Проверяем, что геометрия страниц и свойства шрифтов остались исходными.
from lxml import etree as ET
original = Document(source)
for path in out.glob('*.docx'):
    doc = Document(path)
    assert ET.tostring(doc.sections[0]._sectPr) == ET.tostring(original.sections[0]._sectPr)
    for before, after in zip(original.paragraphs, doc.paragraphs):
        assert before.text == after.text
        assert before.paragraph_format.first_line_indent == after.paragraph_format.first_line_indent
        assert before.paragraph_format.line_spacing == after.paragraph_format.line_spacing
        assert [(run.text, run.font.size, run.font.name) for run in before.runs] == [(run.text, run.font.size, run.font.name) for run in after.runs]
print('Three DOCX variants created; text, fonts, margins and spacing preserved')