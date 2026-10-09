import sys
from pathlib import Path
from openpyxl import load_workbook
from pypdf import PdfReader
from pdf2image import convert_from_path
from PIL import Image, ImageOps, ImageDraw

sys.stdout.reconfigure(encoding='utf-8')
root = Path(__file__).resolve().parents[1] / 'Tasks/01.Food'
work = root / 'Work/2026.10.09'
output_path = root / 'Results/2026.10.09/v2 Рапорти_продовольче_сухпрод_2026.10.09.xlsx'
wb = load_workbook(output_path)
assert len(wb['Сухпрод'].tables['DryRations'].ref.split(':')) == 2
assert wb['Сухпрод'].tables['DryRations'].ref == 'B2:N101'
assert wb['Продовольче'].tables['FoodReports'].ref == 'B2:O3'
assert not any(cell.value is not None for row in wb['Продовольче'].iter_rows(min_row=3) for cell in row)
assert [wb['Сухпрод'].cell(2, col).value for col in range(2, 11)] == ['Звання', 'ПІБ', 'ІПН/Установа', 'Початок', 'Тривалість', 'Припинення', 'Підстава', 'Наказ', 'Додаткова інформація']
assert wb['Сухпрод']['C3'].fill.fgColor.rgb != 'FFFFC7CE'
assert any(rule.type == 'duplicateValues' for rules in wb['Сухпрод'].conditional_formatting._cf_rules.values() for rule in rules)
assert wb['Сухпрод']['B3'].font.name == 'Times New Roman'
reference = load_workbook(work / 'РУХ_last_format_reference.xlsx')
assert all(wb['Сухпрод'].column_dimensions[col].width == reference['Сухпрод'].column_dimensions[col].width for col in 'ABCDEFGHIJ')
counts = {}
for row in range(3, 102):
    reporter = wb['Сухпрод'].cell(row, 13).value
    counts[reporter] = counts.get(reporter, 0) + 1
assert counts == {'КОРНЕНКО': 76, 'МОРОЗОВИЧ': 3, 'МІЗЯК': 20}
assert sum(wb['Сухпрод'].cell(row, 6).value for row in range(3, 102)) == 221
for sheet in wb:
    for row in sheet.iter_rows():
        assert all('.docx' not in str(cell.value).lower() for cell in row)
for sheet in wb:
    assert sheet.freeze_panes is None
    assert wb._fills[wb._cell_styles[sheet.column_dimensions['A'].style].fillId].fgColor.rgb[-6:] == '262626'
    assert sheet.column_dimensions['AO'].style or any(c.min <= 16384 <= c.max and c.style for c in sheet.column_dimensions.values())
    print(sheet.title, sheet.max_row, 'no freeze, dark canvas')
for row in range(3, 102):
    assert wb['Сухпрод'][f'G{row}'].value in ['=[@Початок]+[@Тривалість]', '=DryRations[[#This Row],[Початок]]+DryRations[[#This Row],[Тривалість]]']
    assert wb['Сухпрод'][f'E{row}'].value.year == 2026
    assert wb['Сухпрод'][f'G{row}'].number_format == reference['Сухпрод']['G1380'].number_format
cached = load_workbook(output_path, data_only=True)
assert cached['Сухпрод']['G3'].value.day == 14
poppler = 'C:/Users/isrepeat/.cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin'
thumbs = []
for file in work.glob('*.pdf'):
    pages = convert_from_path(str(file), dpi=100, poppler_path=poppler)
    print(file.name, len(pages), 'pages')
    for index, page in enumerate(pages, 1):
        page.save(work / f'{file.stem}-{index}.png')
        thumb = ImageOps.contain(page, (420, 600))
        tile = Image.new('RGB', (440, 630), 'white')
        tile.paste(thumb, (10, 25))
        ImageDraw.Draw(tile).text((5, 5), f'{file.stem} / {index}', fill='black')
        thumbs.append(tile)
contact = Image.new('RGB', (440 * 4, 630 * ((len(thumbs) + 3) // 4)), '#888888')
for index, tile in enumerate(thumbs):
    contact.paste(tile, ((index % 4) * 440, (index // 4) * 630))
contact.save(work / 'extracts-contact.png')