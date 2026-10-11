"""Создаёт шаблон зачисления, сохраняя оформление действующего шаблона снятия."""

import zipfile
from pathlib import Path

from lxml import etree as ET

from build_food_extracts import replace_placeholders, W

folder = Path(__file__).resolve().parents[1] / 'Documents/Templates/Extracts [food]'
source = next(folder.glob('*снять с продовольствия*.docx'))
target = folder / source.name.replace('снять с продовольствия', 'зачислить на продовольствие')
if target.exists():
    raise FileExistsError(target)
with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as result:
    for item in original.infolist():
        data = original.read(item.filename)
        if item.filename == 'word/document.xml':
            tree = ET.fromstring(data)
            for paragraph in tree.iter(W + 'p'):
                replace_placeholders(paragraph, {
                    'зняти з продовольчого забезпечення': 'зарахувати на продовольче забезпечення',
                    '{RemovalDate}': '{EnrollmentDate}',
                })
            data = ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        result.writestr(item, data)
print(target.name)