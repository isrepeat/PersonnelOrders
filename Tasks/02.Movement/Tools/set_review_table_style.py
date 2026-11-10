"""Именованный кофейный стиль таблиц и акцент ПІБ при добавлении строк."""
import posixpath
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


def set_table_style(path):
    path=Path(path)
    ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel_ns='http://schemas.openxmlformats.org/package/2006/relationships'
    office_ns='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    tag=lambda name:f'{{{ns}}}{name}'
    ET.register_namespace('',ns)
    with ZipFile(path) as archive:
        entries=[(entry,archive.read(entry.filename)) for entry in archive.infolist()]
    contents={entry.filename:data for entry,data in entries}
    styles=ET.fromstring(contents['xl/styles.xml'])
    dxfs=styles.find(tag('dxfs'))
    if dxfs is None:
        dxfs=ET.Element(tag('dxfs'),count='0')
        anchor=styles.find(tag('tableStyles'))
        styles.insert(list(styles).index(anchor) if anchor is not None else len(styles),dxfs)

    def differential(color,bold=False):
        index=len(dxfs)
        dxf=ET.SubElement(dxfs,tag('dxf'))
        font=ET.SubElement(dxf,tag('font'))
        ET.SubElement(font,tag('name'),val='Arial')
        ET.SubElement(font,tag('sz'),val='11')
        ET.SubElement(font,tag('color'),rgb='FFFFFFFF')
        if bold:ET.SubElement(font,tag('b'))
        pattern=ET.SubElement(ET.SubElement(dxf,tag('fill')),tag('patternFill'),patternType='solid')
        ET.SubElement(pattern,tag('fgColor'),rgb=color)
        ET.SubElement(pattern,tag('bgColor'),indexed='64')
        border=ET.SubElement(dxf,tag('border'))
        for side in ['left','right','top','bottom']:
            ET.SubElement(ET.SubElement(border,tag(side),style='thin'),tag('color'),rgb='FF000000')
        return index

    whole,header,stripe=differential('FF302D25'),differential('FF5D5142',True),differential('FF38342C')
    gold=len(dxfs)
    ET.SubElement(ET.SubElement(ET.SubElement(dxfs,tag('dxf')),tag('font')),tag('color'),rgb='FFF0B56A')
    dxfs.set('count',str(len(dxfs)))
    table_styles=styles.find(tag('tableStyles'))
    if table_styles is None:table_styles=ET.SubElement(styles,tag('tableStyles'))
    style=ET.SubElement(table_styles,tag('tableStyle'),name='CoffeeMovement',pivot='0',table='1',count='4')
    for element,index in [('wholeTable',whole),('headerRow',header),('firstRowStripe',whole),('secondRowStripe',stripe)]:
        ET.SubElement(style,tag('tableStyleElement'),type=element,dxfId=str(index),**({'size':'1'} if 'Stripe' in element else {}))
    table_styles.set('count',str(len(table_styles)))
    contents['xl/styles.xml']=ET.tostring(styles,encoding='utf-8',xml_declaration=True)
    for name,data in list(contents.items()):
        if not re.fullmatch(r'xl/tables/[^/]+\.xml',name):continue
        table=ET.fromstring(data)
        info=table.find(tag('tableStyleInfo'))
        if info is None:info=ET.SubElement(table,tag('tableStyleInfo'))
        info.set('name','CoffeeMovement')
        info.set('showRowStripes','1')
        info.set('showColumnStripes','0')
        contents[name]=ET.tostring(table,encoding='utf-8',xml_declaration=True)
    # Произвольный столбец ПІБ поддерживается условным форматированием до конца листа.
    for name,data in list(contents.items()):
        if not re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name):continue
        rel_path=posixpath.join(posixpath.dirname(name),'_rels',posixpath.basename(name)+'.rels')
        if rel_path not in contents:continue
        rels={r.get('Id'):r.get('Target') for r in ET.fromstring(contents[rel_path]).findall(f'{{{rel_ns}}}Relationship')}
        sheet=ET.fromstring(data)
        for part in sheet.findall(f'{tag("tableParts")}/{tag("tablePart")}'):
            target=rels[part.get(f'{{{office_ns}}}id')]
            table_path=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(name),target))
            table=ET.fromstring(contents[table_path])
            headers=[c.get('name') for c in table.findall(f'{tag("tableColumns")}/{tag("tableColumn")}')]
            fio=next((i for i,h in enumerate(headers) if h=='ПІБ' or h=='ПІБ / ідентифікатор'),None)
            if fio is None:continue
            start=re.match(r'([A-Z]+)(\d+)',table.get('ref'))
            first=0
            for char in start[1]:first=first*26+ord(char)-64
            column=first+fio
            letters=''
            while column:column,rem=divmod(column-1,26);letters=chr(65+rem)+letters
            row=int(start[2])+1
            cf=ET.Element(tag('conditionalFormatting'),sqref=f'{letters}{row}:{letters}1048576')
            priority=1+max([int(r.get('priority','0')) for r in sheet.iter(tag('cfRule'))] or [0])
            rule=ET.SubElement(cf,tag('cfRule'),type='expression',dxfId=str(gold),priority=str(priority))
            ET.SubElement(rule,tag('formula')).text=f'ROW()>={row}'
            following={'dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'}
            anchor=next((n for n in sheet if n.tag.rsplit('}',1)[-1] in following),None)
            sheet.insert(list(sheet).index(anchor) if anchor is not None else len(sheet),cf)
        contents[name]=ET.tostring(sheet,encoding='utf-8',xml_declaration=True)
    temporary=path.with_suffix('.coffee.tmp')
    with ZipFile(temporary,'w') as archive:
        for entry,data in entries:archive.writestr(entry,contents[entry.filename])
    temporary.replace(path)