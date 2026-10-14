"""Явно задаёт украинский язык и включает проверку текста Word без изменения компоновки."""

from lxml import etree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def ukrainian_tree(tree):
    # Язык конца абзаца также задаём явно, иначе пустые строки наследуют язык шаблона.
    for paragraph in tree.iter(W + 'p'):
        properties = paragraph.find(W + 'pPr')
        if properties is None:
            properties = ET.Element(W + 'pPr')
            paragraph.insert(0, properties)
        if properties.find(W + 'rPr') is None:
            ET.SubElement(properties, W + 'rPr')
    for run in tree.iter(W + 'r'):
        properties = run.find(W + 'rPr')
        if properties is None:
            properties = ET.Element(W + 'rPr')
            run.insert(0, properties)
    for properties in tree.iter(W + 'rPr'):
        for excluded in list(properties.findall(W + 'noProof')):
            properties.remove(excluded)
        language = properties.find(W + 'lang')
        if language is None:
            language = ET.SubElement(properties, W + 'lang')
        for attribute in ('val', 'eastAsia', 'bidi'):
            language.set(W + attribute, 'uk-UA')
    if tree.tag == W + 'settings':
        for tag in ('hideSpellingErrors', 'hideGrammaticalErrors'):
            for node in list(tree.findall(W + tag)):
                tree.remove(node)
        state = tree.find(W + 'proofState')
        if state is not None:
            state.set(W + 'spelling', 'dirty')
            state.set(W + 'grammar', 'dirty')
        theme_language = tree.find(W + 'themeFontLang')
        if theme_language is not None:
            theme_language.set(W + 'val', 'uk-UA')
    return tree


def ukrainian_part(name, data):
    if not (name.startswith('word/') and name.endswith('.xml')):
        return data
    tree = ET.fromstring(data)
    before = ET.tostring(tree)
    ukrainian_tree(tree)
    if ET.tostring(tree) == before:
        return data
    return ET.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)