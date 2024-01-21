import docx
import lxml.etree as ET


class Node:
    def __init__(self, id, tag, name, parent=None, value=None):
        self.id = id
        self.tag = tag
        self.name = name
        self.parent = parent
        self.value = value
        self.children = []


def parse_content_description_to_tree(doc):

    tocs = (styles['toc 1'], styles['toc 2'], styles['toc 3'],)


    root = Node(0, 'main', 'ROOT')
    first_level_id = 1

    in_content_description = False
    for i in range(len(doc.paragraphs)):
        para = doc.paragraphs[i]

        if in_content_description:
            s_name = para.style.name

            if s_name not in tocs:
                return root

            if s_name == styles['toc 1']:
                node = Node(first_level_id, 'sec', ' '.join(para.text.split()[:-1]), root)
                root.children.append(node)
                first_level_id += 1
            else:
                id = para.text.split()[0].split('.')
                parent = root
                for i in range(len(id) - 1):
                    for ch in parent.children:
                        if ch.id == id[i]:
                            parent = ch
                            break
                node = Node(id[-1], 'subsec', ' '.join(para.text.split()[:-1]), parent)
                parent.children.append(node)

        if para.style.name == styles['content_description']:
            in_content_description = True


def parse_content_description_to_etree(doc):

    tocs = (styles['toc 1'], styles['toc 2'], styles['toc 3'],)


    root = ET.Element('content_description')
    last_first_level = None
    last_second_level = None

    in_content_description = False
    for i in range(len(doc.paragraphs)):
        para = doc.paragraphs[i]

        if in_content_description:
            s_name = para.style.name

            if s_name not in tocs:
                return root

            if s_name == styles['toc 1']:
                node = ET.SubElement(root, 'sec')
                name = ' '.join(para.text.split()[:-1])
                if name[0].isdigit():
                    name = ' '.join(name.split()[1:])
                node.set('name', name)

                last_first_level = node

            if s_name == styles['toc 2']:
                node = ET.SubElement(last_first_level, 'subsec1')
                name = ' '.join(para.text.split()[:-1])
                if name[0].isdigit():
                    name = ' '.join(name.split()[1:])
                node.set('name', name)

                last_second_level = node

            if s_name == styles['toc 3']:
                node = ET.SubElement(last_second_level, 'subsec2')
                name = ' '.join(para.text.split()[:-1])
                if name[0].isdigit():
                    name = ' '.join(name.split()[1:])
                node.set('name', name)

        if para.style.name == styles['content_description']:
            in_content_description = True
    return root


def parse_em_all(doc):
    root = parse_content_description_to_etree(doc)
    prev = root

    tocs = (styles['toc 1'], styles['toc 2'], styles['toc 3'],)

    levels = {
        'sec': 1,
        'subsec1': 2,
        'subsec2': 3,
    }

    doc_index = 0

    for elem in root.iter():
        if elem == root:
            continue

        if prev == root or levels[prev.tag] < levels[elem.tag]:
            prev = elem
            continue

        print(elem.get('name'))
        doc_index = 0

        while doc_index < len(doc.paragraphs):
            if prev.get('name') in doc.paragraphs[doc_index].text and doc.paragraphs[doc_index].style.name not in tocs:
                s_name = doc.paragraphs[doc_index].style.name
                doc_index += 1
                for i in range(doc_index, len(doc.paragraphs)):
                    para = doc.paragraphs[i]

                    if para.style.name == s_name:
                        doc_index = i
                        break

                    node = ET.SubElement(prev, 'p')
                    node.text = para.text
                break
            doc_index += 1
        prev = elem
    doc_index = 0
    while doc_index < len(doc.paragraphs):
        if prev.get('name') in doc.paragraphs[doc_index].text and doc.paragraphs[doc_index].style.name not in tocs:
            s_name = doc.paragraphs[doc_index].style.name
            doc_index += 1
            for i in range(doc_index, len(doc.paragraphs)):
                para = doc.paragraphs[i]

                if para.style.name == s_name:
                    doc_index = i
                    break

                node = ET.SubElement(prev, 'p')
                node.text = para.text
            break
        doc_index += 1
    return root


styles = {
    'content_description': 'Название-caps',  # Содержание
    'toc 1': 'toc 1',
    'toc 2': 'toc 2',
    'toc 3': 'toc 3',  # Перечисления в содержании (индексация длины 1, 2 и 3 соответственно)
    'subsec v': 'Нумерованный заголовок 3'
}

doc = docx.Document("input.docx")
root = parse_em_all(doc)

et = ET.ElementTree(root)
et.write('output.xml', encoding="utf-8", pretty_print=True)
