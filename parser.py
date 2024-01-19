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


    root = ET.Element('data')
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
                node.set('name', ' '.join(para.text.split()[:-1]))

                last_first_level = node

            if s_name == styles['toc 2']:
                node = ET.SubElement(last_first_level, 'subsec')
                node.set('name', ' '.join(para.text.split()[:-1]))

                last_second_level = node

            if s_name == styles['toc 3']:
                node = ET.SubElement(last_second_level, 'subsec')
                node.set('name', ' '.join(para.text.split()[:-1]))

        if para.style.name == styles['content_description']:
            in_content_description = True
    return root


def print_tree(root):
    print(root.name)
    for ch in root.children:
        print_tree(ch)


styles = {
    'content_description': 'Название-caps',  # Содержание
    'toc 1': 'toc 1',
    'toc 2': 'toc 2',
    'toc 3': 'toc 3',  # Перечисления в содержании (индексация длины 1, 2 и 3 соответственно)
}

doc = docx.Document("input.docx")
root = parse_content_description_to_etree(doc)

et = ET.ElementTree(root)
et.write('output.xml', encoding="utf-8", pretty_print=True)
