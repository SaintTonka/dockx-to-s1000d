import docx


class Node:
    def __init__(self, id, name, parent=None, value=None):
        self.id = id
        self.name = name
        self.parent = parent
        self.value = value


styles = {
    'content_description': 'Название-caps',  # Содержание
}

doc = docx.Document("input.docx")
