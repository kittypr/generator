from docx import Document


class DocxWriter:
    def __init__(self):
        self.document = Document('style.docx')

    def add_paragraph(self, text):
        p = self.document.add_paragraph(text)
        p.style = 'Normal'

    def add_heading(self, text, level):
        h = self.document.add_heading(text, level)
        h.style = 'Heading ' + str(level)

    def add_table(self, table):
        new_table = self.document.add_table(rows=len(table), cols=len(table[0]))
        new_table.style = 'Table'
        for i in range(0, len(table)):
            for j in range(0, len(table[0])):
                new_table.rows[i].cells[j].text = table[i][j]

    def add_bullet_element(self, string, level):
        p = self.document.add_paragraph()
        p.style = 'List Bullet'
        if level in range(1, 5):
            p.style = 'List Bullet ' + str(level)
        p.add_run(string)

    def save(self, path):
        self.document.save(path)
