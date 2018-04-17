from docx import Document


class DocxWriter:
    def __init__(self):
        self.document = Document('test_1.docx')

    def add_paragraph(self, text):
        self.document.add_paragraph(text)

    def add_heading(self, text, level):
        self.document.add_heading(text, level)

    def add_table(self, table):
        new_table = self.document.add_table(rows=len(table), cols=len(table[0]))
        for i in range(0, len(table)):
            for j in range(0, len(table[0])):
                new_table.rows[i].cells[j].text = table[i][j]

    def add_bullet_element(self, string):
        p = self.document.add_paragraph()
        p.style = 'List Bullet'
        p.add_run(string)

    def save(self, path):
        self.document.save(path)
