from docx import Document


class DocxWriter:
    def __init__(self):
        self.document = Document()

    def add_paragraph(self, text):
        self.document.add_paragraph(text)

    def add_heading(self, text, level):
        self.document.add_heading(text, level)

    def add_table(self, table):
        new_table = self.document.add_table(rows=len(table), cols=len(table[0]))
        for i in range(0, len(table)):
            for j in range(0, len(table[0])):
                new_table.rows[i].cells[j].text = table[i][j]

    def save(self, path):
        self.document.save(path)
