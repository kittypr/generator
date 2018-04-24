import json
from subprocess import Popen, PIPE


class JsonDocParser:

    def __init__(self, doc):
        self.doc = doc
        self.immediate_writing = True
        self.data = ''

        # ****** INDICATORS ******

        self.current_header_level = 0
        self.bullet_level = 0
        self.fmt = {'Emph': 0, 'Strong': 0, 'Strikeout': 0}

    # ***** INDICATOR`s METHODS *****

    def collect_data(self):
        self.immediate_writing = False

    def new_data(self):
        self.data = ''

    def add_data(self, new_data):
        self.data += new_data

    # ***** WRITING METHODS *****

    def write_data(self):
        if self.bullet_level > 0:
            self.doc.add_bullet_element(self.data, self.bullet_level)
        elif self.current_header_level != 0:
            self.doc.add_heading(self.data, self.current_header_level)
        else:
            self.doc.add_paragraph(self.data)
        self.new_data()

    def write_table(self, input_table):
        if self.immediate_writing and self.data:
            self.write_data()
        previous_state = self.immediate_writing
        self.collect_data()
        table = list()
        headers = input_table['c'][3]
        is_headers = False
        if headers:
            is_headers = True
            header_row = list()
            for column in headers:
                self.list_parse(column)
                cell = self.data
                self.new_data()
                header_row.append(cell)
            table.append(header_row)
        table_rows = input_table['c'][4]
        for table_row in table_rows:
            row = list()
            for column in table_row:
                self.list_parse(column)
                cell = self.data
                self.new_data()
                row.append(cell)
            table.append(row)
        self.immediate_writing = previous_state
        self.doc.add_table(table, is_headers)

    # ***** PARSING METHODS *****

    def write_special_block(self, block):
        if self.immediate_writing and self.data:
            self.write_data()
        con = 1  # we chose in what part of block content is
        if block['t'] == 'Header':
            self.current_header_level = block['c'][0]
            con = 2
        previous_state = self.immediate_writing
        self.collect_data()
        self.list_parse(block['c'][con])
        self.immediate_writing = previous_state
        if self.immediate_writing:
            self.write_data()
        self.current_header_level = 0

    def write_bullet(self, bull_list):
        self.bullet_level = self.bullet_level + 1
        self.list_parse(bull_list['c'])
        self.bullet_level = self.bullet_level - 1

    def dict_parse(self, dictionary):
        try:
            if dictionary['t'] in self.fmt.keys():
                self.fmt[dictionary['t']] = 1
            if dictionary['t'] == 'Table':
                self.write_table(dictionary)
            elif dictionary['t'] == 'CodeBlock' or dictionary['t'] == 'Code':
                pass
            elif dictionary['t'] == 'Div' or dictionary['t'] == 'Span' or dictionary['t'] == 'Header':
                self.write_special_block(dictionary)
            elif dictionary['t'] == 'Math':
                pass
            elif dictionary['t'] == 'Link':
                pass
            elif dictionary['t'] == 'BulletList':
                self.write_bullet(dictionary)
            elif dictionary['t'] == 'OrderedList':
                pass
            elif dictionary['t'] == 'Image':
                pass
            elif 'c' in dictionary:
                if type(dictionary['c']) == str:
                    self.add_data(dictionary['c'])
                if type(dictionary['c']) == list:
                    self.list_parse(dictionary['c'])
            else:
                if dictionary['t'] == 'Space':
                    self.add_data(' ')
                elif dictionary['t'] == 'SoftBreak':
                    self.add_data('\n')
                elif dictionary['t'] == 'LineBreak':
                    self.add_data('\n\n')
                    if self.immediate_writing:
                        self.write_data()
            if dictionary['t'] == 'Para' or dictionary['t'] == 'Plain':
                if self.immediate_writing:
                    self.write_data()
        except KeyError as err:
            print('Incompatible Pandoc version. Data could be excepted.')
            print(err)

    def list_parse(self, content_list):
        for item in content_list:
            if type(item) == dict:
                self.dict_parse(item)
            if type(item) == list:
                self.list_parse(item)


def get_json(filename):
    command = 'pandoc -f markdown -t json ' + '"' + filename + '"'
    print(command)
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    res = proc.communicate()

    if res[1]:
        print('ERROR', str(res[1].decode('cp866')))  # sending stderr output to user
        return None
    else:
        document_json = json.loads(res[0])
        return document_json


def main(filename, doc):
    document_json = get_json(filename)
    doc_parser = JsonDocParser(doc)
    if type(document_json) == dict:
        doc_parser.list_parse(document_json['blocks'])
    elif type(document_json) == list:
        doc_parser.list_parse(document_json[1])
    else:
        print('Incompatible Pandoc version')
