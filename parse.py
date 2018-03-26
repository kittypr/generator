import json
from subprocess import Popen, PIPE

class JsonDocParser:

    def __init__(self, gdoc):
        self.gdoc = gdoc
        self.immediate_writing = True
        self.data = ''
        self.fmt ={'Emph': 0, 'Strong': 0, 'Strikeout': 0}

    # ***** INDICATOR`s METHODS *****

    def collect_data(self):
        self.immediate_writing = False

    def new_data(self):
        self.data = ''

    def add_data(self, new_data):
        self.data += new_data

    # ***** WRITING METHODS *****

    def write_data(self):
        self.gdoc.write(self.data)
        self.new_data()

    # ***** PARSING METHODS *****

    def dict_parse(self, dictionary):
        try:
            if dictionary['t'] in self.fmt.keys():
                self.fmt[dictionary['t']] = 1
            if dictionary['t'] == 'Table':
                pass
            elif dictionary['t'] == 'CodeBlock' or dictionary['t'] == 'Code':
                pass
            elif dictionary['t'] == 'Div' or dictionary['t'] == 'Span' or dictionary['t'] == 'Header':
                pass
            elif dictionary['t'] == 'Math':
                pass
            elif dictionary['t'] == 'Link':
                pass
            elif dictionary['t'] == 'BulletList':
                pass
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
            if dictionary['t'] == 'Para':
                self.add_data('\n')
                if self.immediate_writing:
                    self.write_data()
        except KeyError:
            print('Incompatible Pandoc version. Data could be excepted.')

    def list_parse(self, content_list):
        for item in content_list:
            if type(item) == dict:
                self.dict_parse(item)
            if type(item) == list:
                self.list_parse(item)


def get_json(filename):
    command = 'pandoc -f markdown -t json ' + filename
    print(command)
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    res = proc.communicate()

    if res[1]:
        print(str(res[1].decode('cp866')))  # sending stderr output to user
        return None
    else:
        document_json = json.loads(res[0])
        return document_json


def main(filename, gdoc):
    document_json = get_json(filename)
    doc_parser = JsonDocParser(gdoc)
    if type(document_json) == dict:
        doc_parser.list_parse(document_json['blocks'])
    elif type(document_json) == list:
        doc_parser.list_parse(document_json[1])
    else:
        print('Incompatible Pandoc version')