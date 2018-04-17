#!/usr/bin/env python3
import argparse
import json
import sys
from subprocess import Popen, PIPE


from odf.draw import Image, Frame
from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style, TableColumnProperties, TableRowProperties, TextProperties
from odf.table import Table, TableRow, TableCell, TableColumn
from odf.text import P, A

from pandocodswriter.limages import load_images
from pandocodswriter.lstyle import load_style, add_fmt, st_dict
# usage - python odswriter.py yourInputFile.yourExetention yourOutputFile.ods -s *YOUR POSITIVE NUMBER*
# check README.md for more information.
# DO NOT mix up places of intput and output.

# Style names.
# header0 - just for correct index.
# If in input file more, than two levels of headers, next level header will generate automatically
#  with name = "header" + str(level).
header = ['header0', 'header1', 'header2']
table_header = 'tablehead'
table_content = 'tablebody'
simple_text = 'text'


# Read the command-line arguments.
parser = argparse.ArgumentParser(description='Pandoc ODS writer. This is Pandoc filter, but there is no opportunity '
                                             'write .ods files easier way. So, use "out.ods" '
                                             'option to write .ods files with this filter')
parser.add_argument('input', help='Input file. Use Pandoc`s input formats.', action='store')
parser.add_argument('output', help='Output file. Use .ods filename extension.', action='store')
parser.add_argument('-s', '--separator', nargs=1, help='Header level to separate sheets, 0 by default(no separation).',
                    action='store')
parser.add_argument('-r', '--reference', nargs=1, help='Reference to file with styles', action='store')
args = parser.parse_args()

# It is important for auto-height in text-rows:
# if you want to change width by default (10 cm), change it in 'write_sheet()',
# count how much PT in your length (in CM) and change this constant:
PTINTENCM = 284
# count how much IN in your length (in CM) and change this constant:
ININTENCM = 3.9

# I need this global variables, because there are two recursive functions call each other, so it would be very hard work
# without global "string_to_write". Other ones are just make those functions much more easy to read.

ods = OpenDocumentSpreadsheet()
table = Table()  # creating the first sheet

content = P()
string_to_write = ''
header_level = 0
bullet = 0  # indicating bullet lists
ordered = 0  # indicating bullet list and used as order at item lines
image_counter = 0
saved_hr = None  # list of hardreferences to loaded images
saved_styles = {}  # We will save styles in order to not downloading it again each time we use it.
separator = 0  # level of separating header

# Dictionary of formatting indicators.
fmt = {'Emph': 0,
       'Strong': 0,
       'Strikeout': 0}


def write_sheet():
    wide = Style(name="Wide", family="table-column")
    wide.addElement(TableColumnProperties(columnwidth="10cm"))
    ods.automaticstyles.addElement(wide)
    table.addElement(TableColumn(stylename='Wide'))
    ods.spreadsheet.addElement(table)


def count_height(row, cell):
    """Counting height that shows all text in cell.
    This functions uses width of text-column and font size.
    Args:
        row - current row.
        cell - current cell.
    """
    style_name = cell.getAttribute('stylename')
    try:
        style = saved_styles[style_name]
        text_prop = style.getElementsByType(TextProperties)
        try:
            text_prop = text_prop[0]
            font_size = str(text_prop.getAttribute('fontsize'))
            font_size = font_size.replace('pt', '')
            font_size = int(font_size)
        except IndexError:
            font_size = 10
    except KeyError:
        font_size = 10

    symbols_in_string = PTINTENCM // font_size + 1
    length = 0
    for p in cell.getElementsByType(P):
        length += len(p.__str__())
    height = font_size*(length // symbols_in_string + 1) + 4
    height = str(height) + 'pt'
    new_name = 'heightsuit' + height
    height_suit = Style(name=new_name, family='table-row')
    height_suit.addElement(TableRowProperties(rowheight=height))

    ods.automaticstyles.addElement(height_suit)
    row.setAttribute(attr='stylename', value=new_name)


def count_size(wh_list, row):
    """Count height of image row.
    Args:
        wh_list - list with attributes, contains width and height:
        row - image row.
    """
    height, width = -1, -1
    for l in wh_list:
        if l[0] == 'width':
            width = float(l[1].replace('in', ''))
        if l[0] == 'height':
            height = float(l[1].replace('in', ''))
    if height == -1 or width == -1:
        width = ININTENCM
        height = ININTENCM
    if width > ININTENCM:
        new_width = ININTENCM
        new_height = height * new_width / width
    else:
        new_width = width
        new_height = height
    height_set = str(new_height)+'in'
    new_name = 'image' + str(image_counter)
    height_suit = Style(name=new_name, family='table-row')
    height_suit.addElement(TableRowProperties(rowheight=height_set))

    ods.automaticstyles.addElement(height_suit)
    row.setAttribute(attr='stylename', value=new_name)

    new_width = str(new_width) + 'in'
    new_height = height_set
    return new_width, new_height


def add_style(cell, name):
    """Add style to cell element.
    This function calls style loading from 'lstyle' and saves it, in order to use it again quickly.
    Args:
        cell - cell that needs to be styled.
        name - style name that will be set.
    """
    if args.reference:
        styles_source = args.reference[0]
    else:
        styles_source = str(sys.argv[0])
        styles_source = styles_source.replace('odswriter.py', '')
        styles_source = styles_source + 'styles.ods'
    global saved_styles
    global ods
    try:
        saved_styles[name]
    except KeyError:
        style = load_style(name, styles_source)
        if style is not None:
            saved_styles[name] = style
            ods.styles.addElement(style)
    cell.setAttribute(attr='stylename', value=name)


def write_text():
    """Write to output file ordinary elements.
    This function is called every tame, we collect whole paragraph or block of elements in 'string_to_write'
    We write every block or paragraph in it's own cell in the first column of output file.
    After writing we shift down current row and clean 'string_to_write' in order to collect next elements.
    """
    global string_to_write
    global header_level
    global ordered
    global bullet
    global table
    global separator
    global content
    row = TableRow()
    cell = TableCell()
    if header_level != 0 and header_level > 0:
        if header_level > (len(header) - 1):  # if there are headers with lvl bigger than 2
            for i in range(len(header), header_level+1):  # creating names for headers with lvl bigger than 2
                header.append('header' + str(i))
        add_style(cell, header[header_level])
        if header_level == separator:  # if separator was set, we will create new sheet in document
            if table.hasChildNodes():
                write_sheet()
            table = Table(name=string_to_write)  # creating new sheet with separating header as name
    else:
        add_style(cell, simple_text)
    if bullet:
        string_to_write = '- ' + string_to_write
    if ordered > 0:
        string_to_write = str(ordered) + ') ' + string_to_write
        ordered = ordered + 1
    content.addText(string_to_write)
    cell.addElement(content)
    content = P()
    count_height(row, cell)
    row.addElement(cell)
    table.addElement(row)
    string_to_write = ''


def write_image(image):
    """Write to output file image elements.
    Since, element with title 'Image' has special structure of 'c'(Content) field, that looks like:
    [[0], [1], [2]]
    where:
        [0] - list of attributes: identifier, classes, key-value pairs:
            ['id', [], [ ... , ['weight', '...in'], ['height', '...in'], ... ] - we get sizes there.
        [1] - caption.
        [2] - ['src', 'title'] - source and title of image.
    we should parse it especially.
    Args:
        image - element with title 'Image'.
    """
    global image_counter
    global saved_hr
    if image_counter == -1:
        return
    if image_counter == 0:
        saved_hr = load_images(args.input, ods)
        if len(saved_hr) == 0:
            image_counter = -1
            return
    if string_to_write:
        write_text()
    row = TableRow()
    cell = TableCell()
    w, h = count_size(image['c'][0][2], row)
    frame = Frame(width=w, height=h)
    img = Image(href=saved_hr[image_counter])

    table.addElement(row)
    row.addElement(cell)
    cell.addElement(frame)
    frame.addElement(img)

    image_counter = image_counter + 1


def write_bullet(bull_list, without_write):
    global bullet
    bullet = 1
    list_parse(bull_list['c'], without_write)
    bullet = 0


def write_ord(ord_list, without_write):
    global ordered
    ordered = 1
    list_parse(ord_list['c'], without_write)
    ordered = 0


def write_code(code):
    """Write to output file code elements.
    Since, element with title 'Code' or 'CodeBlock' has special structure of 'c'(Content) field, that looks like:
    [[0], 'code']
    where:
        [0] - list of attributes: identifier, classes, key-value pairs.
        'code' - string with code.
    we should parse it especially.
    Args:
        code - element with title 'Code' or 'CodeBlock'.
    """
    global string_to_write
    string_to_write = string_to_write + code['c'][1]


def write_link(link):
    """Write special blocks with attributes.
    Since, element with title 'Link' has special structure of 'c'(Content) field, that looks like:
    [[atr], [1},['target', 'title']]
        where:
        [atr] - list of attributes: identifier, classes, key-value pairs.
        [1] - list with objects (list of dictionaries) - visible text of hyperlink.
        ['target', 'title'] - list with two strings, 'target' - URL, 'title' - title.
    we should parse it especially.
    Args:
        link - element with title 'Link'.
    """
    global string_to_write
    global content
    content.addText(string_to_write)
    string_to_write = ''
    list_parse(link['c'][1], without_write=True)
    a = A(href=link['c'][2][0], text=string_to_write)
    string_to_write = ''
    content.addElement(a)


def write_math(math):
    """Write to output file code elements
    Since, element with title 'Math' has special structure of 'c'(Content) field, that looks like:
    [{0}, 'math'].
    where:
        {0} - dictionary contains type of math.
        'math' - string with math.
    we should parse it especially.
    TeX Math format.
    Args:
        raw - element with title 'Math'.
    """
    # TODO: write it
    global string_to_write
    string_to_write = string_to_write + math['c'][1]


def write_raw(raw):
    """Write to output file raw elements.
    Since, element with title 'RawBlock' or 'RawInline' has special structure of 'c'(Content) field, that looks like:
    [format, 'raw text']
    where:
        format - format of raw text.
        'raw text' - string with raw text.
    we should parse it especially.
    Args:
        raw - element with title 'RawBlock' or 'RawInline'.
    """
    global string_to_write
    string_to_write = string_to_write + raw['c'][1]


def write_special_block(block, without_write):
    """Write special blocks with attributes.
    Since, element with title  'Div' or 'Span' or 'Header' has special structure of 'c'(Content) field, that looks like:
    [[0], [1]]*
    where:
    [0] - list of attributes: identifier, classes, key-value pairs.
    [1] - list with objects (list of dictionaries) - content.
    * with 'Header' title - [level, [0], [1]] - level - int, [0], [1] - the same as above.
    we should parse it especially.
    This function writes block itself.
    Args:
        block - element with title 'Div' or 'Span' or 'Header'.
        without_write - indicate calling write_text() functions. By default calls it.
    """
    global string_to_write
    global header_level
    con = 1
    if block['t'] == 'Header':
        header_level = block['c'][0]
        con = 2
    if (not without_write) and string_to_write:
        write_text()
    list_parse(block['c'][con], without_write=True)
    if not without_write:
        write_text()
    header_level = 0


def write_table(tab):
    """Write to output file table elements.
    This function is called every time, we meet 'Table' dictionary's title.
    Firstly, if we have some information in 'string_to_write' we record it, because we'll use this
    variable to collect information from table's cells.
    Table in pandoc's json has following structure:
    dict: { 't': 'Table'
            'c': [ [0] [1] [2] [3] [4] ]
          }
    Where:
    [0] - caption.
    [1] - is list of aligns by columns, looks like: [ { t: 'AlignDefault' }, ... ].
    [2] - widths of columns.
    [3] - is list of table's headers (top cell of every column), can be empty.
    [4] - list of rows, and row is list of cells.
    Since every cell's structure is the same as text's one, we just parse them as list and write one by one.
    Args:
        tab - dictionary with 't': 'Table".
    """
    global table
    global string_to_write
    global fmt

    for k in fmt.keys():  # setting to zero all outside-table formatting, we use formatting ONLY inside table-cell
        fmt[k] = 0

    if string_to_write:  # writing all text from buffer, table has it's own rules for rows, cols and their shift-rules
        write_text()

    row = TableRow()  # adding empty line before table
    table.addElement(row)

    row = TableRow()
    headers = tab['c'][3]
    if headers:
        cell = TableCell()  # adding empty first cell in row (first column in document - text column).
        row.addElement(cell)
        for col in headers:
            cell = TableCell()
            list_parse(col, without_write=True)
            add_style(cell, table_header)
            cell_content = P(text=string_to_write)
            for key in fmt.keys():
                if fmt[key] == 1:
                    new_style = add_fmt(style=st_dict[cell.getAttribute(attr='stylename')], key=key)
                    ods.styles.addElement(new_style)
                    fmt[key] = 0
                    cell = TableCell(stylename=new_style.getAttribute(attr='name'))
            cell.addElement(cell_content)
            string_to_write = ''
            row.addElement(cell)
        table.addElement(row)
    t_content = tab['c'][4]
    for line in t_content:
        row = TableRow()
        cell = TableCell()
        row.addElement(cell)
        for col in line:
            cell = TableCell()
            list_parse(col, without_write=True)
            add_style(cell, table_content)
            cell_content = P(text=string_to_write)
            for key in fmt.keys():
                if fmt[key] == 1:
                    new_style = add_fmt(style=st_dict[cell.getAttribute(attr='stylename')], key=key)
                    ods.styles.addElement(new_style)
                    fmt[key] = 0
                    cell = TableCell(stylename=new_style.getAttribute(attr='name'))
            cell.addElement(cell_content)
            string_to_write = ''
            row.addElement(cell)
        table.addElement(row)

    row = TableRow()  # adding empty line after table
    table.addElement(row)


# This two functions - 'dict_parse()' and 'list_parse()', has purpose of extract readable information
# from json object.
# Pandoc's json object in field 'blocks' has list of dictionaries or lists,
# which represent another objects and dictionaries has the following structure:
# { 't': '*Name of objects type*',
#   'c': '*Content of object*' }
# e.g.:
# { 't': 'Str',
#   'c': 'Hello' }
# (sometimes there can be no 'c'-field (e.g. 'Space' - object)
# So, 'c'-field - content, can be a list of dictionaries, or a string, or a list of lists,
# and we should parse list again and etc.
# That's why we have there two functions - for parsing lists and for parsing dictionaries - that should call each other

def dict_parse(dictionary, without_write=False):
    """Parse dictionaries.
    Dictionary represent some json-object. Kind of json object depends on 't' (title) field of it.
    We will parse it differently  depending on different titles. Sometimes this function write block,
    sometimes it leaves writing special functions.
    Args:
        dictionary - object with 't' and sometimes 'c' fields.
        without_write - indicate calling write_text() functions. By default calls it.
    """
    global string_to_write
    global fmt
    try:
        if dictionary['t'] in fmt.keys():
            fmt[dictionary['t']] = 1
        if dictionary['t'] == 'Table':
            write_table(dictionary)
        elif dictionary['t'] == 'CodeBlock' or dictionary['t'] == 'Code':
            write_code(dictionary)
        elif dictionary['t'] == 'Div' or dictionary['t'] == 'Span' or dictionary['t'] == 'Header':
            write_special_block(dictionary, without_write)
        elif dictionary['t'] == 'Math':
            write_math(dictionary)
        elif dictionary['t'] == 'Link':
            write_link(dictionary)
        elif dictionary['t'] == 'BulletList':
            write_bullet(dictionary, without_write)
        elif dictionary['t'] == 'OrderedList':
            write_ord(dictionary, without_write)
        elif dictionary['t'] == 'Image':
            write_image(dictionary)
        elif 'c' in dictionary:
            if type(dictionary['c']) == str:
                string_to_write = string_to_write + dictionary['c']
            if type(dictionary['c']) == list:
                list_parse(dictionary['c'], without_write)
        else:
            if dictionary['t'] == 'Space':
                string_to_write = string_to_write + ' '
            elif dictionary['t'] == 'SoftBreak':
                string_to_write = string_to_write + '\n'
            elif dictionary['t'] == 'LineBreak':
                string_to_write = string_to_write + '\n\n'
                if not without_write:
                    write_text()
            else:
                string_to_write = string_to_write
        if dictionary['t'] == 'Para':
            string_to_write = string_to_write + '\n'
            if not without_write:
                write_text()
    except KeyError:
        print('Incompatible Pandoc version')


def list_parse(content_list, without_write=False):
    """Parse list.
    Args:
        content_list - list with different parts of content from input-document.
        without_write - indicate calling write_text() functions. By default calls it.
    """
    for item in content_list:
        if type(item) == dict:
            dict_parse(item, without_write)
        if type(item) == list:
            list_parse(item, without_write)
        else:
            continue


def main(doc):
    """Main function.
    Get JSON object from pandoc, parse it, save result.
    Args:
        doc - json object as python dictionary or list.
              In case of dictionary it has representation like:
              { 'pandoc-version': ...
                'meta': ...
                'blocks': .......}
              in blocks we have all file-content, we will parse doc['blocks'].
              In case of list it has representation like:
              [[info_list], [content_list]], so we will parse doc[1].
    Raises:
        PermissionError: when we can't write output file.
    """
    global table
    output = args.output

    if type(doc) == dict:
        list_parse(doc['blocks'])
    elif type(doc) == list:
        list_parse(doc[1])
    else:
        print('Incompatible Pandoc version')

    write_sheet()

    try:
        ods.save(output)
    except PermissionError as err:
        print("No access to ", output)
        print(err.strerror)


if __name__ == '__main__':
    source = args.input

    # Use subprocess to call pandoc and convert input file into json-object.

    command = 'pandoc ' + source + ' -t json'
    proc = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
    res = proc.communicate()

    if res[1]:
        print(str(res[1]))  # sending stderr output to user
    else:
        if args.separator is None:
            d = json.loads(res[0])
            main(d)
        if args.separator is not None:
            try:
                s = int(args.separator[0])
                separator = s
                d = json.loads(res[0])
                main(d)
            except IndexError:
                print('You entered invalid separator')
            except ValueError:
                print('You entered invalid separator')