import argparse
import parse
import creator
from oauth2client import tools

# parser = argparse.ArgumentParser(description='Generator of Google Documents', parents=[tools.argparser])
parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter, parents=[tools.argparser])
# flags = parser.parse_args()
parser.add_argument('input', help='Input file. Use Pandoc`s input formats', action='store')
args = parser.parse_args()




def main():
    print(args.input)
    writer = creator.GDocsWriter(args)
    doc = writer.create_new_document('Test_Report')
    parse.main(args.input, doc)
    # doc.write_table([['asd', 'wds'], ['bsd', 'lds']])

if __name__ == '__main__':
    main()
