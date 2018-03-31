import argparse
import parse
import creator
import os
from oauth2client import tools


parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter, parents=[tools.argparser])
parser.add_argument('input', help='Input file. Use Pandoc`s input formats', action='store')
args = parser.parse_args()


def main():
    writer = creator.GDocsWriter(args)
    doc = writer.create_new_document(os.path.basename(args.input))
    parse.main(args.input, doc)


if __name__ == '__main__':
    main()
