import argparse
import parse
import docx_writer
import uploader
from oauth2client import tools


parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter, parents=[tools.argparser])
parser.add_argument('input', help='Input file. Use Pandoc`s input formats', action='store')
parser.add_argument('output', help='Output file.', action='store')
args = parser.parse_args()


def main():
    upl = uploader.Uploader(args)
    doc = docx_writer.DocxWriter()
    parse.main(args.input, doc)
    doc.save(args.output + '.docx')
    upl.upload()


if __name__ == '__main__':
    main()
