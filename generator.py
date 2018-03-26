import argparse
import parse
import creator

parser = argparse.ArgumentParser(description='Generator of Google Documents')
parser.add_argument('input', help='Input file. Use Pandoc`s input formats', action='store')
args = parser.parse_args()




def main():
    print(args.input)
    writer = creator.GDocsWriter()
    writer.create_new_document('Test_Report')
    parse.main(args.input, writer)


if __name__ == '__main__':
    main()
