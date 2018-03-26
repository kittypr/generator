import argparse
import parse
import creater

parser = argparse.ArgumentParser(description='Generator of Google Documents')
parser.add_argument('input', help='Input file. Use Pandoc`s input formats', action='store')
args = parser.parse_args()


def main():
    print(args.input)
    parse.main(args.input)


if __name__ == '__main__':
    main()
