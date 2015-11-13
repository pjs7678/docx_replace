#-*- coding: utf-8 -*-

import argparse
from docx_replace import replace

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Replace words in docx file according to the rules')
    parser.add_argument('--in', dest='infile', required=True, help='input docx file name')
    parser.add_argument('--out', dest='outfile', required=True, help='output docx file name')
    parser.add_argument('--format', dest='filter_format',
                        help="json format file which is used for changing words. It should be like this {'keyword' : 'after_replace'} where in docx {keyword} change to after_replace")
    args = parser.parse_args()

    replace(infile = args.infile,
            outfile = args.outfile,
            filter_format = args.filter_format)

