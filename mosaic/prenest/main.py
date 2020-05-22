
import pandas
import logging
import os
import time

from mosaic.prenest.workbook import PrenestDataWorkbook
from argparse import ArgumentParser, REMAINDER

from mosaic import config

def main():
    parser = ArgumentParser()
    parser.add_argument('--caller',  '-c', help='Calling file full name')
    parser.add_argument('--job',     '-j', help='Job number', default='')
    parser.add_argument('--fetch',   '-f', action='store_true', help='Fetch flange data from engineering')
    parser.add_argument('--manual',  '-m', action='store_true', help='Generate flange data from manual entry')
    parser.add_argument('--load',    '-l', action='store_true', help='Load flange data')
    parser.add_argument('--analyze', '-a', action='store_true', help='Analyze nesting')
    parser.add_argument('--export',  '-e', action='store_true', help='Export to SigmaNest for nesting')
    parser.add_argument('args', nargs=REMAINDER)
    args = parser.parse_args()

    logging.info(args)

    wb = PrenestDataWorkbook(args.job, args.caller)

    if args.fetch:
        wb.import_flange_data()
    if args.manual:
        wb.generate_flange_data_manual()
    if args.load:
        wb.import_flange_data()
    if args.analyze:
        wb.analyze()
    if args.export:
        wb.export_for_prenesting()


if __name__ == "__main__":
    LOG_FILE = os.path.join(os.path.dirname(__file__), 'log.txt')
    logging.basicConfig(filename=LOG_FILE, filemode='w', level=logging.INFO)
    try:
        main()
    except Exception as error:
        logging.exception(error)
