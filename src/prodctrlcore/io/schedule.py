
import os
import xlwings

from collections import defaultdict
from argparse import ArgumentParser
from prodctrlcore.monday.client import JobBoard, DevelopmentJobBoard

BASE_DIR = os.path.dirname(os.path.realpath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "Job Ship Dates.xlsx")

COST_CENTERS = {
    2005: "WB",
    2006: "NB",
    2007: "SB",
    2030: "MBN",
    2031: "MBS",
}

# JobBoardType = JobBoard
JobBoardType = DevelopmentJobBoard


def init_argparser():
    parser = ArgumentParser()

    parser.add_argument('-d',
                        '--dev',     action='store_true',        help='Execute on development instance')
    parser.add_argument('-r',
                        '--restore', action='store_true',        help='Execute restore mode')
    parser.add_argument('-f',
                        '--file',    action='store',             help='File to restore(restore flag only)')
    parser.add_argument('-j',
                        '--jobs',
                        '--job',     action='extend', nargs='*', help='Jobs to update/restore')

    return parser.parse_args()


def main():
    args = init_argparser()

    if args.dev:
        JobBoardType = DevelopmentJobBoard

    if args.restore:
        # get file
        if args.file == 'last':
            restore_file = 'most recent file'
        else:
            restore_file = args.restore

        with open(restore_file, 'r') as restore_file_stream:
            data = restore_file_stream.read()
        # restore_job_board(data, args.jobs)
    else:
        print("Update")
        # update_job_board(args.jobs)


def update_job_board(jobs=None):
    job_board = JobBoardType()
    wb = xlwings.Book(DATA_FILE)
    jobs = defaultdict(dict)

    # Early Start and Main Start dates
    i = 2
    sheet = wb.sheets['Dates']
    while sheet.range(i, 1).value:
        job = sheet.range(i, 1).value
        jobs[job]['early_start'] = sheet.range(i, 2).value
        jobs[job]['main_start'] = sheet.range(i, 3).value

        i += 1

    # Project Manager
    i = 2
    sheet = wb.sheets['PM']
    while sheet.range(i, 1).value:
        job, pm = sheet.range(i, 1).value
        jobs[job]['pm'] = sheet.range(i, 2).value

        i += 1

    # Products/Types
    i = 2
    sheet = wb.sheets['Products']
    job, products = None, list()
    while sheet.range(i, 2).value:
        if sheet.range(i, 1).value:
            if job:  # not first row
                jobs[job]['text6'] = ','.join(products)
            job = sheet.range(i, 1).value
            products = list()

        products.append(sheet.range(i, 2).value)
        i += 1

    # Fab Bays
    i = 2
    sheet = wb.sheets['Bays']
    job, bays = None, list()
    while sheet.range(i, 2).value:
        if sheet.range(i, 1).value:
            if job:  # not first row
                jobs[job]['text2'] = ','.join(bays)
            job = sheet.range(i, 1).value
            bays = list()

        bays.append(sheet.range(i, 2).value)
        i += 1

    # update monday.com board
    for job, kwargs in jobs.items():
        job_board.update_job_data(job, **kwargs)

    wb.save()
    if len(wb.app.books) > 1:
        wb.close()
    else:
        wb.app.quit()


def restore_job_board(data, jobs=None):
    job_board = JobBoardType()


if __name__ == "__main__":
    main()
