
import logging

from os import scandir
from os.path import join, exists
from os.path import dirname, realpath

from argparse import ArgumentParser
from datetime import datetime
from re import fullmatch

from prodctrlcore.io import schedule
from prodctrlcore.monday.custom import DevelopmentJobBoard, JobBoard


BASE_DIR = dirname(realpath(__file__))
DATA_FILE = join(BASE_DIR, "data", "Job Ship Dates.xlsx")

timestamp = datetime.now().date().isoformat()
LOG_FILE = join(BASE_DIR, "logs", '{}.log'.format(timestamp))
logging.basicConfig(filename=LOG_FILE, level=logging.INFO)

# JobBoardType = JobBoard
JobBoardType = DevelopmentJobBoard
DATE_REGEX = "([0-9]{1,2})/([0-9]{1,2})/([0-9]{0,4})"


def init_argparser():
    parser = ArgumentParser()

    parser.add_argument('-d', '--dev',
                        action='store_true',            help='Execute on development instance')
    parser.add_argument('-r', '--restore',
                        action='store_true',            help='Execute restore mode')
    parser.add_argument('-f', '--file',
                        action='store', default="last", help='File to restore(restore flag only)')
    parser.add_argument('-j', '--jobs', '--job',
                        action='extend', nargs='*',     help='Jobs to update/restore')

    return parser.parse_args()


def main():
    args = init_argparser()

    if args.dev:
        JobBoardType = DevelopmentJobBoard

    if args.restore:
        # get file
        if args.file == 'last':
            last_modified = None
            mtime = 0
            for entry in scandir(join(BASE_DIR, "logs")):
                if entry.name.endswith('.log'):
                    if entry.stat().st_mtime > mtime:
                        restore_file = entry.path
                        mtime = entry.stat().st_mtime

        elif exists(log_file(args.file)):
            restore_file = log_file(args.file)

        elif fullmatch(DATE_REGEX, args.file):
            groups = fullmatch(DATE_REGEX, args.file).groups()

            if len(groups) == 3:
                month, day, year = groups
            else:  # month and day only
                month, day = groups
                year = datetime.now().year

            args.file = datetime(year, month, day).date().isoformat()
            restore_file = log_file(args.file)

        else:
            print("Log file specified does not match expected patterns")
            print("->", args.file)

        if not exists(restore_file):
            print("Log file does not exist")
            return

        updates = parse_log_file(restore_file)
        # restore_job_board(updates)

    else:
        update_job_board()


def log_file(log_file_name):
    if log_file_name.endswith('.log'):
        return join(BASE_DIR, "logs", log_file_name)
    return join(BASE_DIR, "logs", '{}.log'.format(log_file_name))


def update_job_board():
    job_board = JobBoardType()

    jobs = schedule.get_update_data(DATA_FILE)

    # update monday.com board
    for job, kwargs in jobs.items():
        job_board.update_job_data(job, **kwargs)


def parse_log_file(log_file):
    with open(restore_file, 'r') as restore_file_stream:
        for line in restore_file_stream:
            pass


def restore_job_board(data):
    job_board = JobBoardType()


if __name__ == "__main__":
    main()
