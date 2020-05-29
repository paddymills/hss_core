
import logging

from os import scandir
from os.path import join, exists
from os.path import dirname, realpath

from argparse import ArgumentParser
from datetime import datetime
from re import compile as regex

from prodctrlcore.io import schedule
from prodctrlcore.monday.custom import DevelopmentJobBoard, JobBoard

from collections import defaultdict

BASE_DIR = dirname(realpath(__file__))
DATA_FILE = join(BASE_DIR, "data", "Job Ship Dates.xlsx")

timestamp = datetime.now().date().isoformat()
LOG_FILE = join(BASE_DIR, "logs", '{}.log'.format(timestamp))
logging.basicConfig(filename=LOG_FILE, level=logging.INFO)

JobBoardType = JobBoard
# JobBoardType = DevelopmentJobBoard
DATE_REGEX = "([0-9]{1,2})/([0-9]{1,2})/([0-9]{0,4})"

logger = logging.getLogger(__name__)


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

    init_str = "\n\n"
    init_str += tilde_str("PROCESS START") + "\n"
    init_str += tilde_str(datetime.now().isoformat()) + "\n"
    init_str += 'Args: {}'.format(args) + "\n\n"
    logger.info(init_str)

    if args.dev:
        JobBoardType = DevelopmentJobBoard

    if args.restore:
        # get file
        date_match = regex(DATE_REGEX).fullmatch(args.file)

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

        elif date_match:

            if len(date_match.groups()) == 3:
                month, day, year = date_match.groups()
            else:  # month and day only
                month, day = date_match.groups()
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
        update_job_board(updates)

    else:
        update_job_board()


def log_file(log_file_name):
    if log_file_name.endswith('.log'):
        return join(BASE_DIR, "logs", log_file_name)
    return join(BASE_DIR, "logs", '{}.log'.format(log_file_name))


def update_job_board(jobs=None):
    job_board = JobBoardType()

    if jobs in None:
        jobs = schedule.get_update_data(DATA_FILE)

    # update monday.com board
    i = 1
    total = len(jobs)
    for job, kwargs in jobs.items():
        print("\r[{}/{}] Running updates".format(i, total), end='')
        # logger.info("Updating Job: [{}] {}".format(job, kwargs))
        job_board.update_job_data(job, **kwargs)

        i += 1


def parse_log_file(log_file):
    LOGGING_FORMAT = regex(
        r"(?P<msg_type>[A-Z]+):(?P<logger>[\w.]):(?P<update_type>\w):(?P<job>[\w-])/(?P<column>\w):(?P<old_val>.)->(?P<new_val>.)")
    # INFO:prodctrlcore.monday.custom:UPDATE:D-1160253C-04/main_start:None->{'date': '2020-11-25', 'changed_at': '2020-05-29T14:19:10.745Z'}

    jobs = defaultdict(dict)

    with open(log_file, 'r') as restore_file_stream:
        for line in restore_file_stream:
            match = LOGGING_FORMAT.match(line)
            if match:
                m_g = match.groups
                jobs[m_g('job')][m_g('column')] = m_g('old_val')

    return jobs


def tilde_str(str_val, str_len=50):
    str_val = "~ {} ~".format(str_val)
    while len(str_val) < str_len:
        str_val = "~{}~".format(str_val)

    return str_val


if __name__ == "__main__":
    main()
