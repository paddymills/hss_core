
from os.path import join, dirname, realpath
from argparse import ArgumentParser

from prodctrlcore.io import schedule
from prodctrlcore.monday.custom import DevelopmentJobBoard


BASE_DIR = dirname(realpath(__file__))
DATA_FILE = join(BASE_DIR, "data", "Job Ship Dates.xlsx")

# JobBoardType = JobBoard
JobBoardType = DevelopmentJobBoard


def init_argparser():
    parser = ArgumentParser()

    parser.add_argument('-d',
                        '--dev',     action='store_true',            help='Execute on development instance')
    parser.add_argument('-r',
                        '--restore', action='store_true',            help='Execute restore mode')
    parser.add_argument('-f',
                        '--file',    action='store', default="last", help='File to restore(restore flag only)')
    parser.add_argument('-j',
                        '--jobs',
                        '--job',     action='extend', nargs='*',     help='Jobs to update/restore')

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


def update():
    job_board = JobBoardType()

    jobs = schedule.get_update_data(DATA_FILE)

    # update monday.com board
    for job, kwargs in jobs.items():
        job_board.update_job_data(job, **kwargs)


def restore_job_board(data, jobs=None):
    job_board = JobBoardType()


if __name__ == "__main__":
    main()
