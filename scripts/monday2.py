
from os import getenv

from moncli import MondayClient, create_column_value, ColumnType
from prodctrlcore.hssformats import schedule
from prodctrlcore.utils import CountingIter

import re
import logging

from os.path import join, dirname, realpath

logger = logging.getLogger(__name__)

BASE_DIR = dirname(realpath(__file__))
DATA_FILE = join(BASE_DIR, "data", "Job Ship Dates.xlsx")

TOKEN = getenv('MONDAY_TOKEN')
SKIP_GROUPS = ['Jobs Completed Through PC']
JOB_REGEX = re.compile("([A-Z])-([0-9]{7})[A-Z]?-([0-9]{1,2})")

col_id_map = dict(
    pm='text',
    product='text6',
    bay='text2',
    early_start='early_start',
    main_start='date',
)
col_type_map = dict(
    pm=ColumnType.text,
    product=ColumnType.text,
    bay=ColumnType.text,
    early_start=ColumnType.date,
    main_start=ColumnType.date,
)


def main():
    jobs = schedule.get_job_ship_dates(DATA_FILE)

    keys = list(jobs.keys())
    for key in keys:
        without_struct = '-'.join(JOB_REGEX.match(key).groups())
        jobs[without_struct] = jobs[key]
        del jobs[key]

    client = MondayClient('pimiller@high.net', TOKEN, TOKEN)
    board = client.get_board(name='Jobs')

    for group in board.get_groups():
        if group.title in SKIP_GROUPS:
            continue

        count = 0
        for item in group.get_items():
            if item.name in jobs:
                print('\r[{}] Updating: {}'.format(count, item.name), end='')

                update_vals = list()
                for k, v in jobs[item.name].items():
                    col_id = col_id_map[k]
                    col_type = col_type_map[k]

                    if v is None:
                        kwargs = {}
                    elif col_type is ColumnType.date:
                        kwargs = {'date': v.date().isoformat()}
                    else:
                        kwargs = {'value': v}

                    col = create_column_value(col_id, col_type, **kwargs)

                    update_vals.append(col)

                item.change_multiple_column_values(column_values=update_vals)
                logger.info("Updating {}: {}".format(
                    item.name, str(update_vals)))
                count += 1
            else:
                logger.info("Not in active jobs: {}".format(item.name))


if __name__ == "__main__":
    main()
