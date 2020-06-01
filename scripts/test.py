import prodctrlcore
from prodctrlcore.io import schedule

import time
from os.path import join, dirname, realpath

BASE_DIR = dirname(realpath(__file__))
DATA_FILE = join(BASE_DIR, "data", "Job Ship Dates.xlsx")

data = schedule.get_update_data(DATA_FILE)

for k, v in sorted(data.items()):
    print(k, v)
