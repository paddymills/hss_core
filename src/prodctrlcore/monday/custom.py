
import json
import re

from os import getenv
from os.path import realpath, dirname
from inflection import underscore
from datetime import datetime
from functools import partial

import logging

from .client import MondayBoardClient, js_utc_now

logger = logging.getLogger(__name__)

JOB_REGEX = re.compile("[A-Z]-([0-9]{7})[A-Z]?-([0-9]{1,2})")
JOB_FORMAT = "{}-{:0>2}"

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
# convenience classes for accessing specific boards #
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #

# the Jobs board


class JobBoard(MondayBoardClient):

    def __init__(self, **kwargs):
        init_kwargs = dict(
            endpoint='https://api.monday.com/v2',
            board_name='Jobs',
            token=getenv('MONDAY_TOKEN'),  # user environment variable
            skip_groups=['Jobs Completed Through PC'],
        )
        init_kwargs.update(kwargs)
        super().__init__(**init_kwargs)

        self.job_ids = dict()
        self.column_aliases = dict(
            job='name',
            product='Type',
            bay='Location',
        )

        self.init_job_board()
        logger.info("{} board initialized".format(init_kwargs['board_name']))

    def init_job_board(self):
        response = self.execute('get_jobs', group_ids=self.groups)

        for group in response['groups']:
            for item in group['items']:
                self.job_ids[item['name']] = int(item['id'])

    def execute(self, query, **kwargs):
        # TODO: log updates
        return self._board_execute(query, board_id=self.board_id, **kwargs)

    def update_job_data(self, job, **kwargs):
        job_id = self.get_job_id(job)
        exec_job = partial(self.execute, job_id=job_id)

        if job_id is None:
            logger.error("Job not in monday.com active groups: " + job)
            return

        for key, val in kwargs.items():
            if key in self.column_aliases:
                key = self.column_aliases[key]
            column_id = self.columns[key]['id']
            column_type = self.columns[key]['type']

            if type(val) is datetime:
                val = val.date().isoformat()

            # get job data
            response = exec_job('get_job_data', column_ids=[column_id])
            data = response['items'][0]['column_values'][0]
            if data['text'] != val:
                if column_type == 'date':
                    val = {'date': val, 'changed_at': js_utc_now()}

                exec_job('update_job', column_id=column_id,
                         column_val=json.dumps(val))
                update_type = "UPDATE"
            else:
                update_type = 'SKIP'
                val = data['text']
            logger.info(
                '{}:{}/{}:{}->{}'.format(update_type, job, key, data['text'], val))

    def get_job_id(self, job):
        if job in self.job_ids:
            return self.job_ids[job]

        # JOB_REGEX = re.compile("[A-Z]-([0-9]{7})[A-Z]?-([0-9]{1,2})")
        # JOB_FORMAT = "{}-{:0>2}"

        job_without_structure = JOB_FORMAT.format(
            *JOB_REGEX.match(job).groups())
        for _monday_job, _id in self.job_ids.items():
            match = JOB_REGEX.match(_monday_job)
            if match:
                monday_job = JOB_FORMAT.format(*match.groups())
                if monday_job == job_without_structure:
                    # # update monday job name
                    # logger.info("Updating job name: {} -> {}".format(_monday_job, job))
                    # self.execute('update_job', job_id=_id, column_id='name', column_val=job)
                    # # Did not work. I think it is because 'name' is not a column_value

                    return _id

        return None


class DevelopmentJobBoard(JobBoard):

    def __init__(self, **kwargs):
        super().__init__(board_name='Development', **kwargs)
