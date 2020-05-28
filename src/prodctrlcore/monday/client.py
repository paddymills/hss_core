
import json

from os import scandir
from os.path import join, exists, realpath, dirname
from inflection import underscore
from graphqlclient import GraphQLClient
from datetime import datetime, timezone

import logging

ROOT_DIRECTORY = realpath(dirname(__file__))


class MondayBoardClient(GraphQLClient):

    def __init__(self, endpoint, board_name=None, token=None, skip_groups=[]):
        super().__init__(endpoint)

        self.board_id = None
        self.complexity = 10000000    # base complexity is 10,000,000 per minute

        self.scripts = dict()
        self.groups = list()
        self.columns = dict()

        self.skip_groups = skip_groups

        # pre-load GraphQL scripts
        for script in scandir(join(ROOT_DIRECTORY, 'graphql')):
            self.add_script(script.name, script.path)

        if token:
            self.inject_token(token)
            if board_name:
                self.init_board(board_name)

    def inject_token(self, token):
        if exists(token):  # token is path to file
            with open(token, 'r') as token_file_stream:
                token = token_file_stream.readline()

        super(MondayBoardClient, self).inject_token(token)

    def init_board(self, board_name):
        # get board by name
        response = self._board_execute('get_boards')
        for board in response:
            if board['name'] == board_name:
                self.board_id = int(board['id'])
                break
        assert self.board_id is not None

        # get group ids
        response = self._board_execute(
            'get_board_config', board_id=self.board_id)
        for group in response['groups']:
            if group['title'] not in self.skip_groups:
                self.groups.append(group['id'])

        # get column ids
        for column in response['columns']:
            val = dict(id=column['id'], type=column['type'])
            self.columns[column['title']] = val
            # Early Start -> early_start
            self.columns[underscore(column['title'].replace(' ', ''))] = val

    def _board_execute(self, query, variables=dict(), **kwargs):
        variables.update(kwargs)

        if query in self.scripts.keys():
            query = self.scripts[query]

        response = json.loads(
            super(MondayBoardClient, self).execute(query, variables))

        if "data" in response.keys():
            if "complexity" in response['data'].keys():
                self.complexity = response['data']['complexity']['after']

            if len(response['data']['boards']) == 1:
                return response['data']['boards'][0]
            return response['data']['boards']
        elif "errors" in response.keys():
            for err in response['errors']:
                logging.error(err)
            return response['errors']

        return response

    def execute(self, query, variables=dict(), **kwargs):
        # gets overloaded by child classes
        return self._board_execute(query, variables, **kwargs)

    def add_script(self, script_name, script_file):
        script_name = underscore(script_name.split('.')[0])
        with open(script_file, 'r') as script_file_stream:
            self.scripts[script_name] = script_file_stream.read()

    def get_complexity(self):
        return self._board_execute('get_complexity')


def js_utc_now():
    return datetime.now(tz=timezone.utc).isoformat(timespec='milliseconds').replace('+00:00', 'Z')
