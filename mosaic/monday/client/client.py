
import os
import json

from inflection    import underscore
from graphqlclient import GraphQLClient
from datetime      import datetime, timezone

ROOT_DIRECTORY = os.path.realpath(os.path.dirname(__file__))


class MondayBoardClient(GraphQLClient):

    def __init__(self, endpoint, board_name=None, token=None, skip_groups=[]):
        super().__init__(endpoint)

        self.board_id = None
        self.complexity = 10000000    # base complexity is 10,000,000 per minute

        self.scripts = dict()
        self.groups = list()
        self.columns = dict()

        self.skip_groups = skip_groups

        if token:
            self.inject_token(token)
            if board_name:
                self.init_board(board_name)

        # pre-load GraphQL scripts
        for script in os.scandir(os.path.join(ROOT_DIRECTORY, 'graphql')):
            self.add_script(script.name, script.path)


    def inject_token(self, token):
        if os.path.exists(token):
            with open(token, 'r') as token_file_stream:
                token = token_file_stream.readline()

        self.inject_token(token)


    def init_board(self, board_name):
        # get board by name
        response = self.execute('get_boards')
        for board in response['boards']:
            if board['name'] == board_name:
                self.board_id = board['id']
                break
                
        # get group and column ids
        response = self.execute('get_board_config', board_id=self.board_id)
        for group in response['groups']:
            if group['title'] not in self.skip_groups:
                self.groups.append(group['id'])
        for column in response['columns']:
            val = dict(id=column['id'], type=column['type'])
            self.columns[column['title']] = val
            self.columns[underscore(column['title'].replace(' ',''))] = val  # Early Start -> early_start


    def execute(self, query, variables=dict(), **kwargs):
        variables.update(kwargs)

        if query in self.scripts.keys():
            query = self.scripts['query']
        response = json.loads(super().execute(query, variables=variables))

        if "data" in response.keys():
            if "complexity" in response['data'].keys():
                self.complexity = response['data']['complexity']['after']

            return response['data']['boards']
        elif "errors" in response.keys():
            # TODO: log errors
            #     response['errors']['message']
            # and response['errors']['problems']['explanation']
            return response['errors']

        return response


    def add_script(self, script_name, script_file):
        script_name = underscore(script_name.split('.')[0])
        with open(script_file, 'r') as script_file_stream:
            self.scripts[script_name] = script_file_stream.read()


    def get_complexity(self):
        return self.execute('get_complexity')


def js_utc_now():
    return datetime.now(tz=timezone.utc).isoformat(timespec='milliseconds').replace('+00:00', 'Z')
