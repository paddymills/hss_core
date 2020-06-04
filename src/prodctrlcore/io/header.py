
import inflection

from xlwings import Range
from re import compile as regex

PARAM_RE = regex(r'(?P<text>[a-zA-Z_]+)(?P<id>[0-9]*)')


class HeaderParser:

    """
        HeaderParser: A class to parse an excel header row

        Can be used multiple ways:
        1) To access column ids based off of header text
        2) To access items of a row based off of header text

        To access row items, must parse the row before use

        Can also pass in partial header names to access columns
        for example:
            'op1' matches 'Operation1'
            'matl' matches 'Material'
        In these methods, trailing ids (if applicable) must match
    """

    def __init__(self, sheet, header=None, header_range="A1", expand_header=True):

        if header:
            self.header = header
        else:
            rng = sheet.range(header_range)
            if expand_header:
                rng = rng.expand('right')

            self.header = rng.value

        self.indexes = dict()
        self._init_header()

        self.row = None

    def __getattr__(self, name):
        try:
            index = self.indexes[name]
        except KeyError:
            index = self.infer_key(name)

        if self.row:
            return self.row[index]

        return index

    def _init_header(self):
        for index, column in enumerate(self.header):
            self.indexes[column] = index
            self.indexes[column.lower()] = index
            self.indexes[to_(column)] = index

    def parse_row(self, row):
        if type(row) is Range:
            row = row.value

        self.row = row

    def infer_key(self, key):
        key = key.lower()

        _m = PARAM_RE.match(key)
        key_text = _m.group('text')
        key_id = _m.group('id')

        # infer based on key that has the
        # - str.startswith match
        # - same trailing ID (if applicable)
        for col in self.indexes:
            _m = PARAM_RE.match(col)
            col_text = _m.group('text')
            col_id = _m.group('id')

            if key_id == col_id:
                if col_text.startswith(key_text):
                    return self.indexes[col]

        # infer based on key that has the
        # - key is found in column text
        # - same trailing ID (if applicable)
        # - CRITICAL: can only have one match
        match = None
        for col in self.indexes:
            _m = PARAM_RE.match(col)
            col_text = _m.group('text')
            col_id = _m.group('id')

            if key_id == col_id:
                if semi_sequential_match(key_text, col_text):
                    if match:
                        raise KeyError(
                            "Multiple matching keys found when inferring column headers. Key={}".format(key))
                    return self.indexes[col]

        raise KeyError


def to_(text):
    return inflection.parameterize(text, separator='_')


def semi_sequential_match(find_str, within_str):
    """
        Returns if find string exists in
        search string sequentially.

        Find string can skip characters in
        search string.

        i.e. find 'matl' in 'material'
    """

    find_chars = list(find_str)
    current_char = find_chars.pop(0)
    for char in within_str:
        if char == current_char:
            if find_chars:
                char = find_chars.pop(0)
            else:
                return False

    return True
