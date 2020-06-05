
from os import makedirs
from os.path import join, exists
from xlwings import Book, Sheet

from re import compile as regex

from . import HeaderParser, ParsedRow


JOBSHIP_RE = regex(
    r"1?(?P<year>\d{2})(?P<id>\d{4})(?P<structure>[a-zA-Z]?)-?(?P<shipment>\d{0,2})")


class JobParser:

    def __init__(self, job, shipment=1, assign_to=None):
        match = JOBSHIP_RE.match(job)

        if match is None:
            raise ValueError(
                "[{}] does not match expected pattern".format(job))

        groups = match.groupdict()

        self.job = '1{year}{id}{structure}'.format(**groups)
        self.job_year = '20{year}'.format(**groups)
        self.shipment = int(groups['shipment'] or shipment)

        # add to other objects attributes
        if assign_to:
            assign_to.__dict__.update(self.__dict__)


class JobBookReader(Book):
    """
        Excel Book reader for jobs that are stored by year
        i.e. directory > 2020 > Job-Shipment.xls

        if file does not exist,
        template file will be created and saved in place
    """

    def __init__(self, job, shipment=None, **kwargs):
        JobParser(job, shipment, assign_to=self)

        self.job_shipment = '{}-{}'.format(self.job, self.shipment)
        self.proper_job_shipment = '{}-{:0>2}'.format(self.job, self.shipment)

        self.folder_suffix = kwargs.get('folder_suffix', '')
        self.file_suffix = kwargs.get('file_suffix', '')

        self.root_dir = kwargs.get('directory')
        self.template = join(self.root_dir, kwargs.get('template'))

        if exists(self.file):
            self.__init_file__(self.file, **kwargs)
        else:
            self.__init_file__(self.template, **kwargs)
            if not exists(self.year_folder):
                makedirs(self.year_folder)
            self.save(self.file)

    def __init_file__(self, file, **kwargs):
        super().__init__(file, **kwargs)

    @property
    def year_folder(self):
        return join(self.root_dir, self.job_year + self.folder_suffix)

    @property
    def file(self):
        xl_file = '{}{}.xls'.format(self.job_shipment, self.file_suffix)

        return join(self.root_dir, self.year_folder, xl_file)

    def sheet(self, sheet_name, **kwargs):
        sheet = self.sheets[sheet_name].impl

        return JobSheetReader(impl=sheet, **kwargs)


class JobSheetReader(Sheet):

    def __init__(self, sheet=None, **kwargs):
        if 'impl' in kwargs:
            super().__init__(impl=kwargs['impl'])
        else:
            super().__init__(sheet)

        if 'header_range' in kwargs:
            self.set_header(**kwargs)
        else:
            self.header = HeaderParser(sheet=self, **kwargs)
            self.first_data_row = 2

    def set_header(self, header_range, first_data_row=0):
        self.first_data_row = first_data_row

        if type(header_range) is not list:
            header_rng = [header_rng]

        _header = list()
        for range in header_range:
            _rng = self.range(range)
            _header.extend(_rng.value)
            if self.first_data_row <= _rng.last_cell.row:
                self.first_data_row = _rng.last_cell.row + 1

        self.header = HeaderParser(header=_header)

    def _data_range(self):
        start = (self.first_data_row, self.min_col)
        end = (self.first_data_row, self.max_col)

        return self.range(start, end).expand('down')

    @property
    def min_col(self):
        return min(self.header.indexes.values())

    @property
    def max_col(self):
        return max(self.header.indexes.values())

    def get_rows(self):
        return self._data_range().value

    def iter_rows(self):
        for row in self.get_rows():
            yield self.header.parse_row(row)

    def add_row(self, row=None, compare_cols=list(), **kwargs):
        if type(row) is ParsedRow:
            new_row = row
        else:
            new_row = self.construct_row(row, **kwargs)

        # in case using compare_cols (only need to do this one once)
        new_cols = map(new_row.get_item, compare_cols)

        # check if row already exists
        for self_row in self.iter_rows():

            # compares selected columns to compare
            if compare_cols:
                self_cols = map(self_row.get_item, compare_cols)

                if new_cols == self_cols:
                    break

            # compares all indexed columns
            elif self_row == new_row:
                break

        # row not in sheet
        else:
            row_index = self._data_range().last_cell.row + 1

            col_index = self._data_range().column
            self.range(row_index, col_index).value = new_row._data

    def construct_row(self, row=None, **kwargs):
        if row:
            return ParsedRow(row, self.header)

        # construct row from scratch
        max_index = max(self.header.indexes.values())
        blanks = [None] * (max_index + 1)

        row = ParsedRow(blanks, self.header)

        for key, value in kwargs.items():
            index = self.header.get_index(key)
            row[index] = value

        return row
