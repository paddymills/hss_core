
from os import makedirs
from os.path import join, exists
from xlwings import Book, Sheet

from re import compile as regex

from .header import HeaderParser


JOBSHIP_RE = regex(
    r"1?(?P<year>\d{2})(?P<id>\d{4})(?P<structure>[a-zA-Z]?)-(?P<shipment>\d{0,2})")


class JobBookReader(Book):
    """
        Excel Book reader for jobs that are stored by year
        i.e. directory > 2020 > Job-Shipment.xls

        if file does not exist,
        template file will be created and saved in place
    """

    def __init__(self, job, shipment=None, **kwargs):
        groups = JOBSHIP_RE.match(job).groupdict()

        self.job = '1{year}{id}{structure}'.format(**groups)
        self.job_year = '20{year}'.format(**groups)

        self.shipment = int(groups['shipment'] or shipment)

        self.job_shipment = '{}-{}'.format(self.job, self.shipment)
        self.proper_job_shipment = '{}-{:0>2}'.format(self.job, self.shipment)

        self.suffix = kwargs.get('folder_suffix', '')
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
        return join(self.root_dir, self.job_year + self.suffix)

    @property
    def file(self):
        xl_file = '{}.xls'.format(self.job_shipment)

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

    def add_row(self, row=None, **kwargs):
        if row is None:
            row = self.construct_row_from_kwargs(kwargs)

        # check if row already exists
        for self_row in self.get_rows():

            # essentially, self_row.startswith(row)
            if row:
                if self_row[:len(row)] == row:
                    break
                continue

            # all kwargs matched row -> row exists in sheet
            else:
                break  # rows for-loop

        # row not in sheet
        else:
            row_index = self._data_range().last_cell.row + 1

            col_index = self._data_range().column
            self.range(row_index, col_index).value = row

    def construct_row_from_kwargs(self, kwargs):
        row = list()

        self.header.clear_parsed()
        for key, value in kwargs.items():
            index = self.header.get_index(key)

            try:
                row[index] = value
            except IndexError:
                while len(row) < index:
                    row.append(None)
                row.append(key)

        return row
