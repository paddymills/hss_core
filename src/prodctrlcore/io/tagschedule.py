
from os import makedirs
from os.path import join, exists
from xlwings import Book

from .header import HeaderParser

TAGSCHED_DIR = r"\\hssfileserv1\HSSShared\HSSI Lean\CAD-CAM\TagSchedule"
TEMPLATE = join(TAGSCHED_DIR, "TagSchedule_Template.xls")


class TagSchedule(Book):

    def __init__(self, job=None, shipment=None, job_shipment=None, **kwargs):
        if job and shipment:
            job_shipment = '{}-{}'.format(job, int(shipment))
        self.job_shipment = job_shipment
        assert self.job_shipment is not None

        self.job_year = '20' + self.job_shipment[1:3].zfill(2)

        self.year_folder = join(TAGSCHED_DIR, self.job_year)
        self.file = join(self.year_folder, '{}.xls'.format(self.job_shipment))

        if exists(self.file):
            self.init_file(self.file, **kwargs)
        else:
            self.init_file(TEMPLATE, **kwargs)
            if not exists(self.year_folder):
                makedirs(self.year_folder)
            self.save(self.file)

    def init_file(self, file, **kwargs):
        super().__init__(file, **kwargs)
        self.webs = WebFlangeSheet(self.sheets['WEBS'])
        self.flanges = WebFlangeSheet(self.sheets['FLANGES'])
        self.code = CodeDeliverySheet(self.sheets['CODE DELIVERY'])


class TagSchedSheet:

    def __init__(self, sheet, header_rng, first_data_row=2):
        if type(header_rng) is not list:
            header_rng = [header_rng]

        _header = list()
        for rng in header_rng:
            _header.extend(sheet.range(rng).value)
        self.header = HeaderParser(header=_header)

        self.min_col = min([sheet.range(x).row for x in header_rng])
        self.max_col = max([sheet.range(x).last_cell.row for x in header_rng])

        self.sheet = sheet
        self.first_data_row = first_data_row

    def _data_range(self):
        start = (self.first_data_row, self.min_col)
        end = (self.first_data_row, self.max_col)
        return self.sheet.range(start, end).expand('down')

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

            row = self.header.parse_row(self_row)
            for key, value in kwargs:
                if self_row.get_item(key) != value:
                    break  # kwargs for-loop

            # all kwargs matched row -> row exists in sheet
            else:
                break  # rows for-loop

        # row not in sheet
        else:
            row_index = self._data_range().last_cell.row + 1

            col_index = self._data_range().column
            self.sheet.range(row_index, col_index).value = row

    def construct_row_from_kwargs(self, kwargs):
        row = list()

        self.header.clear_parsed()
        for key, value in kwargs:
            index = self.header.get_index(key)

            try:
                row[index] = value
            except IndexError:
                while len(row) < index:
                    row.append(None)
                row.append(key)

        return row


class WebFlangeSheet(TagSchedSheet):

    def __init__(self, sheet):
        SHEET_KWARGS = dict(
            header_range=['C1:N1', 'O2:N2'],
            first_data_row=4,
        )
        super().__init__(sheet, **SHEET_KWARGS)


class CodeDeliverySheet(TagSchedSheet):

    def __init__(self, sheet):
        SHEET_KWARGS = dict(
            header_range='A2:G2',
        )
        super().__init__(sheet, **SHEET_KWARGS)
