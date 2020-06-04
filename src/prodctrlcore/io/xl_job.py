
from os import makedirs
from os.path import join, exists
from xlwings import Book, Sheet

from re import compile as regex

from .header import HeaderParser

WORKORDER_DIR = r"\\hssieng\DATA\HS\SAP - Material Master_BOM\SigmaNest Work Orders"
TEMPLATE = join(WORKORDER_DIR, "TagSchedule_Template.xls")

JOBSHIP_RE = regex(
    r"1?(?P<year>\d{2})(?P<id>\d{4})(?P<structure>[a-zA-Z]?)-(?P<shipment>\d{0,2})")


class JobBookReader(Book):

    def __init__(self, job, shipment=None, **kwargs):
        groups = JOBSHIP_RE.match(job).groupdict()

        self.job = '1{year}{id}{structure}'.format(**groups)
        self.job_year = '20{year}'.format(**groups)

        self.shipment = int(groups['shipment']) or shipment
        assert self.shipment is not None

        self.job_shipment = '{}-{:0>2}'.format(self.job, int(self.shipment))

        if exists(self.file):
            self.init_file(self.file, **kwargs)
        else:
            self.init_file(TEMPLATE, **kwargs)
            if not exists(self.year_folder):
                makedirs(self.year_folder)
            self.save(self.file)

        super().__init__(file, **kwargs)

    @param
    def year_folder(self):
        return join(WORKORDER_DIR, self.job_year + " Work Orders Created")

    @param
    def file(self):
        xl_file = '{}.xls'.format(self.job_shipment)

        return join(WORKORDER_DIR, self.year_folder, xl_file)


class JobSheetReader(Sheet):

    def __init__(self, sheet, **kwargs):
        self.sheet = self.sheets[sheet_name]

        self.header = HeaderParser(sheet=self)
        header.add_header_aliases(HEADER_ALIASES)

    def init_file(self, file, **kwargs):
        super().__init__(file, **kwargs)
