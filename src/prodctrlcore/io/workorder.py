
from os import makedirs
from os.path import join, exists
from xlwings import Book

from .header import HeaderParser

WORKORDER_DIR = r"\\hssieng\DATA\HS\SAP - Material Master_BOM\SigmaNest Work Orders"
TEMPLATE = join(WORKORDER_DIR, "TagSchedule_Template.xls")


class WorkOrder(Book):

    def __init__(self, job=None, shipment=None, job_shipment=None, **kwargs):
        if job and shipment:
            job_shipment = '{}-{}'.format(job, int(shipment))
        self.job_shipment = job_shipment
        assert self.job_shipment is not None

        if exists(self.file):
            self.init_file(self.file, **kwargs)
        else:
            self.init_file(TEMPLATE, **kwargs)
            if not exists(self.year_folder):
                makedirs(self.year_folder)
            self.save(self.file)

    def init_file(self, file, **kwargs):
        super().__init__(file, **kwargs)

    @param
    def year_folder(self):
        job_year = '20' + self.job_shipment[1:3].zfill(2)

        return join(WORKORDER_DIR, job_year + " Work Orders Created")

    @param
    def file(self):
        xl_file = '{}.xls'.format(self.job_shipment)

        return join(WORKORDER_DIR, self.year_folder, xl_file)
