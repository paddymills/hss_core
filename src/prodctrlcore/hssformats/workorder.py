
from glob import glob
from os.path import join, basename
from re import compile as regex

from prodctrlcore.io import JobBookReader, JobParser
from ._alias import workorder as HEADER_ALIASES

WORKORDER_DIR = r"\\hssieng\DATA\HS\SAP - Material Master_BOM\SigmaNest Work Orders"
TEMPLATE = "TagSchedule_Template.xls"

ADD_KEYS = [
    'TransType',
    'District',
    'ItemName',
    'Qty',
    'Material',
    'DueDate',
    'Customer',
    'DwgNumber',
    'Remark',
    'ItemData1',
    'ItemData2',
    'ItemData3',
    'ItemData4',
    'Process',
    'ChargeRefNumber',
    'Operation2',
    'Operation3',
    'Operation4',
    'Mark',
    'MaterialMaster',
    'RawSize',
    'PartSize',
    'HeatMarkKeyWord',
]

FOLDER_SUFFIX = " Work Orders Created"
FILE_SUFFIX = "_SimTrans_WO"


class WorkOrder(JobBookReader):

    def __init__(self, job, shipment=None, **kwargs):
        kwargs.update(
            directory=WORKORDER_DIR,
            template=TEMPLATE,
            folder_suffix=FOLDER_SUFFIX,
            file_suffix=FILE_SUFFIX
        )
        super().__init__(job, shipment, **kwargs)

        self.data_sheet = self.sheet('WorkOrders_Template', header_range='A2')
        self.data_sheet.add_header_aliases(HEADER_ALIASES)

    def add(self, row):
        row_data = dict()
        for key in ADD_KEYS:
            row_data[key] = row.get_item(key)

        self.data_sheet.add_row(**row_data)


class WorkOrderJobData:
    # get all data for work orders for a given structure
    # i.e. '1180078B'

    def __init__(self, job):
        JobParser(job, assign_to=self)
        glob_path = join(
            WORKORDER_DIR,
            self.job_year + FOLDER_SUFFIX,
            "{}*.xls".format(job))

        self._data = dict()
        for path in glob(glob_path):
            # JobParser can parse files starting with job-shipment
            file = WorkOrder(basename(path))

            for row in file.data_sheet.iter_rows():
                if row.mark in self._data:
                    stored = self._data[row.mark]

                    # update grade and operations, if not None
                    stored.matl = stored.matl or row.matl
                    stored.remark = stored.remark or row.remark
                    stored.op2 = stored.op2 or row.op2
                    stored.op3 = stored.op3 or row.op3
                    stored.op4 = stored.op4 or row.op4
                else:
                    self._data[row.mark] = row

    def __getattr__(self, name):
        if name in self._data:
            return self.get_part(name)

        raise AttributeError

    def get_part(self, mark):
        return self._data.get(mark, None)
