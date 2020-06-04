
from .xljobfile import JobBookReader

WORKORDER_DIR = r"\\hssieng\DATA\HS\SAP - Material Master_BOM\SigmaNest Work Orders"
TEMPLATE = "TagSchedule_Template.xls"

HEADER_ALIASES = {
    'Part': 'ItemName',
    'State': 'Customer',
    'Job': 'ItemData1',
    'Ship': 'ItemData2',
    'Shipping Group': 'ItemData4',
    'ChargeRefNumber': 'SAP Network Number',
    'Mark': 'Operation5',
    'MaterialMaster': 'Operation6',
    'HeatMarkKeyWord': 'Operation10'
}


class WorkOrder(JobBookReader):

    def __init__(self, job, shipment=None, **kwargs):
        kwargs.update(
            directory=WORKORDER_DIR,
            template=TEMPLATE,
            folder_suffix=" Work Orders Created"
        )

        super().__init__(job, shipment, **kwargs)

    def get_data_sheet(self):
        _sheet = self.sheet('WorkOrders_Template', header_range='A2')
        _sheet.header.add_header_aliases(HEADER_ALIASES)

        return _sheet


def get_all_workorder_data_for_structure(structure):
    # TODO: glob to find files
    # get all data for work orders for a given structure
    # i.e. '1170082C'
    pass
