
from .xljobfile import JobBookReader

TAGSCHED_DIR = r"\\hssfileserv1\HSSShared\HSSI Lean\CAD-CAM\TagSchedule"
TEMPLATE = "TagSchedule_Template.xls"

WEBFLG_SHEET_KWARGS = dict(
    header_range=['C1:N1', 'O2:N2'],
    first_data_row=4,
)
CODE_SHEET_KWARGS = dict(
    header_range='A2:G2',
)


class TagSchedule(JobBookReader):

    def __init__(self, job, shipment=None, **kwargs):
        kwargs.update(
            directory=TAGSCHED_DIR,
            template=TEMPLATE
        )

        super().__init__(job, shipment, **kwargs)
        self.webs = self.sheet('WEBS', **WEBFLG_SHEET_KWARGS)
        self.flanges = self.sheet('FLANGES', **WEBFLG_SHEET_KWARGS)
        self.code = self.sheet('CODE DELIVERY', **CODE_SHEET_KWARGS)
