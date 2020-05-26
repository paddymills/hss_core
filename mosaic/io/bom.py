#!/usr/bin/env python

import xlwings
import os
import re
import logging

from glob import glob
from types import SimpleNamespace

ENG_DIR = r"\\hssieng\DATA\HS\JOBS"
part_re = re.compile(r'[a-zA-Z][0-9]+[a-zA-Z]+')
part_mark_pattern = re.compile()(r'[a-zA-Z][0-9]+[a-zA-Z]+')
# TODO: Cross reference JobStandards with rest of data?

# cSpell:disable
skip = SimpleNamespace()
skip.books = ['Products', 'JobStandards', 'ShipPieces']
skip.sheets = ['Special', 'Template', 'Total', 'L Weights']
skip.types = [
    'Anch Bolt',
    'Anch Wash',
    'Bronze Pl',
    'Elast Brg',
    'Fab Pad',
    'Grating',
    'HS Bolt',
    'HSS',
    'Nut',
    'RB',
    'Rebar',
    'Screw',
    'Std Wash',
    'Stud',
    'Valid Comms.',
    'DTI Wash',
    'STD. WASH',
    'HS BOLT',
    'NUT',
]
# cSpell:enable

bom_header = SimpleNamespace()
bom_header.QTY = 'QTY EA.'
bom_header.MARK = 'MARK'
bom_header.TYPE = 'COMM'
bom_header.THK = 'DESCRIPTION'
bom_header.THK_TO_WID_OFFSET = 2
bom_header.LEN = 'LENGTH'
bom_header.LEN_INCHES_OFFSET = 1
bom_header.SPEC = 'SPEC'
bom_header.GRADE = 'GRADE'
bom_header.TEST = 'TEST'
bom_header.REMARKS = 'REMARKS'
bom_header.ITEM = 'ITEM'
bom_header.WT = 'SHIP WT. EA. (LBS|KG)'
bom_header.WT_startswith = 'SHIP WT. EA.'

eng_data_maps = SimpleNamespace()
eng_data_maps.THICKNESS = {
    '10 GA.': 0.125,
    '16 GA.': 0.062,
}
eng_data_maps.GRADE = {  # (spec, grade, test)
    'A606 Type 4': ('A606', 'TYPE4', 'N/A'),
}

force_cvn_mode = False
all_parts = dict()


class BomData:

    def __init__(self, job, shipment=None, requires_full_load=False, force_cvn=False):
        self.job = job
        self.shipment = shipment

        self.part_data = dict()

        self.fetched_job_standards = False
        self.fetched_full_bom = False

        if force_cvn:
            force_cvn_mode = True

        if requires_full_load:
            self.load_bom()

    def get_part_data(self):
        pass

    def load_bom(self):
        # fetch full bom

        self.fetched_full_bom = True

    def load_job_standards(self):
        # fetch job standards data

        self.fetched_job_standards = True


class Part:

    def __init__(self, *args, **kwargs):
        self.assemblies = list()    # assemblies that part occurs on

        for attr, value in kwargs.items():
            setattr(self, attr, value)

        if type(self.thk) is str:
            self.thk = eng_data_maps.THICKNESS[self.thk]

        if 'HPS' in self.grade:
            self.zone = 3
        else:
            self.zone = 2

    def __repr__(self):
        return "[{qty}]|{part} ({grade})".format(qty=self.qty, part=self.name, grade=self.material_grade)

    @property
    def material_grade(self):
        if type(self.grade) in (float, int):     # float -> int  i.e. 50 -> '50'
            self.grade = int(self.grade)

        grade = "{spec}-{grade}".format(spec=self.spec, grade=self.grade)
        if self.test == 'N/A':
            return grade
        elif self.test:
            return grade + "{test}{zone}".format(test=self.test[0], zone=self.zone)
        elif force_cvn_mode:
            return grade + "T2"
        return grade

    @property
    def qty(self):
        total = 0
        for assembly, qty in self.assemblies:
            total += assembly.qty * qty

        return total

    def add_assembly(self, assembly, part_qty):
        self.assemblies.append((assembly, part_qty))

        return self


class Assembly:

    def __init__(self, *args, **kwargs):
        self.parts = list()
        self.mark = kwargs['name']
        self.qty = kwargs['qty']

    def add_part(self, part, qty):
        part.add_assembly(self, qty)
        self.parts.append(part)


def parse_header(header):
    header_mapping = SimpleNamespace()

    header_mapping.QTY = header.index(bom_header.QTY)
    header_mapping.MARK = header.index(bom_header.MARK)
    header_mapping.TYPE = header.index(bom_header.TYPE)
    header_mapping.THK = header.index(bom_header.THK)
    header_mapping.WID = header.index(bom_header.WID)
    header_mapping.LEN = header.index(bom_header.LEN)
    header_mapping.SPEC = header.index(bom_header.SPEC)
    header_mapping.GRADE = header.index(bom_header.GRADE)
    header_mapping.TEST = header.index(bom_header.TEST)
    header_mapping.REMARKS = header.index(bom_header.REMARKS)
    header_mapping.ITEM = header.index(bom_header.ITEM)
    header_mapping.QTY = header.index(bom_header.QTY)

    for column in header:
        if column.startswith(bom_header.WT_startswith):
            if "LBS" in column:
                header_mapping.units = "IMPERIAL"
            elif "KG" in column:
                header_mapping.units = "METRIC"
            break

    return header_mapping


def parse_line(header, line):
    qty = line[header.QTY]
    mark = line[header.MARK]
    type = line[header.TYPE]

    if mark is None or type in skip.types:
        return None

    if type is None:    # ASSEMBLY

        return Assembly(name=mark, qty=qty)

    # ~~~~~~~~~~~~~~~~~~ PART (not previously parsed/created) ~~~~~~~~~~~~~~~~~~~~~~~
    if mark not in all_parts.keys():
        thk = line[header.THK]
        width = line[header.THK + header.THK_TO_WID_OFFSET]

        # length
        if header.units == 'IMPERIAL':
            length = line[header.LEN] * 12 + \
                line[header.LEN + header.LEN_INCHES_OFFSET]
        else:
            length = line[header.LEN]

        spec = line[header.SPEC]
        if spec in eng_data_maps.GRADE.keys():
            spec, grade, test = eng_data_maps.GRADE[spec]
        else:
            grade = line[header.GRADE]
            test = line[header.TEST]
        # data.remarks = line[header.REMARKS]
        item = line[header.ITEM]
        all_parts[mark] = Part(mark=mark, thk=thk, wid=width,
                               len=length, spec=spec, grade=grade, test=test, item=item)

    return all_parts[mark]


def extract_sheet_data(sheet):
    end = sheet.range('Print_Area').last_cell.row

    # part and assembly queues
    assemblies = list()

    previous_line_type = 'ASSEMBLY'
    header = parse_header(sheet.range("B2:AB2").value)
    for row in sheet.range("B4:AB{}".format(end)).value:
        line = parse_line(header, row)

        if type(line) is Assembly:
            if previous_line_type == 'PART':
                assemblies = list()
                previous_line_type = 'ASSEMBLY'
            assemblies.append(line)
        elif type(line) is Part:
            is_main_component = (not part_mark_pattern.match(line.mark))
            for assembly in assemblies:
                if is_main_component:   # add to all parts as parent-component
                    proper_mark = "{parent}-{component}".format(
                        parent=assembly.mark, component=line.mark)
                    all_parts[proper_mark] = line
                assembly.add_part(line, row[header.QTY])

            if is_main_component:
                del all_parts[line.mark]
            previous_line_type = 'PART'


def get_bom_files():
    # find path
    base_folder = os.path.join(
        config['Paths']['EngJobs'], config['Instance']['Job'])
    for folder in glob(base_folder + '-*'):
        folder_shipment = folder.split('-')[1]
        if folder_shipment.isalpha():
            continue
        if config['Instance']['Shipment'] in folder_shipment:
            bom_folder = os.path.join(folder, 'BOM')
            break
    else:
        bom_folder = os.path.join(base_folder, 'BOM')
    logging.info("Using engineering BOM folder: {}".format(bom_folder))

    # workbook list generator
    for entry in os.scandir(bom_folder):
        if entry.is_file() and entry.name.split('.')[0] not in skip.books:
            yield entry.path


def get_bom_data(job=None, shipment=None):
    if job:
        config['Instance']['Job'] = job
        config['Instance']['Year'] = '20' + job[1:3]
    if shipment:
        config['Instance']['Shipment'] = shipment

    xl_app = xlwings.App()
    for bom_file in get_bom_files():
        wb = xl_app.books.open(bom_file)

        for sheet in wb.sheets:
            if sheet.name not in skip.sheets:
                try:
                    extract_sheet_data()
                except:
                    logging.exception(
                        "Error at workbook {} sheet {}".format(wb.name, sheet.name))

        wb.close()
    xl_app.quit()


def get_part_data():
    get_bom_data()

    return all_parts
