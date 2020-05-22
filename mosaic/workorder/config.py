
import os

from configparser import ConfigParser, ExtendedInterpolation
from types import SimpleNamespace
from re import compile as regex

ROOT = os.path.dirname(__file__)
with open(os.path.join(ROOT, 'config.ini')) as file_stream:
    config = ConfigParser(interpolation=ExtendedInterpolation())
    config.read_file(file_stream)
    
# compute actual downloads folder path
config['Paths']['Downloads'] = os.path.expanduser(config['Paths']['Downloads'])

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

part_mark_pattern = regex(r'[a-zA-Z][0-9]+[a-zA-Z]+')