
from os.path import join, dirname

# file handler imports
from mosaic.file_handlers.bom import *
from mosaic.file_handlers.flgdata import *
from mosaic.file_handlers.schedule import *
from mosaic.file_handlers.tagschedule import *
from mosaic.file_handlers.workorder import *


# YAML library
from yaml import load
try:
    from yaml import CLoader as Loader
except ImportError:
    from yaml import Loader

with open(join(dirname(__file__), 'config.yaml')) as config_stream:
    config = load(config_stream, Loader=Loader)

for key, val in config['files'].items():
    if type(val) is list:
        config['files'][key] = join(*val)
