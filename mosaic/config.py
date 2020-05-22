#/bin/python3

import os
# YAML library
from yaml import load
try:
    from yaml import CLoader as Loader
except ImportError:
    from yaml import Loader

with open('config.yaml') as cfg_stream:
    cfg = load(cfg_stream, Loader=Loader)

for key, val in cfg['files'].items():
    if type(val) is list:
        cfg['files'][key] = os.path.join(*val)
