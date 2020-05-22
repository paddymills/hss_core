
from os.path import join, dirname

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
