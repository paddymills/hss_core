
from argparse import ArgumentParser

parser = ArgumentParser()
    
parser.add_argument('-d', '--dev',              action='store_true',        help='Execute on development instance')
parser.add_argument('-r', '--restore',          action='store_true',        help='Execute restore mode')
parser.add_argument('-f', '--file',             action='store',             help='File to restore(restore flag only)')
parser.add_argument('-j', '--jobs', '--job',    action='extend', nargs='*', help='Jobs to update/restore', default=[])

parser.print_help()
print(parser.parse_args('--jobs 1170083 --job 1170082 1180301'.split()))
print(parser.parse_args('-r'.split()))