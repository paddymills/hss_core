#!/usr/bin/env python

import argparse
import pyodbc
import re
from datetime import datetime

from prodctrlcore.io.db import get_sndb_conn

sndb_conn = get_sndb_conn()
sndb = sndb_conn.cursor()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('program', nargs='?', default='',
                        help='script args: UH or US')

    args = parser.parse_args()
    arg = args.program.upper()

    opts = [
        ['Update Program Heat Number', 'UH', update_heat, '\nProgram: '],
        ['Update Sheet Size', 'US', update_size, '\nSheet Name: '],
        ['Change Burned Part Name', 'UPIPArc',
            update_partname, '\nOriginal Part Name: '],
    ]

    if arg:
        for x in opts:
            if x[1] == arg:
                script, val_str = x[2:]
                break
        else:
            print(get_updated(args.program))
            exit()
    else:
        print('Available Archive Scripts:')
        indexed = enumerate([x[0] for x in opts], start=1)
        print(''.join(['\n  %s::%s' % (i, v) for i, v in indexed]))
        index = int(input('\nScript to run: '))
        if index > len(opts):
            print('Index out of range')
            exit()
        script, val_str = opts[index - 1][2:]

    print('\n\n')
    while 1:
        val = input(val_str).upper()
        if not val:
            break
        ret = script(val)
        if ret:
            print(ret + '\n')

    sndb_conn.close()


# arg :: UH >> update heat number, PO number and SAP MM if given
def update_heat(prog, heat=None, po=None, mm=None):
    sndb.execute("""
        SELECT HeatNumber, BinNumber, PrimeCode
        FROM StockArchive
        WHERE ProgramName=?
        ORDER BY ArcDateTime DESC
    """, prog)
    orig_heat, orig_po, orig_mm = sndb.fetchone()
    heat = heat or input('Heat Number: ').upper() or orig_heat
    po = po or input('PO Number: ').upper() or orig_po
    mm = mm or input('SAP MM: ').upper() or orig_mm

    val = input(f'Heat :: {heat}\nPO :: {po}\nSAP MM :: {mm}\n\nCommit? ')
    if not val or val.upper()[0] != 'Y':
        return None

    # update heat, po and sap mm
    sndb.execute("""UPDATE StockHistory
                 SET HeatNumber=?, BinNumber=?, PrimeCode=?
                 WHERE ProgramName=?""", (heat, po, mm, prog))
    sndb.execute("""UPDATE StockArchive
                 SET HeatNumber=?, BinNumber=?, PrimeCode=?
                 WHERE ProgramName=?""", (heat, po, mm, prog))
    sndb_conn.commit()

    return None


# arg :: US >> update sheet size
def update_size(sheet, wid=None, len=None):
    sndb.execute('SELECT Width, Length FROM Stock WHERE SheetName=?', sheet)
    db_wid, db_len = sndb.fetchone()
    wid = wid or input('Width: ').strip() or db_wid
    len = len or input('Length: ').strip() or db_len
    area = float(len) * float(wid)

    sndb.execute('''
    UPDATE Stock
    SET Width=?, Length=?, Area=?
    WHERE SheetName=?
    ''', (wid, len, area, sheet))
    sndb_conn.commit()

    return None


def update_partname(oldPart, newPart=None):
    newPart = newPart or input('New Part Name: ').strip()

    sndb.execute('''
        UPDATE PIPArchive
        SET PartName=?
        WHERE PartName=?
    ''', (newPart, oldPart))

    val = input(f'Increase work order quantity of {oldPart}? ')
    if val and val.upper()[0] != 'Y':
        sndb.execute('SELECT QtyOrdered FROM Part WHERE PartName=?', oldPart)
        qty = sndb.fetchone()[0]
        sndb.execute('''
            UPDATE Part
            SET QtyOrdered=?
            WHERE PartName=?
        ''', (qty + 1, oldPart))

    sndb_conn.commit()
    print('Update complete')
    return None


if __name__ == '__main__':
    main()
