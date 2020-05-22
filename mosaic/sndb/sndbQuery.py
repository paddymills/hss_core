#!/usr/bin/env python

import pyodbc
import re
import sys

from datetime import datetime

import cli_stream
import updatedPrograms


def formatDateTime(x): return datetime.strftime(x, '%m/%d/%Y %H:%M')


REPLACEMENTS = ['*', '#', '+']

cs = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445;"
conn = pyodbc.connect(cs)
cur = conn.cursor()


def size(thkWidLen):
    thk, wid, length = thkWidLen
    if int(thk) == thk:
        thk = int(thk)
    return ' x '.join([str(thk), str(round(wid, 3)), str(round(length, 3))])


def recursive_replace(_str, find_vals, replace_val):
    return_val = _str.replace(find_vals.pop(0), replace_val)
    if find_vals:
        return recursive_replace(return_val, find_vals, replace_val)
    return return_val


def input_handler(val):
    patterns = [
        (sheet, '^[A-Za-z]{1,2}[0-9]+$'),
        (part, '^[0-9]+[A-Za-z]?([-_][0-9A-Za-z#+]+)+$'),
        (part, '^[0-9]{7}[A-Za-z]?$'),
        (program, '^[0-9]{5}([-][0-9]+)?$'),
        (material_master, '^[0-9]+[A-Za-z][0-9]{0,2}-[0-9]{2,5}[A-Za-z]?$'),
        (material_master, '^[50wW/]+-[0-9]{4,}[A-Za-z]?$'),
    ]
    for func, pattern in patterns:
        if re.match(pattern, val):
            for x in func(recursive_replace(val, REPLACEMENTS.copy(), '%').upper() + '%'):
                if x[0] == datetime(1900, 1, 1):
                    x[0] = '[....Active....]'
                elif type(x[0]) is datetime:
                    x[0] = datetime.strftime(x[0], '%m/%d/%Y %H:%M')
                cli_stream.inline_print('\n', ' :: '.join([str(a) for a in x]))
            cli_stream.inline_print('\n')


def sheet(sheet):
    cur.execute("""
        SELECT
            Program.ArcDateTime, Stock.SheetName, Program.ProgramName,
            Stock.HeatNumber, Stock.BinNumber,
            Stock.Thickness, Stock.Width, Stock.Length
        FROM StockHistory AS Stock
            INNER JOIN ProgArchive AS Program
                ON Program.SheetName=Stock.SheetName
                AND Program.ProgramName=Stock.ProgramName
        WHERE Stock.SheetName LIKE ? AND Program.TransType='SN102'
        UNION
        SELECT
            0, Stock.SheetName, Program.ProgramName,
            Stock.HeatNumber, Stock.BinNumber,
            Stock.Thickness, Stock.Width, Stock.Length
        FROM Stock
            LEFT JOIN Program
                ON Stock.SheetName=Program.SheetName
        WHERE Stock.SheetName LIKE ?
    """, [sheet] * 2)

    return [list(x[:-3]) + [size(x[-3:])] for x in cur.fetchall()]


def part(part):
    cur.execute("""
        SELECT
            PIP.ArcDateTime, PIP.PartName, PIP.ProgramName,
            Stock.HeatNumber, Stock.BinNumber, Stock.PrimeCode
        FROM PIPArchive AS PIP
            INNER JOIN StockArchive AS Stock
                ON PIP.ProgramName=Stock.ProgramName
        WHERE PIP.PartName LIKE ? AND PIP.TransType='SN102'
        UNION
        SELECT
            0, PIP.PartName, PIP.ProgramName,
            Stock.HeatNumber, Stock.BinNumber, Stock.PrimeCode
        FROM PIP
            INNER JOIN Program
                ON PIP.ProgramName=Program.ProgramName
            INNER JOIN Stock
                ON Program.SheetName=Stock.SheetName
        WHERE PIP.PartName LIKE ?
        ORDER BY PIP.ArcDateTime
    """, ['%' + part.replace('-', '%')] * 2)
    return list(cur.fetchall())


def program(prog):
    return [updatedPrograms.check_status(prog[:-1], cursor=cur)]


def material_master(sapmm):
    cur.execute("""
        SELECT
            Program.ArcDateTime, Stock.PrimeCode, Stock.SheetName,
            Program.ProgramName, Stock.HeatNumber, Stock.BinNumber
        FROM StockArchive AS Stock
            INNER JOIN ProgArchive AS Program
                ON Stock.SheetName=Program.SheetName
        WHERE Stock.PrimeCode LIKE ? AND Program.TransType='SN102'
        UNION
        SELECT
            0, Stock.PrimeCode, Stock.SheetName,
            Program.ProgramName, Stock.HeatNumber, Stock.BinNumber
        FROM Stock
            LEFT JOIN Program
                ON Stock.SheetName=Program.SheetName
        WHERE Stock.PrimeCode LIKE ?
    """, [sapmm] * 2)

    records = list(cur.fetchall())
    if len(records) > 10:
        records = sorted(
            records, key=lambda x: datetime.max if x[0] == datetime(1900, 1, 1) else x[0])[-20:]

    return records


if __name__ == '__main__':
    if len(sys.argv) > 1:
        for x in sys.argv[1:]:
            input_handler(x)
    else:
        cli_stream.IOLoop(input_handler, inputPrompt='Value: ')
    conn.close()
