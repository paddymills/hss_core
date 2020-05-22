#!/usr/bin/env python

import pyodbc
import sys
import datetime as dt

from collections import defaultdict
from itertools import zip_longest, islice


def formatDateTime(x): return dt.datetime.strftime(x, '%m/%d/%Y %H:%M')


cs = "DRIVER={SQL Server};SERVER=HIIWINBL18;UID=SNUser;PWD=BestNest1445;"
conn = pyodbc.connect(cs)
cur = conn.cursor()

lastProgram = None


def check_status(prog, cursor=None):
    cur = cursor or conn.cursor()
    cur.execute("""
        SELECT
            Comp.CompletedDateTime, Comp.OperatorName,
            Stock.HeatNumber, Stock.BinNumber, Stock.SheetName
        FROM SNDbase91.dbo.StockHistory as Stock
            INNER JOIN OYSProgramUpdate.dbo.CompletedProgram AS Comp
                ON Comp.ProgramName = Stock.ProgramName
                AND Comp.SheetName = Stock.SheetName
        WHERE Stock.ProgramName=?
    """, prog)
    try:
        record = list(cur.fetchone())
        record[0] = formatDateTime(record[0])
        return ['Updated', ' :: '.join([str(x) for x in record])]
    except TypeError:
        cur.execute("""
            SELECT
                ArcDateTime, TransType, AutoID, SheetName
            FROM SNDbase91.dbo.ProgArchive as ProgArchive
            WHERE ProgArchive.ProgramName=?
            ORDER BY ArcDateTime DESC, TransType ASC
        """, prog)
        try:
            record1 = cur.fetchone()
            record2 = cur.fetchone()
            if record2:
                lessThan10 = abs(record1[0] - record2[0]
                                 ) < dt.timedelta(seconds=10)
            else:
                lessThan10 = False

            if record1[1] == 'SN100' or lessThan10:
                # TODO: Add file finder (in post or backup folders)
                return ['Still Active']
            elif record1[1] == 'SN101':
                if record1[3][:4] == "SLAB":
                    return ['Slab Nest Deleted', formatDateTime(record1[0])]
                else:
                    return ['Deleted', formatDateTime(record1[0])]
            elif record1[1] == 'SN102':
                cur.execute("""
                    SELECT
                        Prog.ArcDateTime, 'SN Update',
                        Stock.HeatNumber, Stock.BinNumber, Stock.SheetName
                    FROM SNDbase91.dbo.StockHistory as Stock
                        INNER JOIN SNDBase91.dbo.ProgArchive AS Prog
                            ON Prog.ProgramName = Stock.ProgramName
                            AND Prog.SheetName = Stock.SheetName
                    WHERE Stock.ProgramName=?
                """, prog)
                record = list(cur.fetchone())
                record[0] = formatDateTime(record[0])
                return ['Updated', ' :: '.join([str(x) for x in record])]
        except TypeError:
            return ['Program does not exist, possible input error']


def recent_updates(week=False, elseDays=1):
    if week:
        days = 7
    elif dt.date.today().weekday() == 0:
        days = 3
    else:
        days = elseDays
    start = dt.date.today() - dt.timedelta(days=days)

    cur.execute("""
        SELECT DISTINCT ProgramName
        FROM SNDbase91.dbo.ProgArchive
        WHERE TransType='SN102' AND ArcDateTime > ? AND
            MachineName IN ('Gemini', 'MG_OXY_GLOBAL', 'MG_TITAN_GLOBAL')
        ORDER BY ProgramName
    """, dt.datetime.combine(start, dt.datetime.min.time()))
    data = [x[0] for x in cur.fetchall()]
    display_many(data)


def main_updates(days=0):
    if days:
        pass
    elif dt.date.today().weekday() == 0:
        days = 3
    else:
        days = 1
    start = dt.date.today() - dt.timedelta(days=days)

    cur.execute("""
        SELECT DISTINCT ProgramName
        FROM SNDbase91.dbo.ProgArchive
        WHERE TransType='SN102' AND ArcDateTime > ? AND
            MachineName NOT IN ('Gemini', 'MG_OXY_GLOBAL', 'MG_TITAN_GLOBAL')
        ORDER BY ProgramName
    """, dt.datetime.combine(start, dt.datetime.min.time()))
    data = [x[0] for x in cur.fetchall()]
    display_many(data)


def pl3_updates():
    start = dt.date(2019, 1, 1)

    cur.execute("""
        SELECT DISTINCT ProgramName
        FROM SNDbase91.dbo.ProgArchive
        WHERE TransType='SN102' AND ArcDateTime > ? AND
            MachineName LIKE 'Plant_3_%'
        ORDER BY ProgramName
    """, dt.datetime.combine(start, dt.datetime.min.time()))
    data = [x[0] for x in cur.fetchall()]
    display_many(data)


def display_many(data):
    MAX_ITEMS_IN_COLUMN = 10
    for prefixLength in range(2, 5):  # 2-4 characters
        prefixes = set(map(lambda x: x[:prefixLength], data))
        try:
            table = dict()
            for pfx in prefixes:
                prefixItems = list(filter(lambda x: x.startswith(pfx), data))
                if len(prefixItems) > MAX_ITEMS_IN_COLUMN:
                    raise Exception()  # next prefixLength
                table[pfx] = prefixItems
            else:
                break
        except Exception:
            continue

    rows = list()
    rowPrefix = "NOT_INITIALZIED"
    for key in sorted(table.keys()):
        if not key.startswith(rowPrefix):
            # add row and update rowPrefix
            rows.append(list())
            rowPrefix = key[:-1]

        rows[-1].append(table[key])

    MAX_COLUMNS = 8
    consolidatedRows = list()
    tempRow = list()
    for row in rows:
        if len(tempRow) + len(row) > MAX_COLUMNS:
            consolidatedRows.append(tempRow)
            tempRow = list()

        tempRow.extend(row)
        tempRow.append([])
    else:
        if tempRow:
            consolidatedRows.append(tempRow)

    print("\n", "=" * 12 * MAX_COLUMNS, "\n")  # visual row separator
    for values in consolidatedRows:
        for x in zip_longest(*values, fillvalue=''):
            print(" ", ".   ".join([a.ljust(8) for a in x]))
        print("\n", "=" * 12 * MAX_COLUMNS, "\n")  # visual row separator


def input_handler(program):
    global lastProgram
    spaces = 0
    if lastProgram and len(program) <= 5:
        spaces = len(program)

        start = 5 - len(program.split('-')[0])
        program = lastProgram[:start] + program
    lastProgram = program
    print(program, *check_status(program))


if __name__ == '__main__':
    if sys.argv[1:]:
        for x in sys.argv[1:]:
            if x == "r":
                recent_updates(elseDays=2)
            elif x == "w":
                recent_updates(week=True)
            elif x == "m":  # Main :: Default(1 day)
                main_updates()
            elif x == "mw":  # Main :: Week
                main_updates(days=7)
            elif x == "mw+":  # Main :: 2 Weeks
                main_updates(days=14)
            elif x == "mm":  # Main :: 1 Month
                main_updates(days=30)
            elif x == "mr":  # Main :: Recent(3 days)
                main_updates(days=3)
            elif x == "mm":  # Main :: Month
                main_updates(days=30)
            elif x == "p3":  # Main Plant 3 :: Editable
                pl3_updates()
            else:
                input_handler(x)
    else:
        while 1:
            in_str = input("Program: ")
            if not in_str:
                break
            input_handler(in_str)
    conn.close()
