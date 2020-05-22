
from collections import defaultdict
import csv
import xlwings
import pyodbc
import re
import datetime
import os

import records

RAW_MM_RE = re.compile(r"[0-9]{7}[A-Z][0-9]{,2}-[0-9]{5}[A-Z]?")
STOCK_MM_RE = re.compile(r"50/50W-[0-9]{4}")
ERR_RE1 = re.compile(r"Deficit of Project stock unr. (\d+\.\d+) ([A-Z]+2)")
ERR_RE2 = re.compile(r"Deficit of SL Unrestricted-use (\d+\.\d+) ([A-Z]+2)")

CS = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;"

def parse_csv(file_name, record_class):
    ls = list()

    with open(file_name, "r") as csvfile:
        for row in csv.DictReader(csvfile):
            _temp = record_class(row)
            ls.append(_temp)

    return ls


def get_active():
    PL_TEMP = "([A-Z]+[0-9]+)?"
    SLAB_PART_RE = re.compile(r"([0-9]{7}[A-Z]-[A-Z]+[0-9]+[A-Z]*-)" + PL_TEMP * 20)
    SLAB_WO_RE = re.compile(r"[0-9]{7}[A-Z]-[0-9]+-SLAB")
    conn = pyodbc.connect(CS)
    cursor = conn.cursor()

    cursor.execute("SELECT WONumber, ProgramName, PartName FROM PIP")
    d = dict()
    for wo, prog, part in cursor.fetchall():
        part = part.replace("_", "-", 1)
        if SLAB_WO_RE.match(wo):
            try:
                match = SLAB_PART_RE.match(part).groups()
                for x in match[1:]:
                    if x:
                        d[match[0] + x] = prog
            except AttributeError:
                pass
        else:
            if part in d.keys():
                continue
            d[part] = prog

    conn.close()
    return d


def main():
    cogi = parse_csv("cogi.csv", records.CogiRecord)
    prod = parse_csv("prod.csv", records.StockRecord)
    orders = parse_csv("orders.csv", records.OrderRecord)

    active = get_active()
    date = datetime.date.today().strftime("%m/%d/%Y")

    prog_update = list()
    sto = list()
    winshuttle = list()
    migo_tr = list()
    cnf = list()

    def cnf_add(part, qty, plant, wbs):
        for i, x in enumerate(cnf):
            if x[0] == part and x[2] == plant and x[3] == wbs:
                cnf[i][1] += qty
                break
        else:
            cnf.append([part, qty, plant, wbs])

    for x in cogi:
        if x.part != "1170082A-X201A":
            continue

        _prod = [p for p in prod if p == x]
        _orders = [o for o in orders if o == x]

        print(x.part, x.plant, x.loc, x.wbs)
        for p in _prod:
            print("PROD ::", p.qty, p.wbs)
        for o in _orders:
            print("ORDER ::", o.qty, o.wbs)


if __name__ == '__main__':
    main()
