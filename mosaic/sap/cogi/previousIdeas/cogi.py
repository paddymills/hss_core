
import csv
import xlwings
import pyodbc
import re
import datetime
import os

import numpy as np

from collections import defaultdict
from copy import copy

import records

RAW_MM_RE = re.compile(r"[0-9]{7}[A-Z][0-9]{,2}-[0-9]{5}[A-Z]?")
STOCK_MM_RE = re.compile(r"50/50W-[0-9]{4}")
ERR_RE1 = re.compile(r"Deficit of Project stock unr. (\d+\.\d+) ([A-Z]+2)")
ERR_RE2 = re.compile(r"Deficit of SL Unrestricted-use (\d+\.\d+) ([A-Z]+2)")

CS = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445;"

def parse_csv(file_name, record_class, constrain_list=None):
    ls = list()

    with open(file_name, "r") as csvfile:
        for row in csv.DictReader(csvfile):
            _temp = record_class(row)
            if constrain_list is None or _temp.part in constrain_list:
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


def update_wbs(plannedOrders):
    d = {}
    for x in plannedOrders:
        d[f"{x.part}-{x.wbs}"] = x.qty

    conn = pyodbc.connect(CS)
    cursor = conn.cursor()

    updates = []
    cursor.execute("""
        SELECT ID, PartName, WBS
        FROM SAPPartWBS
        WHERE QtyReq>QtyConf
    """)
    for id, part, wbs in cursor.fetchall():
        if f"{part}-{wbs}" in d.keys():
            updates.append([d[f"{part}-{wbs}"], id])

    print(updates)
    cursor.execute("UPDATE SAPPartWBS SET QtyConf=QtyReq")
    cursor.executemany("""
        UPDATE SAPPartWBS
        SET QtyConf = QtyReq - ?
        WHERE ID = ?
    """, updates)
    cursor.commit()
    conn.close()


def main():
    cogi = parse_csv("cogi.csv", records.CogiRecord)
    part_names = [x.part for x in cogi]

    prod = parse_csv("prod.csv", records.StockRecord, constrain_list=part_names)
    orders = parse_csv("orders.csv", records.OrderRecord, constrain_list=part_names)
    update_wbs(orders)

    active = get_active()
    date = datetime.date.today().strftime("%m/%d/%Y")

    prog_update = list()
    sto = list()
    winshuttle = list()
    migo_tr = list()
    cnf = list()

    def addToCnf(part, qty, plant, wbs):
        for i, x in enumerate(cnf):
            if x[0] == part and x[2] == plant and x[3] == wbs:
                cnf[i][1] += qty
                break
        else:
            cnf.append([part, qty, plant, wbs])

    def addToSto(part, qty):
        for i, x in enumerate(sto):
            if x[0] == part:
                sto[i][1] += qty
                break
        else:
            sto.append([part, qty])

    def transferQty(cogi, other):
        smaller_number = min(cogi.qty, other.qty)
        other.qty -= smaller_number
        cogi.qty -= smaller_number

        return smaller_number


    def migoTrFormat(part, qty, plant, fromWBS, toWBS):
        fmt = [None] * 18
        fmt[0] = part
        fmt[2] = plant
        fmt[3] = "PROD"
        fmt[4] = qty
        fmt[5] = "EA"
        fmt[13] = "PROD"
        fmt[14] = toWBS
        fmt[17] = fromWBS

        return fmt


    blacklist = (
        "1170066",
    )


    for c in cogi:
        if c.loc != "PROD" or c.part.startswith(blacklist):
            continue

        if c.part in active.keys():
            prog_update.append(active[c.part])
            continue

        for p in filter(lambda x: x==c, prod):
            if c.qty == 0:
                break
            movedQty = transferQty(c, p)
            if p.plant == "HS01" and c.plant == "HS02":
                addToSto(c.part, movedQty)

        for p in filter(c.matchOtherInPlant, prod):
            if c.qty == 0:
                break
            migo_tr.append(
                migoTrFormat(c.part, transferQty(c, p), c.plant, p.wbs, c.wbs)
            )

        for o in filter(lambda x: x==c, orders):
            if c.qty == 0:
                break
            addToCnf(c.part, transferQty(c, o), c.plant, c.wbs)

        if c.qty > 0:
            winshuttle.append([c.order, c.part, c.qty, c.wbs])

    uniqueFirstItems = lambda x: np.unique(np.take(winshuttle, [0], axis=1))

    cnf_parts = uniqueFirstItems(cnf)
    processed_dir = r"\\hssieng\SNData\SimTrans\SAP Data Files\Processed"
    deleted_dir = r"\\hssieng\SNData\SimTrans\SAP Data Files\deleted files"
    old_deleted_dir = r"\\hssieng\SNData\SimTrans\SAP Data Files\old deleted files"
    prod_data = dict()

    for f in os.listdir(processed_dir):
        with open(os.path.join(processed_dir, f), "r") as prod_file:
            for line in prod_file.readlines():
                if line.split("\t")[0] in cnf_parts:
                    prod_data[line.split("\t")[0]] = line.split("\t")

    for f in os.listdir(deleted_dir):
        with open(os.path.join(deleted_dir, f), "r") as prod_file:
            for line in prod_file.readlines():
                if line.split("\t")[0] in cnf_parts and line.split("\t")[0] not in prod_data.keys():
                    prod_data[line.split("\t")[0]] = line.split("\t")
    for f in os.listdir(old_deleted_dir):
        with open(os.path.join(old_deleted_dir, f), "r") as prod_file:
            for line in prod_file.readlines():
                if line.split("\t")[0] in cnf_parts and line.split("\t")[0] not in prod_data.keys():
                    prod_data[line.split("\t")[0]] = line.split("\t")

    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".ready"
    with open(r"Production_" + timestamp, "w") as cnf_file:
        for part, qty, plant, wbs in cnf:
            if part not in prod_data.keys():
                print(part, "not found &&", qty, plant, wbs)
                continue

            d = prod_data[part]
            area_ea = float(d[8]) / float(d[4])
            d[2] = wbs
            d[4] = str(int(qty))
            d[8] = str(round(area_ea * qty, 3))
            cnf_file.write("\t".join(d))


    wb = xlwings.Book()
    while len(wb.sheets) < 6:
        wb.sheets.add(after=wb.sheets[-1])
    wb.sheets[0].name = "STO"
    wb.sheets[1].name = "Confirm"
    wb.sheets[2].name = "Orders to Add Step"
    wb.sheets[3].name = "Add to Orders"
    wb.sheets[4].name = "Update Programs"
    wb.sheets[5].name = "MIGO_TR"
    wb.sheets[0].activate()

    wb.sheets[0].range("A2").value = sto
    wb.sheets[1].range("A2").value = cnf
    wb.sheets[2].range("A2").value = uniqueFirstItems(winshuttle).reshape((-1,1))
    wb.sheets[3].range("A2").value = winshuttle
    wb.sheets[5].range("A2").value = migo_tr

    progs = [[x] for x in set(prog_update)]
    all_parts = []
    for i, x in enumerate(progs):
        _parts = []
        for k, v in active.items():
            if v == x[0]:
                _parts.append(k)
                all_parts.append([k])
        progs[i].append(", ".join(_parts))
    wb.sheets[4].range("A2").value = sorted(progs, key=lambda x: x[1])
    wb.sheets[4].range("D2").value = sorted(all_parts)

    for s in wb.sheets:
        s.autofit()

if __name__ == '__main__':
    main()
