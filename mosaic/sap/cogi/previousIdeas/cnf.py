#!/usr/bin/env python3

import pyodbc
import csv
import datetime
import os

from operator import attrgetter

import records

CS = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445;"

def main():
    # runFromPreDefinedPartsList()
    runFromOrdersList()


def runFromOrdersList():
    partsList = getPartsListFromCSV()

    refinedPartsList = list()
    for part in partsList:
        if not part.wbs.startswith("D-"):
            continue

        if checkIfWorkOrderCreated(part.name, part.shipment):
            refinedPartsList.append(part)

    partNames = [part.name for part in refinedPartsList]
    activeSNParts = checkIfCompleteInSigmaNEST(partNames)
    sapData = getSAPData(partNames)

    sortKey = lambda x: (x.name, [-ord(c) for c in x.wbs])

    outputData = list()
    for part in sorted(refinedPartsList, key=sortKey):
        if activeSNParts[part.name] > 0:
            lowerQty = min(activeSNParts[part.name], part.qty)
            part.qty -= lowerQty
            activeSNParts[part.name] -= lowerQty

        if part.qty == 0:
            continue

        try:
            sapData[part.name][2] = part.wbs
            sapData[part.name][4] = str(int(part.qty))
            outputData.append("\t".join(sapData[part.name]))
        except KeyError:
            print(part, "not in SAP Processed Files")

    createOutputFile(outputData)


def runFromPreDefinedPartsList():
    partsList = embeddedPartsList()
    sapData = getSAPData(partsList)
    activeSNParts = checkIfCompleteInSigmaNEST(partsList)
    sapOrders = getPartsListFromCSV()

    outputData = []
    for part in partsList:
        orders = [x for x in sapOrders if x.part == part]

        if not orders:
            print(part, ":: No orders found")

        for o in sorted(orders, key=lambda x: x.wbs, reverse=True):
            if activeSNParts[part] > 0:
                lowerQty = min(activeSNParts[part], o.qty)
                o.qty -= lowerQty
                activeSNParts[part] -= lowerQty

            if o.qty == 0:
                continue

            try:
                sapData[part][2] = o.wbs
                sapData[part][4] = str(int(o.qty))
                outputData.append("\t".join(sapData[part]))
            except KeyError:
                print(part, "not in SAP Processed Files")

    createOutputFile(outputData)


def parse_csv(file_name, record_class, constrain_list=None):
    ls = list()

    with open(file_name, "r") as csvfile:
        for row in csv.DictReader(csvfile):
            _temp = record_class(row)
            if constrain_list is None or _temp.part in constrain_list:
                ls.append(_temp)
    return ls


def getPartsListFromCSV():
    return parse_csv("orders.csv", records.OrderRecord)


def embeddedPartsList():
    job = "1170155E"
    parts = """X703K"""

    return [f"{job}-{part}" for part in sorted(parts.split("\n"))]


def checkIfWorkOrderCreated(part, shipment):
    SN_WO = r"\\HSSIENG\DATA\HS\SAP - Material Master_BOM\SigmaNest Work Orders"
    wo_year = lambda x: os.path.join(SN_WO, f"20{x[1:3]} Work Orders Created")

    js = f"{part[:8]}-{int(shipment)}"
    if os.path.exists(os.path.join(wo_year(js), f"{js}_SimTrans_WO.xls")):
        return True
    return False


def checkIfCompleteInSigmaNEST(parts):
    active = dict()
    conn = pyodbc.connect(CS)
    cursor = conn.cursor()

    for part in parts:
        active[part] = 0
        snPart = part.replace("-", "_", 1)
        cursor.execute("""
            SELECT QtyOrdered, QtyCompleted
            FROM Part
            WHERE PartName = ?
        """, snPart)
        for ordered, complete in cursor.fetchall():
            active[part] += ordered - complete

    conn.close()
    return active


def getSAPData(parts):
    processed_dir = r"\\hssieng\SNData\SimTrans\SAP Data Files\Processed"
    prod_data = dict()

    for f in os.listdir(processed_dir):
        with open(os.path.join(processed_dir, f), "r") as prod_file:
            for line in prod_file.readlines():
                if line.split("\t")[0] in parts:
                    prod_data[line.split("\t")[0]] = line.split("\t")

    return prod_data


def createOutputFile(data):
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".ready"
    with open(r"Production_" + timestamp, "w") as cnf_file:
        for line in data:
            if type(line) is list:
                cnf_file.write("\t".join(line))
            else:
                cnf_file.write(line)

if __name__ == '__main__':
    main()
