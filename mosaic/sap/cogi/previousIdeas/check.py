#!/usr/bin/env python3

from collections import defaultdict
import pyodbc
import xlwings

COGI_HEADER = ["Material", "WBS Element", "Qty in unit of entry", "Unit of Entry", "Order", "Plant"]
ORDERS_HEADER = ["Material Number", "WBS Element", "Order quantity (GMEIN)"]

def parse_csv(ls, header):
    template = [None] * len(header)
    for i, x in enumerate(ls[0].split(",")):
        if x in header:
            template[header.index(x)] = i

    for line in ls[1:-1]:
        yield [line.split(",")[x] for x in template]


cogi = defaultdict(int)
orders = defaultdict(int)

with open("cogi.csv", "r") as f:
    for x in parse_csv(f.read().split('\n'), COGI_HEADER):
        if x[3] == "EA":
            if x[5] == "HS02" and x[0].count("-") < 2:
                continue
            cogi[(x[0], x[1], x[4])] += int(x[2])

with open("orders.csv", "r") as f:
    for x in parse_csv(f.read().split('\n'), ORDERS_HEADER):
        orders[(x[0], x[1])] += int(x[2])

not_found = [["Material", "WBS Element", "Consuming Order", "Qty"]]
for k, v in cogi.items():
    if k[1][0] == "S":
        # print(k)
        pass
    elif k[:2] not in orders.keys():
        not_found.append([*k, v])
    elif v < orders[k[:2]]:
        not_found.append([*k, v])

create = defaultdict(int)
for x in not_found[1:]:
    create[x[0]] += x[3]

bom, no_bom = [], []
migo = []
for x in not_found[1:]:
    if x[0].count("-") > 1:
        bom.append([x[0], "hs01", "pp01", x[3], "11/12/2018", x[1]])
    else:
        no_bom.append([x[0], "hs01", "pp01", x[3], "11/12/2018", "1170062B-X1A", x[1]])
    migo.append(["415", None, x[0], None, None, "HS01", "PROD", x[3], None, None, None, None, None, None, x[0], "HS01", "PROD", x[1]])

# wb = xlwings.Book("CO01 BOM Exists.xlsx")
# wb.sheets[0].range("A2").value = bom
# wb = xlwings.Book("CO01 without a BOM.xlsx")
# wb.sheets[0].range("A2").value = no_bom

wb = xlwings.books.active.sheets.active
wb.range("A2").value = migo

# cs = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91"
# conn = pyodbc.connect("DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91")
# cur = conn.cursor()
# with open("create.csv", "w") as output:
#     for k, v in create.items():
#         cur.execute(
#             "SELECT ProgramName FROM PIP WHERE PartName=?",
#             k.replace("-", "_", 1)
#         )
#         if not cur.fetchall():
#             output.write(f"{k},{v}\n")
# conn.close()
