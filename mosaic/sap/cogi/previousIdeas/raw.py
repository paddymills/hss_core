
from collections import defaultdict
import csv
import xlwings
import re

RAW_MM_RE = re.compile(r"[0-9]{7}[A-Z][0-9]{,2}-[0-9]{5}[A-Z]?")
STOCK_MM_RE = re.compile(r"50/50W-[0-9]{4}")

COGI_HEADER = ["Material", "Storage Location", "WBS Element", "Qty in unit of entry", "Unit of Entry", "Plant"]
MB52_HEADER = ["Material Number", "Storage Location", "Special stock number", "Unrestricted", "Base Unit of Measure", "Plant"]

def parse_csv(file_name, header):
    with open(file_name, "r") as f:
        ls = [x for x in csv.reader(f)]
        template = [None] * len(header)
        for i, x in enumerate(ls[0]):
            if x in header:
                template[header.index(x)] = i

        for line in ls[1:-1]:
            yield [line[x] for x in template]

cogi = defaultdict(float)
mb52 = defaultdict(float)

sap_quantifier = lambda x: "_".join([x[a] for a in (0, 1, 2, 5)])
for x in parse_csv("raw_cogi.csv", COGI_HEADER):
    cogi[sap_quantifier(x)] += float(x[3].replace(",", ""))
for x in parse_csv("raw_mat.csv", MB52_HEADER):
    qty = float(x[3].replace(",", ""))
    if x[4] != "IN2":
        if x[4] == "FT2":
            qty *= 144
        elif x[4] == "M2":
            qty *= 1550.003
    mb52[sap_quantifier(x)] += qty

gi = []
for k, v in cogi.items():
    if k in mb52.keys():
        if mb52[k] - v <= 0:
            continue
        qty = mb52[k] - v
    else:
        qty = v
    mm, loc, wbs, plant = k.split("_")
    gi.append(["222", "Q", mm, None, plant, loc, qty, None, None, "IN2", None, wbs])

wb = xlwings.Book()
wb.sheets[0].range("A1").value = gi
