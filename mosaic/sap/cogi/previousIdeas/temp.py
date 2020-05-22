
import xlwings
import re

RAW_MM_RE = re.compile(r"[0-9]{7}[A-Z][0-9]{,2}-[0-9]{5}[A-Z]?")
STOCK_MM_RE = re.compile(r"50/50W-[0-9]{4}")
ERR_RE1 = re.compile(r"Deficit of Project stock unr. (\d+\.\d+) ([A-Z]+2)")
ERR_RE2 = re.compile(r"Deficit of SL Unrestricted-use (\d+\.\d+) ([A-Z]+2)")

migo = dict()

wb = xlwings.books.active
data = wb.sheets[0].range("B2:N2").expand("down").value
for x in data:
    if not RAW_MM_RE.match(x[0]) and not STOCK_MM_RE.match(x[0]):
        continue

    mm = x[0]
    loc = x[3]
    plant = x[2]
    total = x[5]
    wbs = x[11]
    err = x[12].split(" : ")[0].replace(",", "")
    if ERR_RE1.match(err):
        qty, unit = ERR_RE1.match(err).groups()
    elif ERR_RE2.match(err):
        qty, unit = ERR_RE2.match(err).groups()
    else:
        print("did not match error text x" + err + "x")
        continue

    key = (mm, plant, loc, wbs)
    qty = float(qty)
    if unit != "IN2":
        if unit == "FT2":
            qty *= 144
        elif unit == "M2":
            qty *= 1550.003

    if key not in migo.keys():
        migo[key] = float(qty)
    else:
        migo[key] += total

dump = []
for k, v in migo.items():
    dump.append([k[0], None, k[1], k[2], v, None, None, "IN2", None, None, k[3]])

s = wb.sheets.add()
s.range("A1").value = dump
