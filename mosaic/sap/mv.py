#!/usr/bin/env python

import pyodbc
import xlwings
from datetime import datetime

cnxn = 'DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445'

wb = xlwings.books.active
data = wb.sheets['snmove'].range('A2:C193').value

conn = pyodbc.connect(cnxn)
db = conn.cursor()


def update():
    for loc, _, sheet in data:
        db.execute("UPDATE Stock SET Location=? WHERE SheetName=?", loc, sheet)


def check():
    for loc, _, sheet in sorted(data):
        db.execute(
            "SELECT SheetName, Location FROM Stock WHERE SheetName=?", sheet)
        for x in db.fetchall():
            print(sheet, "::", x, "::", loc)
            if x[1] != loc:
                print("^^___________________ERROR________________^^")


if __name__ == "__main__":
    # update()
    check()

db.commit()
conn.close()
