import pyodbc
import xlwings
import re

CS = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445;"

wb = xlwings.books.active
sheet = wb.sheets.active
row = 2

with pyodbc.connect(CS) as db:
    cursor = db.cursor()
    while 1:
        part = sheet.range((row, 2)).value
        if not part:
            break

        cursor.execute("""
            SELECT ProgramName
            FROM PIP
            WHERE PartName=?
        """, part.replace("-", "_", 1))
        prog = cursor.fetchone()
        if prog:
            sheet.range((row, 3)).value = "ACTIVE"

        row += 1
