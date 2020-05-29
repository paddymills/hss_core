
import os
import xlwings

from collections import defaultdict
from argparse import ArgumentParser

COST_CENTERS = {
    2005: "WB",
    2006: "NB",
    2007: "SB",
    2030: "MBN",
    2031: "MBS",
}


def get_update_data(xl_file):
    wb = xlwings.Book(xl_file)
    jobs = defaultdict(dict)

    # Early Start and Main Start dates
    i = 2
    sheet = wb.sheets['Dates']
    while sheet.range(i, 1).value:
        print("\r[{}] Reading Dates".format(i-1), end='')

        job = sheet.range(i, 1).value
        jobs[job]['early_start'] = sheet.range(i, 2).value
        jobs[job]['main_start'] = sheet.range(i, 3).value

        i += 1
    print()

    # Project Manager
    i = 2
    sheet = wb.sheets['PM']
    while sheet.range(i, 1).value:
        print("\r[{}] Reading PM's".format(i-1), end='')

        job = sheet.range(i, 1).value
        jobs[job]['pm'] = sheet.range(i, 2).value

        i += 1
    print()

    # Products/Types
    i = 2
    sheet = wb.sheets['Products']
    job, products = None, list()
    while sheet.range(i, 2).value:
        print("\r[{}] Reading Products".format(i-1), end='')

        if sheet.range(i, 1).value:
            if job:  # not first row
                jobs[job]['product'] = ','.join(products)
            job = sheet.range(i, 1).value
            products = list()

        products.append(sheet.range(i, 2).value)
        i += 1
    print()

    # Fab Bays
    i = 2
    sheet = wb.sheets['Bays']
    job, bays = None, list()
    while sheet.range(i, 2).value:
        print("\r[{}] Reading Bays".format(i-1), end='')

        if sheet.range(i, 1).value:
            if job:  # not first row
                jobs[job]['bay'] = ','.join(bays)
            job = sheet.range(i, 1).value
            bays = list()

        cc = sheet.range(i, 2).value
        if type(cc) is str:
            cc = int(cc)
        bays.append(COST_CENTERS[cc])
        i += 1
    print()

    wb.save()
    if len(wb.app.books) > 1:
        wb.close()
    else:
        wb.app.quit()

    return jobs
