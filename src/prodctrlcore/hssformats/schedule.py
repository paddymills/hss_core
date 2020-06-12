
import os
import xlwings

from collections import defaultdict
from argparse import ArgumentParser

from prodctrlcore.utils import CountingIter

COST_CENTERS = {
    2005: "WB",
    2006: "NB",
    2007: "SB",
    2030: "MBN",
    2031: "MBS",
}


def get_job_ship_dates(xl_file, data_connection_name="High Steel Scheduling"):
    jobs = defaultdict(dict)
    wb = xlwings.Book(xl_file)

    print("Refreshing data connections")
    wb.api.Connections(data_connection_name).Refresh()

    # Early Start and Main Start dates
    data = wb.sheets['Dates'].range("DATES_HEADER").expand('down').value
    iter = CountingIter(data, "Reading Dates")
    next(iter)
    for job, early_start, main_start in iter:
        jobs[job]['early_start'] = early_start
        jobs[job]['main_start'] = main_start

    # Project Manager
    data = wb.sheets['PM'].range("PM_HEADER").expand('down').value
    iter = CountingIter(data, "Reading PM's")
    next(iter)
    for job, pm in iter:
        jobs[job]['pm'] = pm

    # Products/Types
    data = wb.sheets['Products'].range("PRODUCTS_HEADER").expand('down').value
    iter = CountingIter(data, "Reading Products")
    next(iter)
    prev_job, products = None, list()
    for job, product, date in iter:
        if job != prev_job:
            jobs[prev_job]['product'] = ','.join(products)
            prev_job, products = job, list()

        if product == 'S' and jobs[prev_job]['main_start'] is None:
            jobs[job]['main_start'] = date

        products.append(product)

    # Fab Bays
    data = wb.sheets['Bays'].range("BAYS_HEADER").expand('down').value
    iter = CountingIter(data, "Reading Bays")
    next(iter)
    prev_job, bays = None, list()
    for job, cc in iter:
        if job != prev_job:
            jobs[prev_job]['bay'] = ','.join(bays)
            prev_job, bays = job, list()

        bays.append(COST_CENTERS[int(cc)])  # cost center might be string

    del jobs[None]  # may have been created from products and bays

    wb.save()
    if len(wb.app.books) > 1:
        wb.close()
    else:
        wb.app.quit()

    return jobs
