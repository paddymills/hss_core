
import os
import xlwings
import logging

from win32 import win32api
from functools import partial

import mosaic.bom as eng
from mosaic.config import config


def determine_processing():
    for app in xlwings.apps:  # active excel applications
        for book in app.books:
            if book.name.startswith(config['Names']['SSRS_ReportName']):
                book.activate()
                post_processing()
                return

    pre_processing()


def pre_processing():
    # 1) Open SSRS exported report
    # 2) Get engineering BOM data (material grades)
    # 3) Get other work order data, if available (material grades and operations)
    # 4) Check for Charge Ref number
    # 5) Save

    # get SSRS report
    wb = open_ssrs_report_file()
    assert wb is not None

    # fill out grade, remark, and operations
    fill_in_data(wb.sheets[0])

    # ensure charge ref number is valid
    if wb.sheets[0].range('O2').value is None:
        win32api.MessageBox(wb.app.hwnd, 'Charge Ref number needs entered.')
    wb.save()


def post_processing():
    # 1) Save pre-subtotalled document
    # 2) Fill out ProductionDemand spreadsheet
    # 3) Fill out CDS(TagSchedule)
    # 4) Import WBS-split data

    pass


def open_ssrs_report_file():
    last_modified = 0
    last_modified_path = None

    # scan through folder looking for most recent SSRS report file
    for dir_entry in os.scandir(config['Paths']['Downloads']):
        if dir_entry.name.startswith(config['Names']['SSRS_ReportName']):
            if dir_entry.stat().st_mtime > last_modified:
                last_modified = dir_entry.stat().st_mtime
                last_modified_path = dir_entry.path

    # no SSRS report file found
    if last_modified_path is None:
        win32api.MessageBox(
            0, "Please download report from SSRS", "Report Not Found")
        return None

    return xlwings.Book(last_modified_path)


def get_archived_work_order_data():
    pass


def fill_in_data(sheet):
    # get header column IDs
    header = sheet.range('A1').expand('right').value  # Row 1
    part = header.index('Operation5')   # part name w/o job
    grade = header.index('Material')
    remark = header.index('Remark')
    op1 = header.index('Operation1')
    op2 = header.index('Operation2')
    op3 = header.index('Operation3')

    # get engineering BOM dataand previous work order data
    # TODO: fetch engineering data on demand
    #   1) read JobStandards on first occurrence (if BOM not read)
    #   2) read BOM on first occurrence of part name not matching regex
    eng.force_cvn_mode = True
    eng_data = eng.get_part_data()
    archived_data = get_archived_work_order_data()

    i = 2
    while sheet.range(i, 1).value:
        item = partial(sheet.range, i)

        part_name = item(part).value
        update_grade = (item(grade).value is None)
        has_archive_data = (part_name in archived_data.keys())

        if update_grade:
            if has_archive_data:
                item(grade).value = archived_data[part_name]['grade']
            elif part_name in eng_data.keys():
                item(grade).value = eng_data[part_name].grade

        if has_archive_data:
            item(remark).value = archived_data[part_name]['remark']
            item(op1).value = archived_data[part_name]['op1']
            item(op2).value = archived_data[part_name]['op2']
            item(op3).value = archived_data[part_name]['op3']

        i += 1


# ~~~~~~~~~~~~~~~~~~~~~~~~~~ PROGRAM RUN ~~~~~~~~~~~~~~~~~~~~~~~~
if __name__ == '__main__':
    # TODO: logging
    LOG_FILENAME = 'log.txt'
    logging.basicConfig(filename=LOG_FILENAME, level=logging.ERROR)

    try:
        determine_processing()
    except:
        logging.exception("Exception in root handler")
        raise
