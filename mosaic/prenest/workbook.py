
import pandas
import os
import logging

from subprocess import call
from openpyxl import Workbook

from mosaic import config
from mosaic.prenest.input_handlers import range_to_list

ENG_DATA_SHEET_NAME = "Engineering Data"
ANALYZE_SHEET_NAME = "Analyzing Data"
MANUAL_INPUT_SHEET_NAME = "Manual Data Entry"

MANUAL_DETAIL_LENGTH_OFFSET = 1.0


class PrenestDataWorkbook:

    def __init__(self, job, caller):
        self.job = job.upper()
        self.xl_writer = pandas.ExcelWriter(caller)


    def import_flange_data(self):
        flg_data_file = os.path.join(
            config['dirs']['eng_jobs'], self.job, 'CAM', 'FlangeData.xlsx')
        if not os.path.exists(flg_data_file):
            call([config['files']['flg_data_exec'], self.job])

        fd = pandas.read_excel(flg_data_file)
        fd.to_excel(self.xl_writer, sheet_name=ENG_DATA_SHEET_NAME, startrow=2, startcol=1)
        logging.info(fd)


    def generate_flange_data_manual(self):
        manual_df = pandas.read_excel(
            self.xl_writer, sheet_name=MANUAL_INPUT_SHEET_NAME)
        eng_df = pandas.read_excel(
            self.xl_writer, sheet_name=ENG_DATA_SHEET_NAME)

        for line in manual_df:
            # handle girder and flange ranges
            girders = range_to_list(line[0])
            flanges = range_to_list(line[1])
            for girder in girders:
                for flange in flanges:
                    eng_df.range(
                        e, 1).value = '{}_{}-{}'.format(self.job, girder, flange)  # Mark
                    # Radius
                    eng_df.range(e, 2).value = line[2]
                    # Width
                    eng_df.range(e, 3).value = line[3]
                    # Detail_Length
                    eng_df.range(
                        e, 4).value = line[4] - MANUAL_DETAIL_LENGTH_OFFSET
                    # Adjust_Length
                    eng_df.range(e, 5).value = MANUAL_DETAIL_LENGTH_OFFSET
                    # Order_Length
                    eng_df.range(e, 6).value = line[4]
                    # Thickness
                    eng_df.range(e, 7).value = line[5]
                    # Grade
                    eng_df.range(e, 17).value = line[6]
                    eng_df.range(e, 18).value = str(
                        int(line[7])).zfill(5)                   # Item

        writer = pandas.ExcelWriter(self.caller_file)
        eng_df.to_excel(sheet_name=ENG_DATA_SHEET_NAME,
                        startrow=2, startcol=1, index=False)

    def load_flange_data(self):
        pass

    def analyze(self):
        # 1) load parts into dataframe
        # 2) load plates into dataframe
        # 3) create nesting groups
        # 4) nest
        # 5) export nesting results

        pass

    def export_for_prenesting(self):
        # 1) generate xml files
        # 2) ask to run xml import
        # 3) if 2, run xml import

        pass
