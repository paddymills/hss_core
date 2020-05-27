
from pandas import read_excel
from os.path import join, exists
from subprocess import call

ENG_JOBS = r"\\hssieng\DATA\HS\JOBS"
FLG_DATA_EXEC = r"\\hssieng\Resources\HS\PROG\FlgXlsData.exe"


class FlangeData:

    def __init__(self, job):
        self.job = job

        self.flg_data_file = join(ENG_JOBS, self.job, 'CAM', 'FlangeData.xlsx')
        if not exists(flg_data_file):
            self.generate_flg_data()

        self.get_data()

    def get_data(self):
        self.data = read_excel(self.flg_data_file)

    def generate_flg_data(self):
        call([FLG_DATA_EXEC, self.job])
