# -*- coding: utf-8 -*-
"""
@author: Z659190

"""

from datetime import date
import time
import pandas as pd
from collections import defaultdict


class FactorAnalyze(object):
    def __init__(self):
        self.root = 'c:/temp/CHN/'
        self.emp_basefile = 'EmployeeBase-20201111.xlsx'
        self.emp_file = 'Personal_Data_2020_11_11_11_00_22.xlsx'
        # self.employment_file = 'Employment_Info_2020_11_11_10_51_43.xlsx'
        self.employment_file = 'MatrixMGR_2020_11_13_02_53_57.xlsx'
        # self.employment_file = 'Employment_Info_2020_11_11_10_47_47.xlsx'
        self.personal_data = pd.DataFrame()
        self.emp_base = pd.DataFrame()
        self.employment = pd.DataFrame()
        self.now_date = date.today().strftime("%Y%m%d")
        self.result = defaultdict(dict)

    def get_country(self,gid):
        return self.result[str(gid)]

    # get map between RU and divisions
    def get_emp_data(self):
        emp_file = self.root + self.emp_basefile
        try:
            self.emp_base = pd.DataFrame(pd.read_excel(io=emp_file, sheet_name='Excel Output',
                                                           dtype = {'ZF Global ID': 'str'}, header=0, skiprows=2))
            # print('Emp data size', self.emp_base.head(2))
            print('Emp data size', self.emp_base.shape)
            self.emp_base.drop_duplicates(keep='first', inplace=True)
            df_dict = self.emp_base[['Employment ID (User ID)', 'Country (ID)']]
            # df_dict.drop_duplicates(keep='first',inplace=True)
            for gid, country in df_dict.itertuples(index=False):
                self.result[gid] = country
        except Exception as e:
            print('Handle Employee Data Exception:', emp_file, e)

    # get employee data and pre cleaning
    def get_personal_data(self):
        emp_file = self.root + self.emp_file
        try:
            self.personal_data = pd.DataFrame(pd.read_excel(io=emp_file, sheet_name='Sheet1',
                                                            dtype= {'person-id-external': 'str'}, header=0,
                                                            skiprows=0))

            self.personal_data['Country'] = self.personal_data.apply(lambda x: self.get_country(x['person-id-external']), axis=1)
            print('Personal Size', self.personal_data.shape)

        except Exception as e:
            print('Handle Change Data Exception:', emp_file, e)

    # get employee data and pre cleaning
    def get_employment_data(self):
        emp_file = self.root + self.employment_file
        try:
            # self.employment = pd.DataFrame(pd.read_excel(io=emp_file, sheet_name='Employment_Info',
            self.employment = pd.DataFrame(pd.read_excel(io=emp_file, sheet_name='Sheet1',
                                                            dtype= {'user-id': 'str'}, header=0,
                                                            skiprows=0))

            self.employment['Country'] = self.employment.apply(lambda x: self.get_country(x['user-id']), axis=1)
            print('Employment Size', self.employment.shape)

        except Exception as e:
            print('Handle employment Data Exception:', emp_file, e)

    def output_results(self):
        print('Prepare summary Data in one  ----------------------------------------------------')
        df_chn_data = pd.DataFrame()
        df_chn_empl = pd.DataFrame()
        file_res = self.root + 'Output_' + self.now_date + '_res.xlsx'
        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(file_res, engine='xlsxwriter')
            workbook = df_writer.book
            sheet_name = '10_GID_COUNTRY'
            df_chn_empl = self.emp_base[['ZF Global ID', 'Country (ID)']]
            df_chn_empl.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '20_personal'
            df_chn_data = self.personal_data.loc[(self.personal_data['Country'] == 'CHN')]
            df_chn_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            print('Total Personal', df_chn_data.shape)

            # sheet_name = '30_employment'
            # self.employment.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            # print('Total Employment', self.employment.shape)

            sheet_name = '31_employment'
            df_chn_empl = self.employment.loc[(self.employment['Country'] == 'CHN')]
            df_chn_empl.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            print('Total Employment', df_chn_empl.shape)

        except Exception as e:
            print('write file failed:', file_res, sheet_name)
            print('error log', e)
        finally:
            workbook.close()
            df_writer.close()

    def main(self):
        self.get_emp_data()
        self.get_personal_data()
        self.get_employment_data()
        self.output_results()


if __name__ == '__main__':
    time1 = time.time()

    obj_factor = FactorAnalyze()
    obj_factor.main()

    print("Done, Total running time", time.time() - time1)
