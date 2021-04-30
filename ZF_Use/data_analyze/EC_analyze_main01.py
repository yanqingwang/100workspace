# -*- coding: utf-8 -*-
"""
@author: Z659190
Input Source: Employee basic data with inactive data / RU and division mapping table

人力资源数据分析尝试
包括静态数据 / 时间相关静态数据 / 增量存量趋势数据 / HR相关数据分析

"""

from datetime import date
import time
import pandas as pd
import re
from collections import defaultdict
from ZF_Use.data_analyze import EC_analyze_function as ea
from ZF_Use.data_analyze import EC_analyze_trends as eat
from ZF_Use.data_analyze import EC_analyze_clean_fous as ecf
from ZF_Use.data_analyze import EC_analyze_lib as el


class FactorAnalyze(object):
    def __init__(self):
        self.country = {"AU": "AUS", "CN": "CHN", 'ID': 'IDN', 'JP': 'JPN', 'KR': "KOR", 'MY': 'MYS',
                        'PH': 'PHL', 'SG': 'SGP', 'TW': 'TWN', 'TH': 'THA', 'AE': 'ARE', 'VN': 'VNM'}
        self.division = ['A', 'B', 'C', 'E', 'I', 'P', 'R', 'T', 'U', 'Z', 'W']
        self.root = 'c:/temp/new_try/'
        # self.emp_file = 'EmployeeBasicInfo_AP_With_Inactive-20210131 - Global.xlsx'
        # self.emp_file = 'EmployeeBasicInfo_AP_With_Inactive-20201231.xlsx'
        self.emp_file = 'EmployeeBasicInfo_AP_With_Inactive-20210131 - Global.xlsx'
        # self.emp_file = 'EmployeeBasicInfo_AP_With_Inactive-20201231 - Copy.xlsx'
        # self.bcs_file = 'Global_DIV_RU.xlsx'
        self.bcs_file = 'Global_DIV_RU_20210208.xlsx'
        self.emp_data = pd.DataFrame()
        self.bcs_data = pd.DataFrame()
        self.merge_data = pd.DataFrame()
        self.now_date = date.today().strftime("%Y%m%d")
        self.result = defaultdict(dict)

    # get map between RU and divisions
    def get_bcs_data(self):
        bcs_file = self.root + self.bcs_file
        try:
            self.bcs_data = pd.DataFrame(pd.read_excel(io=bcs_file, sheet_name='Sheet1', header=0, skiprows=0))
            print('bcs data', self.bcs_data.head(2))
            # self.bcs_data.drop(columns=['Nothing'], inplace=True)
            self.bcs_data = self.bcs_data[self.bcs_data["Country"] != ""]
            self.bcs_data.drop_duplicates(keep='first', inplace=True)
            # print(df.columns)
            self.bcs_data.sort_values(by=['RU'], inplace=True)
        except Exception as e:
            print('Handle BCS Data Exception:', bcs_file, e)
        # print(df.columns)
        # return self.bcs_data.fillna('')

    # get employee data and pre cleaning
    def get_emp_data(self):
        emp_file = self.root + self.emp_file
        try:
            self.emp_data = pd.DataFrame(pd.read_excel(io=emp_file, sheet_name='Excel Output',
                                                       dtype= ecf.format_columns, header=0,
                                                       # skiprows=2))
                                                       # parse_dates=['Termination Date'],infer_datetime_format=True,
                                                       skiprows=2))
            self.emp_data = self.emp_data.rename(columns=ecf.Columns_rename)

            # 数据清理
            # ecf.clean_data(df_raw)
            self.emp_data = self.emp_data[self.emp_data['External Agency Worker'] != 1]
            self.emp_data.drop(columns=ecf.clean_columns, inplace=True)

            # self.emp_data['Hire Date'].to_timestamp(freq='D',axis=1)
            # self.emp_data['Hire Date'].to_timestamp(freq='D',axis=1)
            # self.emp_data['Hire Date'] = self.emp_data.apply(lambda x: el.conv_date(x['Hire Date']),axis=1)

            # 更新关键日期
            self.emp_data['NewH1'] = self.emp_data.apply(lambda x: el.update_date(x['Hire Date'], x['Hire Date.1']),
                                                         axis=1)
            self.emp_data['NewH'] = self.emp_data.apply(lambda x: el.update_date(x['Original Start Date'], x['NewH1']),
                                                        axis=1)
            self.emp_data['NewT'] = self.emp_data.apply(
                lambda x: el.update_date(x['Termination Date'], x['Termination Date.1']), axis=1)

            # self.emp_data.astype({'NewH':'datetime'}, errors='ignore')
            # self.emp_data.DatetimeIndex(['NewH'],axis=1)

            # self.emp_data.drop(columns=ecf.clean_col_date, inplace=True)

            self.emp_data.sort_values(by=['ZF Global ID', 'NewH'], ascending=[True, False], inplace=True)
            self.emp_data.drop_duplicates(subset=['ZF Global ID'], keep='first', inplace=True)

            print(self.emp_data.head(2))

        except Exception as e:
            print('Handle Employee Data Exception:', emp_file, e)

    # get employee data and pre cleaning
    def get_combine_data(self):

        try:
            # self.merge_data = pd.merge(self.emp_data, self.bcs_data, how='left', on='RU',left_index=True,
            self.merge_data = pd.merge(self.emp_data, self.bcs_data, how='left', on='RU',
                                       indicator=False)
            self.merge_data.reset_index(inplace=True)
            print(self.merge_data.head(2))
            self.merge_data['EmploymentType'] = self.merge_data.apply(
                lambda x: el.get_mgr(x['Employment Type (Label)']), axis=1)
            self.merge_data['JF'] = self.merge_data.apply(lambda x: x['Job Classification (Job Code)'][0:2], axis=1)
            self.merge_data['ServiceYear'] = self.merge_data.apply(lambda x: el.get_service_year(x['NewH']), axis=1)
            self.merge_data['Age'] = self.merge_data.apply(lambda x: el.get_age_range(x['Date Of Birth']), axis=1)
            self.merge_data['ServiceMonths'] = self.merge_data.apply(lambda x: el.get_service_year_termination(x['NewH'],x['NewT']), axis=1)
        except Exception as e:
            print('Static Data Exception:', e)

        try:
            self.merge_data['NewHP'] = self.merge_data['NewH'].apply(el.conv_date)
            self.merge_data['NewTP'] = self.merge_data['NewT'].apply(el.conv_date)
            # self.merge_data['NewTP'] = self.merge_data['NewT'].dt.to_period('M')
            # self.merge_data['NewTP'] = self.merge_data['Termination Date'].dt.to_period('M')
            # self.merge_data['NewHP'] = self.merge_data['NewH'].dt.to_period()

        except Exception as e:
            print('Period Data Exception:', e)

    # output static analyze result
    def out_analyze_res(self):
        try:
            file_static = self.root + 'Output_' + self.now_date + '_res_static.xlsx'
            obj_ea = ea.AnalyzeObj(file_static)
            obj_ea.main(self.merge_data)
        except Exception as e:
            print('Static Analyze Exception:', e)

        # Time period trends analyze
        try:
            file_time = self.root + 'Output_' + self.now_date + '_res_time.xlsx'
            obj_eat = eat.AnalyzeTimeObj(file_time)
            obj_eat.main(self.merge_data)
        except Exception as e:
            print('Time Period Analyze Exception:', e)

    def output_results(self):
        print('Prepare summary Data in one  ----------------------------------------------------')
        file_res = self.root + 'Output_' + self.now_date + '_res.xlsx'
        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(file_res, engine='xlsxwriter')
            workbook = df_writer.book
            sheet_name = '10_bcs'
            self.bcs_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '20_emp'
            self.emp_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '30_merge'
            self.merge_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        except Exception as e:
            print('write file failed:', file_res, sheet_name)
            print('error log', e)
        finally:
            workbook.close()
            df_writer.close()

    def main(self):
        print("main")
        self.get_bcs_data()
        self.get_emp_data()
        self.get_combine_data()

        self.out_analyze_res()
        self.output_results()
        pass


if __name__ == '__main__':
    time1 = time.time()

    obj_factor = FactorAnalyze()
    obj_factor.main()

    print("Done, Total running time", time.time() - time1)
