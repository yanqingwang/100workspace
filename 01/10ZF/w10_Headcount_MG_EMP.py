# -*- coding: utf-8 -*-
"""
Get MG Grade Employee
@author: Z659190
"""

from datetime import date
import time
import pandas as pd
import xlsxwriter


def get_div(division_short):
    if division_short in ['A','B','C','E','I','P','R','T','U']:
        return division_short
    else:   #['-','D','F','G','H','M','MK','O','Q','S','V']:
        return 'Z'


class CommonUtility(object):
    def __init__(self):
        self.ap_country = ["ARE","AUS","CHN","JPN","KOR","MYS","PHL","SGP","THA","TWN","VNM","IDN","IND"]

    def get_region(self,country):
        try:
            if country in self.ap_country:
                return "AP"
            else:
                return "Unknown"
        except Exception as e:
            print('error log to get country', e)


class MgEmp(object):
    def __init__(self):
        self.mg_grade = ['Management Group 1','Management Group 2','Management Group 3', 'Management Group 4',
                         'ZF TRW OIP Pool','Executive Mgmt Group','Global Executive Team','Board of Management']
        self.root = 'c:/temp/10headcount/'
        self.empfile = 'EmployeeHeadcount-Page1-20200327.xlsx'
        self.comm = CommonUtility()

    def get_emp_data(self):
        df = pd.DataFrame()
        file = self.root + self.empfile
        try:
            df = pd.DataFrame(pd.read_excel(io=file, sheet_name='Excel Output',  dtype = {'ZF Global ID':str,'Reporting Unit (Reporting Unit ID)':str},header=0, skiprows=2))
            print(df.head(2))
            df['Division'] = df.apply(lambda x: get_div(x['Division Short Text']), axis=1)
            df['Region'] = df.apply(lambda x: self.comm.get_region(x['Country (ID)']), axis=1)
            print(df.columns)
        except Exception as e:
            print('Exception:', file,e)
        return df.fillna('')

    def out_put_rs(self, df_data):
        now_date = date.today().strftime("%Y%m%d")
        out_file = 'Output_' + now_date +'_MG_Level.xlsx'

        df_simple = pd.DataFrame()
        df_AP = pd.DataFrame()
        df_data = df_data[df_data['Employment Type (Label)'].isin(self.mg_grade)]

        df_simple = df_data[['Employment Type (Label)','Region','Country (ID)','Division','Reporting Unit (Reporting Unit ID)','ZF Global ID','ZID','First Name','Last Name','Business email address']]
        df_simple.sort_values(['Employment Type (Label)','Region'], ascending=True,inplace=True)

        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(self.root+out_file,engine='xlsxwriter')
            workbook = df_writer.book

            sheet_name = '10_Initial'
            df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '12_Global_MG'
            df_simple.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            df_AP = df_simple[df_simple['Region'] == 'AP']
            sheet_name = '14_AP_MG'
            df_AP.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            workbook.close()
        except Exception as e:
            print('write file failed:', out_file)
            print('error log', e)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    time1 = time.time()

    mg_handle = MgEmp()
    df_data = mg_handle.get_emp_data()
    mg_handle.out_put_rs(df_data)

    print("Total running time", time.time() - time1)
