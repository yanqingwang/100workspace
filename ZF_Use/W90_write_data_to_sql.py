# -*- coding: utf-8 -*-
"""
Read data from a file, and then write it to sqlite
@author: Z659190
"""

from datetime import date
import time
import pandas as pd
from ZFlib import data_storage as ds

import re

class Headcount_raw_handling(object):
    def __init__(self):
        self.root = 'c:/temp/10headcount/'
        self.empfile = 'EmployeeHeadcount-Page1-20210428.xlsx'
        # self.empfile = 'EmployeeHeadcount-Page1-20200531 - Copy.xlsx'
        self.status = ['Active','Furlough','Unpaid Leave','Paid Leave']
        self.sql_table = 'Headcount'
        self.factors_table = 'Headcount_factors'
        self.df_data = pd.DataFrame()

    def get_date(self):
        date_list = re.findall('[1-2][0-9]{3}[0-1][0-9][0-3][0-9]', self.empfile)
        if len(date_list) > 0:
            return date_list[0]
        else:
            return '00000000'

    def get_emp_data(self):
        df = pd.DataFrame()
        file = self.root + self.empfile
        try:
            df = pd.DataFrame(pd.read_excel(io=file, sheet_name='Excel Output',  dtype = {'ZF Global ID':str,'Reporting Unit (Reporting Unit ID)':str},header=0, skiprows=2))
            df['Source_Date'] = self.get_date()
            df['FileName'] = self.empfile
            df = df.rename(columns={"Reporting Unit (Reporting Unit ID)": "RU"
                                    })
            df['active'] = df.apply((lambda x: 'Active' if x['Employee Status (Label)'] in self.status else ""),axis=1)
            df = df[df['active'] == 'Active']

            df = df[df['Country (ID)','Company (Label)','Division Short Text','External Agency & Contingent Worker','Employee Class (Label)','Employment Type (Label)',
                        'Employee Status (Label)','Source_Date','ZF Global ID','ZID','First Name','Last Name']]

            print(df.head(2))
        except Exception as e:
            print('Exception:', file,e)
        # df = df[df[]]
        return df

    def get_factors(self,data_table):

        df_factors = pd.DataFrame()
        df_factors = data_table.groupby(['Country (ID)','Company (Label)','Division Short Text','External Agency & Contingent Worker','Employee Class (Label)','Employment Type (Label)',
                                         'Employee Status (Label)','Source_Date']).agg({'ZF Global ID':'count','FTE (Group Reporting)':'sum'}).reset_index()
        print(df_factors)
        return df_factors


    def main(self):
        df_data = pd.DataFrame()
        df_data = self.get_emp_data()
        ds.update_raw_data(df_data,self.sql_table)
        df_factors = self.get_factors(df_data)
        ds.update_factor_data(df_factors,self.factors_table)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    time1 = time.time()

    Headcount_raw_handling().main()

    print("Total running time", time.time() - time1)
