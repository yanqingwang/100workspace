# -*- coding: utf-8 -*-
"""
@author: Z659190
人力资源数据分析尝试
包括静态数据 / 时间相关静态数据 / 增量存量趋势数据 / HR相关数据分析

"""

from datetime import date
import time
import pandas as pd
from collections import defaultdict
from ZF_Use.data_analyze import EC_analyze_clean_fous as ecf
from ZF_Use.data_analyze import EC_analyze_lib as el


def conv(x):
    # print(type(x))
    # print(x)
    if not pd.isna(x):
        # print(x.date().strftime("%Y%m"))
        # print(pd.Period(x,'M'))
        return pd.Period(x,'M')
        # return x.date().strftime("%Y%m")

class FactorAnalyze(object):
    def __init__(self):
        self.root = 'c:/temp/new_try/'
        self.emp_file = 'EmployeeBasicInfo_AP_With_Inactive-20201031 - Copy1.xlsx'
        self.emp_data = pd.DataFrame()
        self.emp_data2 = pd.DataFrame()
        self.df2 = pd.DataFrame()
        self.now_date = date.today().strftime("%Y%m%d")
        self.result = defaultdict(dict)
        self.date_range = pd.period_range('2020-01-01', '2020-12-31', freq='M')
        # self.range = pd.Period('2020',freq='A-DEC')
        print(self.date_range)

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

        try:

            # self.emp_data['NewHP1'] = pd.PeriodIndex(self.emp_data['NewH'])
            self.emp_data['NewHP'] = self.emp_data['NewH'].apply(conv)
            self.emp_data['NewTP'] = self.emp_data['NewT'].apply(conv)
            self.emp_data2 = self.emp_data.loc[self.emp_data['NewTP'].isin(self.date_range)].reset_index()
            # df_data = self.data.loc[self.data['Picklist.Code'].isin(target_grps1)].reset_index()
            # self.merge_data['NewHP'] = self.emp_data['NewH'].dt.to_period('M')
            # self.merge_data['NewTP'] = self.merge_data['NewT'].dt.to_period('M')
            # self.emp_data['NewTP'] = self.emp_data['Termination Date'].dt.to_period('M')
            # self.merge_data['NewHP'] = self.merge_data['NewH'].dt.to_period()

            # self.df2 = pd.DataFrame(self.date_range,columns=['Month'])
            # for index,row in self.df2.iterrows():

        except Exception as e:
            print('Period Data Exception:', e)
            # for index,row in self.df2.iterrows():

        df_res = pd.DataFrame()
        for month in self.date_range:
            line = {}
            try:
                line['Month'] = month
                # last_day = month.to_timestamp(how='end')
                # print(last_day)
                line['Active_total'] = len(self.emp_data.loc[(self.emp_data['NewHP'] <= month) &
                                                             ((self.emp_data['NewTP'] > month) | (self.emp_data['NewTP'].isna()))].index)
                                                             # (self.emp_data['NewTP'].isna())].index)
                line['NewHire'] = len(self.emp_data.loc[(self.emp_data['NewHP'] == month)].index)
                line['Termination'] = len(self.emp_data.loc[(self.emp_data['NewTP'] == month)].index)
                df_res = df_res.append(pd.Series(line), ignore_index=True)
            except Exception as e:
                print('Count Data Exception:', e)

        print(df_res)


    def output_results(self):
        print('Prepare summary Data in one  ----------------------------------------------------')
        file_res = self.root + 'Output_' + self.now_date + '_res.xlsx'
        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(file_res, engine='xlsxwriter')
            workbook = df_writer.book
            sheet_name = '20_emp'
            self.emp_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '30_emp'
            self.emp_data2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '40_emp'
            self.df2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        except Exception as e:
            print('write file failed:', file_res, sheet_name)
            print('error log', e)
        finally:
            workbook.close()
            df_writer.close()

    def main(self):
        print("main")
        self.get_emp_data()
        self.output_results()
        pass


if __name__ == '__main__':
    time1 = time.time()

    obj_factor = FactorAnalyze()
    obj_factor.main()

    print("Done, Total running time", time.time() - time1)
