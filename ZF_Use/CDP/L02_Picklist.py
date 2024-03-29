# -*- coding: utf-8 -*-
"""
@author: Z659190
提供数据预处理的公共方法

"""
import pandas as pd
import time
from datetime import date
import csv

target_grps1 =  ['cust_employeeClass','cust_employmentType','cust_eventReason',
                 'cust_EventReason','cust_localjobcode_chn']

target_grps2 =  ['cust_politicalstatus_chn','ETHNICGROUP_CHN','HUKOU_CHN',
                 'employmentType','EmployeeClass','employee-status','cust_localsystem']

target_grps3 = {'X088','X011','X005'}

format_columns = {'External Code': 'str', 'Picklist Value.External Code': 'str'}
target_cols = ['Picklist.Code','External Code','Picklist Value.External Code','US English','Default Value','Chinese (China)']

class FactorAnalyze(object):
    def __init__(self):
        self.root = 'c:/temp/CHN/'
        # self.file = 'Picklist-Values.xlsx'
        self.file = 'Picklist-Values-prd-20210226.xlsx'
        self.data = pd.DataFrame()

    def handle_data(self):
        df_data = pd.DataFrame()
        lv_file = self.root + self.file
        try:
            self.data = pd.DataFrame(pd.read_excel(io=lv_file, sheet_name='Picklist-Values',
                                                   dtype=format_columns, header=0,
                                                   skiprows=1))
            # self.data = pd.DataFrame((pd.read_csv(filepath_or_buffer=lv_file,quoting=csv.QUOTE_NONE,delimiter=',',
            #                                      encoding='utf-8',dtype=format_columns, skiprows=1)))
            # self.data = self.data[target_cols]
            df_data = self.data.loc[self.data['Picklist.Code'].isin(target_grps1)].reset_index()
            # df_data['CDP'] = df_data.apply(lambda x: x['External Code'].startswith('CDP'), axis=1 )
            df_data['CDP'] = df_data['External Code'].str.startswith('CDP')
            # df_data['CDP'] = df_data['Non-unique External Code'].str.startswith('')
            df_data2 = self.data.loc[self.data['Picklist.Code'].isin(target_grps2)].reset_index()
            df_data3 = self.data.loc[self.data['Picklist Value.External Code'].isin(target_grps3)].reset_index()


        except Exception as e:
            print('Handle BCS Data Exception:', self.file, e)

        print('Output Data in one  ----------------------------------------------------')

        file_res = self.root + 'Output_picklist_' + date.today().strftime("%Y%m%d") + '_res.xlsx'
        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(file_res, engine='xlsxwriter')
            workbook = df_writer.book
            sheet_name = '10_res'
            df_data1 = df_data[(df_data['CDP'] )| (df_data['Picklist Value.External Code'] == 'X011')].reset_index()
            df_data1.drop(columns=['CDP'],inplace=True)
            df_data1.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '20_res'
            df_data2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '30_res'
            df_data3.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '10_res_s'
            df_data1[target_cols].to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '20_res_s'
            df_data2[target_cols].to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '30_res_s'
            df_data3[target_cols].to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            df_data4 = df_data1[target_cols].append(df_data2[target_cols])
            df_data4 = df_data4.append(df_data3[target_cols])
            df_data4.drop_duplicates(inplace=True)

            sheet_name = '40_combine'
            df_data4.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        except Exception as e:
            print('write file failed:', file_res, sheet_name)
            print('error log', e)
        finally:
            workbook.close()
            df_writer.close()


if __name__ == '__main__':
    time1 = time.time()

    obj_factor = FactorAnalyze()
    obj_factor.handle_data()

    print("Done, Total running time", time.time() - time1)
