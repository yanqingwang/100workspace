# -*- coding: utf-8 -*-
"""
age range / hire / termination /
@author: Z659190
"""

from datetime import date
import time
import datetime
import pandas as pd
import xlsxwriter


def get_root():
    return 'C:/temp/analyze/'


def write_data(df_out_value, file_name):
    sheet_name = 'sheet1'

    try:
        df_writer = pd.ExcelWriter(file_name)
        # write data
        df_out_value.to_excel(excel_writer=df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        df_writer.save()
        df_writer.close()
        print("output file", file_name)
    except Exception as e:
        print('write file failed:', file_name)
        print('error log', e)


def pre_process(input):
    df = pd.DataFrame()
    df_out = pd.DataFrame(columns=['Group Name','User'])
    df_tmp = pd.DataFrame(columns=['Group Name','User'])
    df_new = pd.DataFrame(columns=['Group Name','User'])
    file = get_root() + input
    try:
        df = pd.DataFrame(pd.read_excel(io=file, sheet_name='Sheet1', header=0, skiprows=0))
    except Exception as e:
        print('Exception:', file,e)
    # print(df.columns)

    # get detail data
    for n_idx,n_row in df.iterrows():
        line_rs = {}
        try:
            df_tmp = df_out[df_out['Group Name'] == n_row['Group Name']]
            if df_tmp.empty:
                line_rs['Group Name'] = n_row['Group Name']
                line_rs['User'] = n_row['User Sys ID'] + '_' +  n_row['Username'] + '_' + n_row['First Name'] + '_' +  n_row['Last Name']
                df_out = df_out.append(pd.Series(line_rs),ignore_index=True)
            else:
                line_rs = pd.Series(df_tmp.iloc[0])
                # line_rs = df_tmp.tolist()
                line_rs['User'] = line_rs['User'] + ';' + n_row['User Sys ID'] + '_' +  n_row['Username'] + '_' + n_row['First Name'] + '_' +  n_row['Last Name']
                df_out.loc[df_out['Group Name'] == n_row['Group Name'],'User'] = line_rs['User']

        except Exception as e:
            line_rs['Group Name'] = n_row['Group Name']
            line_rs['User'] = n_row['User Sys ID'] + '_' +  n_row['Username'] + '_' + n_row['First Name'] + '_' +  n_row['Last Name']
            df_out = df_out.append(pd.Series(line_rs),ignore_index=True)
            print('read file failed :')
            print('error log', e)

    out_file = get_root() + file_out
    write_data(df_out,out_file)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    file_in = 'GUG_Authorization__2019_09_11.xlsx'
    # file_in = 'AP_Employee_Basic_Data_Light-20190802.xlsx'
    file_out = 'Output_' + now_date + '_GUG.xlsx'
    time1 = time.time()

    pre_process(file_in)

    print("Total running time", time.time() - time1)
