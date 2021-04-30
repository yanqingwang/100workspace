# -*- coding: utf-8 -*-
"""
Created on Sun Mar  3 11:05:58 2019
@author: Z659190
"""

from os import chdir,listdir
from datetime import date
from time import sleep
import pandas as pd
import os

import csv

#get path and filenames
root_dir = "C:/temp/Delimit"
chdir(root_dir)


def write_file(filename,df_data):
    try:
        # 创建一个excel

        df_writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        # workbook = df_writer.book

        sheet_name = '80_res'
        df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

    except Exception as e:
        print('write summary factor failed:', filename, sheet_name)
        print('error log', e)
    finally:
        # workbook.close()
        df_writer.close()


#get file list in the path
filenames=listdir(root_dir)
df = pd.DataFrame()
df2 = pd.DataFrame()
#open the target file, if not exist, create new one
now_date = date.today().strftime("%Y%m%d")
file2 = root_dir + '/t1_result_'+now_date+'.csv'
file3 = root_dir + '/t1_result_'+now_date+'.xlsx'
for filename in filenames:
    if filename.startswith('Position'):
        filepath = root_dir+'/'+filename
        #for all files, read and process
        try:
            sheet_to_df_map = pd.read_excel(io=filepath, sheet_name=None, header=0, skiprows=0,parse_dates=False)
            # df.loc[:,'filename'] = filename       #add new column
            # df.loc['new_line'] = '3'  "add new row
            # df['filename'] = filename
            print(filepath)
            for key in sheet_to_df_map.keys():
                if not key in ['Notes', 'Format- various systems', 'Questions']:
                    df = sheet_to_df_map[key]
                    # df_data_tmp.columns = set_columns()
                    df2 = df2.append(df, sort=False)
            # df2 = df2.append(df,sort=False)
        except Exception as e:
            print('Exception:', filepath,e)
df2.to_csv(path_or_buf=file2,sep="|",line_terminator ='!##!'+os.linesep)
# df2.to_csv(path_or_buf=file2,sep="|",line_terminator ='!##!'+'\\n')
write_file(file3,df2)

print("File output to ",file2)

sleep(10)




