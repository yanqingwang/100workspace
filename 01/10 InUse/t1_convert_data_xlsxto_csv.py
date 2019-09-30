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
root_dir = "C:/temp/convert"
chdir(root_dir)

#get file list in the path
filenames=listdir(root_dir)
df = pd.DataFrame()
df2 = pd.DataFrame()
#open the target file, if not exist, create new one
now_date = date.today().strftime("%Y%m%d")
file2 = root_dir + '/t1_result_'+now_date+'.csv'
for filename in filenames:
    if filename.startswith('ZF'):
        filepath = root_dir+'/'+filename
        #for all files, read and process
        try:
            df = pd.DataFrame(pd.read_excel(io=filepath, sheet_name='Sheet1', header=0, skiprows=0))
            # df.loc[:,'filename'] = filename       #add new column
            # df.loc['new_line'] = '3'  "add new row
            # df['filename'] = filename
            print(filepath)
            df2 = df2.append(df,sort=False)
        except Exception as e:
            print('Exception:', filepath,e)
df2.to_csv(path_or_buf=file2,sep="|",line_terminator ='!##!'+os.linesep)
# df2.to_csv(path_or_buf=file2,sep="|",line_terminator ='!##!'+'\\n')
print("File output to ",file2)
sleep(10)




