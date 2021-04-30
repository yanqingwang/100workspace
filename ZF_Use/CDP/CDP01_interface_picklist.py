# -*- coding: utf-8 -*-
"""
Created on Sun Mar  3 11:05:58 2019
@author: Z659190
"""

from os import chdir,listdir
from datetime import date
from time import sleep
from pandas import DataFrame
from pandas import read_csv

import csv

#get path and filenames
root_dir = "C:/temp/31CDP_INT"
# output_root = 'C:/temp/30CDP_OUT'
chdir(root_dir)

#get file list in the path
filenames=listdir(root_dir)
df = DataFrame()
df2 = DataFrame()
df3 = DataFrame()
#open the target file, if not exist, create new one
now_date = date.today().strftime("%Y%m%d")
file2 = root_dir + '/t1_result_'+now_date+'.csv'
for filename in filenames:
    if filename.startswith('PicklistData_zffriedric.csv'):
        filepath = root_dir+'/'+filename
        #for all files, read and process
        try:

            df = DataFrame((read_csv(filepath_or_buffer=filepath,quoting=csv.QUOTE_NONE,delimiter=',',encoding='utf-8',error_bad_lines=False)))
            # df.loc[:,'filename'] = filename       #add new column
            # df.loc['new_line'] = '3'  "add new row
            # df['filename'] = filename
            print(filepath)
            df2 = df2.append(df,sort=False)
        except Exception as e:
            print('Exception:', filepath,e)

print(df2.shape,df2.columns)

df2 = df2[df2['status'] == 'ACTIVE']
# try:
#     df2['CDP_Check'] = df2.apply((lambda x: 'CDP' if (str(x['external_code']).startswith('CDP')) else ""), axis=1)
# except Exception as e:
#     print('Exception:',  e)

picklist = ['ecMaritalStatus','ETHNICGROUP_CHN','country','ISOCountryList','HUKOU_CHN','cust_politicalstatus_chn']
for pk in picklist:
    df3 = df3.append(df2[df2['^picklistId'] == pk],sort=False)


df3 = df3.append(df2[df2['^picklistId'] == 'cust_employmentType' ],sort=False)
df3 = df3.append(df2[df2['^picklistId'] == 'cust_employeeClass' ],sort=False)
df3 = df3.append(df2[df2['^picklistId'] == 'contractType' ],sort=False)
df3 = df3.append(df2[df2['^picklistId'] == 'cust_employmentType' & df2['external_code'] == 'CDP'],sort=False)


df3.to_csv(file2)
print("File output to ",file2)
sleep(10)




