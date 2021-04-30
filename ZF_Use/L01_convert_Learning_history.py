# -*- coding: utf-8 -*-
"""
Created on Sun Mar  3 11:05:58 2019
@author: Z659190
"""

from pandas import DataFrame
import pandas as pd
from pandas import read_csv
from datetime import date

import csv

#get path and filenames
root_dir = "C:/temp/learning/"
df = DataFrame()
df2 = DataFrame()

now_date = date.today().strftime("%Y%m%d")

try:
    # f_name = root_dir + 'China Localization_Training Entry File_Combined 20200123.xlsx'
    f_name = root_dir + 'China Localization_Training Entry File_Combined 20200123 - 16.xlsx'
    df = pd.read_excel(f_name, sheet_name='10 Training Catalogue', skip_blank_lines=True,
                                    parse_dates=False, dtype = {'SUBJ_AREA_1':str})

    df2 = pd.read_excel(f_name, sheet_name='20 Training Catalogue CHN', skip_blank_lines=True,
                                    parse_dates=False)

except Exception as e:
    print('read file failed:', f_name)
    print('error log', e)

file1 = root_dir +'item_data_ZF_CN_'+now_date+'.txt'
# file1 = root_dir +'item_data_ZFDEV_CNtest_'+now_date+'.txt'
# file2 = root_dir + 'item_locale_data_ZFDEV_CNtest_'+now_date + '.txt'
file2 = root_dir + 'item_locale_data_ZF_CN_'+now_date + '.txt'

try:
    df.dropna(how='all', inplace=True)
    df.drop(['!##!'],axis=1,inplace = True)
    # df.drop(['CANCEL_POLICY_ID'],axis=1,inplace = True)
    # df['SUBJ_AREA_1'] = df.apply(lambda x:x['SUBJ_AREA_1']+"!##!",axis=1)
    # df.rename(columns = {"SUBJ_AREA_1":"SUBJ_AREA_1!##!"},inplace=True)

    df['CPNT_SRC_ID'] = df.apply(lambda x:str(x['CPNT_SRC_ID']).upper(),axis=1)
    # df['CPNT_ID'] = df.apply(lambda x:'MIG2_'+str(x['CPNT_ID']).upper(),axis=1)
    df['CPNT_ID'] = df.apply(lambda x:str(x['CPNT_ID']).upper(),axis=1)
    df['DMN_ID'] = df.apply(lambda x:str(x['DMN_ID']).upper(),axis=1)


    df['LEVEL1_SURVEY'] = df.apply(lambda x:str(x['LEVEL1_SURVEY'])+"!##!",axis=1)
    df.rename(columns = {"LEVEL1_SURVEY":"LEVEL1_SURVEY!##!"},inplace=True)
    df.to_csv(file1,index=False,encoding ='UTF-8',sep='|',line_terminator='\n' )
    print('Write date',file1)

    df2.dropna(how='all', inplace=True)
    # df2['CPNT_ID'] = df2.apply(lambda x:'MIG2_'+str(x['CPNT_ID']).upper(),axis=1)
    df2['CPNT_ID'] = df2.apply(lambda x:str(x['CPNT_ID']).upper(),axis=1)
    df2['TGT_AUDNC'] = df2.apply(lambda x:str(x['TGT_AUDNC'])+"!##!",axis=1)
    df2.drop(['!##!'],axis=1,inplace = True)
    df2.rename(columns = {"TGT_AUDNC":"TGT_AUDNC!##!"},inplace=True)
    # print(df2.columns)
    df2.to_csv(file2,index=False,encoding ='UTF-8',sep='|',line_terminator='\n')
    print("write date",file2)
except Exception as e:
    print('Exception:', e)




