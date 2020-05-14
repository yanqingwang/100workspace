# -*- coding: utf-8 -*-


import os
import datetime
import pandas as pd
import numpy as np
import csv

# 获取目标文件夹的路径
root_dir = "C:/temp/30CDP"
output_root = 'C:/temp/30CDP_OUT'
os.chdir(root_dir)
df2 = pd.DataFrame()


# 获取当前文件夹中的文件名称列表

def read_file():
    filenames = os.listdir(root_dir)
    df = pd.DataFrame()
    df2 = pd.DataFrame()
    # 先遍历文件名
    for filename in filenames:
        if filename.startswith('CDP'):
            filepath = root_dir + '/' + filename
            # 遍历单个文件，读取行数
            try:

                df = pd.DataFrame(
                    (pd.read_csv(filepath_or_buffer=filepath, quoting=csv.QUOTE_NONE, delimiter=',', encoding='utf-8')))
                # df.loc[:,'filename'] = filename       #增加列
                # df.loc['new_line'] = '3'  "增加行
                df['filename'] = filename
                print(filepath)
                df2 = df2.append(df)
            except Exception as e:
                print('Exception:', filepath, e)
    return df2


def check_localID():
    # 打开当前目录下的result.txt文件，如果没有则创建

    df3 = pd.DataFrame()
    now_date = datetime.date.today().strftime("%Y%m%d")
    file2 = output_root + '/Output_t1_repeat_check_' + now_date + '.csv'
    df2 = read_file()
    df3 = df2.head(0)
    line2 = df2.loc[0]
    for index, line in df2.iterrows():
        # print('line2',line2['LocalID'],'line',line['LocalID'])
        try:
            if pd.isna(line['LocalID']) :
                # print(line)
                line['reason'] = 'no local id found'
                df3 = df3.append(pd.Series(line))
            elif (not pd.isna(line2['LocalID'])) and (line['LocalID'] == line2['LocalID']) and ( line['GlobalID'] != line2['GlobalID']) :
                line['reason'] = 'repeat_local_id'
                df3 = df3.append(pd.Series(line))
            # if line['ReportingUnit'] != line2['ReportingUnit'] and ( line['GlobalID'] == line2['GlobalID']):
            #     line['reason'] = 'cross_RU_transfer'
            #     df3 = df3.append(pd.Series(line))
        except Exception as e:
            print('check exceptin:',e)
            print('index', index,'line2',line2['LocalID'],'line',line)
        line2 = line
    df3.to_csv(file2)


if __name__ == '__main__':
    check_localID()