# -*- coding: utf-8 -*-
"""
检查有培训历史记录，但是不在中国的global ID
有global ID，但是没有历史记录的员工，需要显示员工姓名和入职日期
检查localID和global ID,姓名矛盾的地方

"""

from os import chdir, listdir
from datetime import date
import time
import datetime
import pandas as pd
import xlsxwriter
import numpy as np

from win32com.client import gencache, DispatchEx

def get_path():
    return 'c:/Users/z659190/Documents/10 Work/10 MyHRSuit/10 Project/131 learning/30_Workpackages/37 Data migration/'


def set_columns():
    return ['Global ID','Local ID', 'Learner Name', 'Item/ Program Name','Training Hours','Item Type','Start Date','End Date',
                           'Completion date','Expiration Date for Certifications','Vendor/ Instructor','Comments/ Remarks']


def get_str_date(l_date):
    if len(str(l_date)) > 10:
        return (str(l_date)[:10])
    else:
        return str(l_date)


def check_msg(gid, name, item, hours):
    msg = ""
    if gid == "":
        msg = msg + 'Global ID empty.'
    if name == "":
        msg = msg + 'Learner empty.'
    if item == '':
        msg = msg +'Item name empty.'
    if hours == "":
        msg = msg + 'Training hours empty.'
    if msg != "":
        return msg


def read_file():
    df_data_tmp = pd.DataFrame()
    df_data = pd.DataFrame()
    filenames = listdir(get_path() + '2nd Submission/')
    for fname in filenames:
        # if fname.startswith('test'):
        if fname.startswith('Learning'):
            try:
                f_name = get_path() + '2nd Submission/' +fname
                sheet_to_df_map = pd.read_excel(f_name, dtype = {'Global ID\n*':str,'Global ID':str}, sheet_name=None, skip_blank_lines=True, parse_dates=False)
                # sheets = pd.ExcelFile(fname)
                # # this will read the first sheet into df
                # for l_sheet in sheets:
                #     df_data = pd.DataFrame(pd.read_excel(io=fname, sheet_name=l_sheet, header=0, skiprows=0))
                # print(sheet_to_df_map.keys())
                for key in sheet_to_df_map.keys():
                    if not key in ['Notes', 'Format- various systems', 'Questions']:
                        df_data_tmp = sheet_to_df_map[key]
                        df_data_tmp.columns = set_columns()
                        df_data_tmp['FileName'] = fname
                        print(fname,df_data_tmp.shape)
                        df_data= df_data.append(df_data_tmp, sort=False)
                # print(df_data_tmp.head(1))

            except Exception as e:
                print('read file failed:', fname)
                print('error log', e)
    # print('df data',df_data.columns)
    print('Learning history',df_data.shape)
    # df_data = df_data.sort_values("Global ID", ascending=True, inplace=True)
    return df_data.fillna("")


def read_file2():
    df_ee = pd.DataFrame()
    df_status = pd.DataFrame()
    df_ee_list = pd.DataFrame()
    try:
        f_name = get_path() + '\Data migration status tracking.xlsx'
        df_status = pd.read_excel(f_name, sheet_name='Data migration', skip_blank_lines=True, parse_dates=False)
        print('df_status',df_status.columns)

        f_name2 = get_path() + '\EmployeeHeadcount-Page1-20191231.xlsx'
        df_ee_list = pd.read_excel(f_name2, sheet_name='Excel Output', dtype = {'ZF Global ID':str}, parse_dates=False,skiprows=2)
        # print(df_ee_list.columns)

        df_ee = df_ee_list[['ZF Global ID','First Name','Last Name','Company (Legal Entity ID)','Company (Label)','Reporting Unit (Reporting Unit ID)','Employee Class (Label)','Employment Type (Label)','External Agency & Contingent Worker']]
        # print(df_ee.columns)

        return df_status.fillna(""),df_ee.fillna("")

    except Exception as e:
        print('read status / employee file failed:')
        print('error log', e)


def check_learning_history(df_data, df_ee):

    df_data_ret = pd.DataFrame()
    df_ee_tmp = pd.DataFrame()

    # table 1
    df_data['RU'] = '9'
    df_data['Error_MSG'] = df_data.apply((lambda x: check_msg(x['Global ID'],x['Learner Name'],x['Item/ Program Name'],x['Training Hours'])),axis=1)
    df_data['Global ID'] = df_data.apply((lambda x: '999999999' if (x['Global ID'] == "") else str(x['Global ID']).strip()),axis=1)

    # df_data_ret = df_data
    df_data.sort_values("Global ID", ascending=True, inplace=True)
    df_data_ret = df_data_ret.append(df_data,sort=False,ignore_index = True)
    for l_idx, i_row in df_data_ret.iterrows():
        try:
            # if str(i_row['Global ID']).strip() == '10147126':
            #     print('gid',l_idx)
            df_ee_tmp = []
            df_ee_tmp = df_ee[df_ee['ZF Global ID'] == str(i_row['Global ID']).strip()]
            # print(df_ee_tmp.shape)
            if not df_ee_tmp.empty:
                ru = pd.Series(df_ee_tmp.iloc[0])['Reporting Unit (Reporting Unit ID)']
                # if str(i_row['Global ID']).strip() == '10147126':
                #     print(l_idx, 'global id', i_row['Global ID'], df_ee_tmp.shape, 'RU',ru )
                # df_data_ret.at[l_idx,'RU'] = df_data_ret.at[l_idx,'Global ID'] + str(ru)
                df_data_ret.at[l_idx,'RU'] = str(ru)
                # df_ee_tmp.drop(df_ee_tmp.index, inplace=True)
        except Exception as e:
            print('read RU failed with error log:', e)
    return df_data_ret


def check_data(df_data, df_status, df_ee):

    df_ee_check = pd.DataFrame()
    df_ee_ru = pd.DataFrame()

    df_ee['ZF Global ID'] = df_ee.apply((lambda x: '999999999' if (x['ZF Global ID'] == "") else str(x['ZF Global ID']).strip()),axis=1)

    cc_list = df_status['CC Code'].tolist()
    # table 2
    df_ee['InScope'] =""
    df_ee['Learning_His'] =""
    df_ee['LearnRecords_No'] =0
    df_ee['EmpCount_WithLearning'] =0
    df_ee['EmpCount'] =1
    for e_idx, e_row in df_ee.iterrows():
        try:
            x,y = 0, 0
            df_data_tmp = df_data[df_data['Global ID'] == str(e_row['ZF Global ID']).strip()]
            x,y = df_data_tmp.shape
            df_ee.at[e_idx,'LearnRecords_No'] = int(x)
            if x > 0:
                df_ee.at[e_idx, 'Learning_His'] = 'Yes'
                df_ee.at[e_idx, 'EmpCount_WithLearning'] = 1
            else:
                df_ee.at[e_idx, 'Learning_His'] = 'No'
                df_ee.at[e_idx, 'EmpCount_WithLearning'] = 0

            if e_row['Company (Legal Entity ID)'] in cc_list:
                df_ee.at[e_idx,'InScope'] = 'In Scope'
            else:
                df_ee.at[e_idx,'InScope'] = 'Out Scope'
            # print(df_ee.head())
        except Exception as e:
            print('check_data_failed with error log:',e )

    df_ee = df_ee.infer_objects()
    # df_ee[['LearnRecords_No']] = df_ee[['LearnRecords_No']].astype(int)
    # df_ee[['LearnRecords_No']] = df_ee[['LearnRecords_No']].apply(pd.to_numeric())

    print(df_ee.head())
    df_ee_check = df_ee.groupby(['InScope','Company (Label)', 'Reporting Unit (Reporting Unit ID)','Employee Class (Label)'])['LearnRecords_No', 'EmpCount_WithLearning', 'EmpCount'].sum().reset_index().sort_values("InScope", ascending=True)
    df_ee_check['Average_Items'] = round(df_ee_check['LearnRecords_No'] / df_ee_check['EmpCount_WithLearning'],2)
    df_ee_check['Percentage_with_learning'] = round(df_ee_check['EmpCount_WithLearning'] / df_ee_check['EmpCount'],2)

    df_ee_ru = df_ee.groupby(['InScope','Company (Label)', 'Reporting Unit (Reporting Unit ID)'])['LearnRecords_No', 'EmpCount_WithLearning', 'EmpCount'].sum().reset_index().sort_values("InScope", ascending=True)
    df_ee_ru['Average_Items'] = round(df_ee_ru['LearnRecords_No'] / df_ee_ru['EmpCount_WithLearning'],2)
    df_ee_ru['Percentage_with_learning'] = round(df_ee_ru['EmpCount_WithLearning'] / df_ee_ru['EmpCount'],2)

    file_out = 'Output_' + now_date + '_check_Learning_history_data.xlsx'
    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(get_path()+file_out,engine='xlsxwriter')
        workbook = df_writer.book

        sheet_name = '00_history'
        df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=True)

        sheet_name = '10_ee_detail'
        df_ee.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '90_Key Info Missing'
        df_data.dropna(axis=0, subset=['Error_MSG']).to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '92_Percentage with EmpClass'
        df_ee_check.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '93_Percentage without EmpClass'
        df_ee_ru.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        # sheet_name = '16_overview'
        sheet_name = '95_Learning records vs Emp no'
        pvt_tmp = pd.pivot_table(df_ee, index=['InScope','Company (Label)','Reporting Unit (Reporting Unit ID)'], values=['LearnRecords_No','EmpCount_WithLearning','EmpCount'],
                                 aggfunc=[np.sum], fill_value=0)
        pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")


        sheet_name = '97_Learning - Employee class'
        pvt_tmp = pd.pivot_table(df_ee, index=['InScope','Company (Label)','Reporting Unit (Reporting Unit ID)','Employee Class (Label)'], values=['LearnRecords_No','EmpCount_WithLearning','EmpCount'],
                                 aggfunc=[np.sum], fill_value=0)
        pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")

        sheet_name = '11_company_check'
        pvt_tmp = pd.pivot_table(df_ee,index=['InScope','Company (Label)'], values=["LearnRecords_No",'EmpCount_WithLearning','EmpCount'],
                                 aggfunc=[np.sum], fill_value=0)
        pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")

        sheet_name = '13_EmpType'
        pvt_tmp = pd.pivot_table(df_ee,index=['InScope','Company (Label)', 'Reporting Unit (Reporting Unit ID)',
                                               'Employment Type (Label)'], values=['LearnRecords_No','EmpCount_WithLearning','EmpCount'],
                                 aggfunc=[np.sum], fill_value=0)
        pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")

        # sheet_name = '14_External'
        sheet_name = 'Learning records vs Emp no'
        pvt_tmp = pd.pivot_table(df_ee,index=['InScope','Company (Label)', 'Reporting Unit (Reporting Unit ID)',
                                               'External Agency & Contingent Worker'], values=['LearnRecords_No','EmpCount_WithLearning','EmpCount'],
                                 aggfunc=[np.sum], fill_value=0)
        pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")

        sheet_name = '15_overview--EEType'
        pvt_tmp = pd.pivot_table(df_ee, index=['InScope','Company (Label)','Reporting Unit (Reporting Unit ID)','Learning_His'], values=["LearnRecords_No"],
                                 aggfunc=[np.sum,'count'], fill_value=0)
        pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")

        sheet_name = '15_overview--External'
        pvt_tmp = pd.pivot_table(df_ee, index=['InScope','Company (Label)','Reporting Unit (Reporting Unit ID)','External Agency & Contingent Worker','Learning_His'], values=["LearnRecords_No"],
                                 aggfunc=[np.sum,'count'], fill_value=0)
        pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")



        workbook.close()
    except Exception as e:
        print('write file failed:', file_out)
        print('error log', e)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    df_data2 = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    time1 = time.time()

    chdir(get_path())

    df_data = read_file()
    df_status, df_ee = read_file2()
    df_data2 = check_learning_history(df_data, df_ee)
    check_data(df_data2, df_status, df_ee)

    print("Total running time", time.time() - time1)
