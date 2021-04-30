# -*- coding: utf-8 -*-
"""
Created on Apr 25 11:05:58 2019
@author: Z659190
"""

import pandas as pd
import time
import xlsxwriter

# read file
def read_data(file_name):
    try:
        df_detail = pd.DataFrame(pd.read_excel(io=file_name, sheet_name='payroll', header=0, skiprows=2))
        df_si = pd.DataFrame(pd.read_excel(io=file_name, sheet_name='si', header=0, skiprows=2))
        df_conf = pd.DataFrame(pd.read_excel(io=file_name, sheet_name='conf', header=0, skiprows=0,index_col=0))

        return df_detail.fillna(0),df_si,df_conf.fillna(0)
    except Exception as e:
        print('read file failed:', file_name)
        print('error log', e)


def get_path():
    return "C:/temp/90own/"


def gl_summary_detail(df_data,df_si,df_conf,out_columns):
    file_name = get_path() + time.strftime("%Y-%m-%d",time.localtime()) + "_all_in_one_gl_report.xlsx"

    wage_list = []
    wage_used = []
    df_res = pd.DataFrame(columns=(out_columns))
    df_res2 = pd.DataFrame()

    wage_list = df_conf.index.tolist()
    # print(type(wage_list))

    df_data_columns = df_detail.columns.values.tolist()
    df_si_columns= df_si.columns.values.tolist()
    df_log_amt = pd.DataFrame()
    # print('Table columns:',df_data_columns)

    for i in range(len(wage_list)):
        if wage_list[i] in df_data_columns:
            wage_used.append(wage_list[i])
        elif wage_list[i] in df_si_columns:
            wage_used.append(wage_list[i])
    # print('Handled wages:',str(wage_used))
    df_wages = pd.DataFrame(wage_used, columns=['Handled Wages'])

    # split according to the files
    for n_idx,n_row in df_data.iterrows():
        if n_row['PayMonth'] != 0:
        # print(n_row['Employee No.'],)
        # print(n_row['AltiumStaffNumber'])
        # print(n_row['Employee No.'],)
            for i in range(len(wage_used)):
                if wage_used[i] in df_data_columns:

                    line_rs = {'AltiumStaffNumber':n_row['AltiumStaffNumber'],'PayMonth':n_row['PayMonth'],'Location':n_row['Location'],'CostCenter':n_row['CostCenter']}

                    line_conf = pd.Series(df_conf.loc[wage_used[i]]).to_dict()
                    # print(wage_used[i],line_conf)
                    line_rs['Wages'] = wage_used[i]
                    if line_conf['GLWages'] == 0 or line_conf['GLWages'] == "":
                        line_rs['GLWages'] = wage_used[i]
                    else:
                        line_rs['GLWages'] = line_conf['GLWages']

                    line_rs['GLWages-SEQ'] = str(line_conf['Sequence'])+'_'+line_rs['GLWages']

                    line_rs['GL Account'] = line_conf['GL Account']
                # print(wage_used[i])
                    if line_conf['Credit'] == 'X':
                        line_rs['Credit'] = n_row[wage_used[i]]

                    if line_conf['Debit'] == 'X':
                        line_rs['Debit'] = n_row[wage_used[i]]

                    if n_row[wage_used[i]] != 0 :
                        df_res = df_res.append(pd.Series(line_rs),ignore_index=True)
            line_conf = {}
            line_rs={}


    for n_idx,n_row in df_si.iterrows():
        if n_row['PayMonth'] != 0:
        # print(n_row['Employee No.'],)
            for i in range(len(wage_used)):
                if wage_used[i] in df_si_columns:
                    line_rs = {'AltiumStaffNumber':n_row['AltiumStaffNumber'],'PayMonth':n_row['PayMonth'],'Location':n_row['Location'],'CostCenter':n_row['CostCenter']}

                    # try:
                    #     line_master = pd.Series(df_data[df_data['AltiumStaffNumber'] == str(int(n_row['AltiumStaffNumber']))].loc[0]).to_dict()
                    #
                    #     print(line_master['Location'])
                    #     print(line_master['CostCenter'])
                    # except Exception as e:
                    #     print("Error in handling employee ",n_row['AltiumStaffNumber'],e)

                    line_conf = pd.Series(df_conf.loc[wage_used[i]]).to_dict()
                    # print(wage_used[i],line_conf)
                    line_rs['Wages'] = wage_used[i]
                    if line_conf['GLWages'] == 0 or line_conf['GLWages'] == "":
                        line_rs['GLWages'] = wage_used[i]
                    else:
                        line_rs['GLWages'] = line_conf['GLWages']
                    line_rs['GLWages-SEQ'] = str(line_conf['Sequence'])+'_'+line_rs['GLWages']

                    line_rs['GL Account'] = line_conf['GL Account']


                    # print(wage_used[i])
                    if line_conf['Credit'] == 'X':
                        line_rs['Credit'] = n_row[wage_used[i]]

                    if line_conf['Debit'] == 'X':
                        line_rs['Debit'] = n_row[wage_used[i]]

                    if n_row[wage_used[i]] != 0 :
                        df_res = df_res.append(pd.Series(line_rs),ignore_index=True)
            line_conf = {}
            line_rs={}

    df_res = df_res.fillna(0)
    print("finished processing\n",df_res.head(3))

    df_detail_gl = df_res.groupby(['AltiumStaffNumber','PayMonth','Location','CostCenter','GLWages-SEQ','GL Account','GLWages'])['Credit','Debit'].sum().reset_index()
    df_detail_person = df_res.groupby(['AltiumStaffNumber','PayMonth','Location','CostCenter'])['Credit','Debit'].sum().reset_index()
    # print('Results\n',df_detail_gl.head())

    df_log_amt = df_res.groupby(['Wages'])['Credit','Debit'].sum().reset_index()

    df_res = df_res.groupby(['PayMonth','Location','CostCenter','GLWages-SEQ','GL Account','GLWages'])['Credit','Debit'].sum().reset_index()
    # df_res = df_res.sort_values(by=['PayMonth','Location','CostCenter','GLWages-SEQ','GL Account','GLWages'],ascending=False).reset_index()
    df_res = df_res.sort_values(by=['PayMonth','Location','CostCenter','GLWages-SEQ','GL Account','GLWages']).reset_index()
    df_res = df_res[['PayMonth','Location','CostCenter','GL Account','GLWages','Credit','Debit']]
    # df_res = df_res.sort_values(['PayMonth','Location','CostCenter','GL Account'],inplace=True)
    # for l_loc in ['BenQ',]
    print('Results\n',df_res.head())

    df_res2 = df_res.groupby(['PayMonth','CostCenter','GL Account','GLWages'])['Credit','Debit'].sum().reset_index()

    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(file_name,engine='xlsxwriter')
        workbook = df_writer.book

        sheet_name = 'Final_report'
        df_res.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)


        sheet_name = 'Cost_Center'
        df_res2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'handled_wages'
        df_wages.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'log_amt'
        df_log_amt.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'gl_detail_report'
        df_detail_gl.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)


        sheet_name = 'gl_person'
        df_detail_person.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        workbook.close()
    except Exception as e:
        print('write file failed:', file_name)
        print('error log', e)




if __name__ == '__main__':

    # get file path
    line_columns = ('AltiumStaffNumber','PayMonth','Location', 'CostCenter',  'GL Account', 'GLWages-SEQ','Wages','GLWages','Credit', 'Debit')

    time1 = time.time()

    # file_in = root_dir +  'altium_gl_data_source_v2.xlsx'
    file_in = get_path() +  'Altium Payroll Report - 201909-v2-gl.xlsx'

    df_detail, df_si, df_conf = read_data(file_in)

    print('conf data\n', df_conf.head(1))
    # print('detail data\n', df_detail.head(2))
    gl_summary_detail(df_detail,df_si,df_conf,line_columns)

    # summary_data(output_root,df_detail)
    time2 = time.time()

    print("Total running time", time2 - time1 )
