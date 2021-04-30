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


def get_columns():
    return ['Country','Company','Division','BU','RU','Job','EMGroup','Status','Location','Month']


def get_hire_chg(l_date, l_reason):
    try:

        if pd.isnull(l_date):
            return 'Unknown'
        else:
            l_dateDay = pd.to_datetime(l_date)
            # if l_reason == 'New Hire' or l_reason == 'Rehire':
            # if l_date != '2019-Ing-Ing' and l_date != '2019-03-08':
            if l_date != '2019-Ing-Ing' and l_date != '2019-03-08' and l_dateDay.year >= 2019:
                return (str(l_dateDay)[0:7])
            else:
                return 'Unknown'
    except Exception as e:
        print('Convert hire date failed:', l_date)
        print('error log', e)
        return 'Unknown'


def get_noshow_chg(l_date, l_reason):
    try:
        l_dateDay = pd.to_datetime(l_date)
        if l_reason == 'Terminated - No Show' :
            return (str(l_dateDay)[0:7] if pd.notnull(str(l_dateDay)[0:7] ) else 'Unknown' )
        else:
            return 'Unknown'
    except Exception as e:
        print('Convert no show date failed:', l_date)
        print('error log', e)
        return 'Unknown'


def get_leave_chg(l_date, l_reason):
    try:
        if pd.isnull(l_date):
            return 'Unknown'
        l_dateDay = pd.to_datetime(l_date)
        # if l_reason != 'Terminated - No Show':
        if l_reason != 'Terminated - No Show' and l_dateDay.year <= 2019:
            return (str(l_dateDay)[0:7])
        else:
            return 'Unknown'
    except Exception as e:
        print('Convert terminated date failed:', l_date)
        print('error log', e)
        return 'Unknown'


def get_root():
    return 'C:/temp/analyze/'


def pre_process(input):
    df = pd.DataFrame()
    file = get_root() + input
    try:
        df = pd.DataFrame(pd.read_excel(io=file, sheet_name='Excel Output', header=0, skiprows=2))
        # df = df[df['Employee Status (Label)'] == 'Active']
        # df['AgeRange'] = df.apply(lambda x:get_age_range(x['Date Of Birth']),axis=1)
        # df['Hire_Month'] = df.apply(lambda x:get_hire_chg(x['Hire Date'],x['Event Reason Icode (Label)']),axis=1)
        df['Hire_Month'] = df.apply(lambda x:get_hire_chg(x['Hire Date'],x['Event Reason Icode (Label)']),axis=1)
        df['Hire_No'] = df.apply((lambda x: 1 if x['Hire_Month'] != 'Unknown' else 0 ),axis=1)
        df['Leave_Month'] = df.apply(lambda x:get_leave_chg(x['Termination Date'],x['Event Reason Icode (Label)']),axis=1)
        df['Leave_No'] = df.apply((lambda x: 1 if x['Leave_Month'] != 'Unknown' else 0 ),axis=1)
        df['No_Show_Month'] = df.apply(lambda x:get_noshow_chg(x['Termination Date'],x['Event Reason Icode (Label)']),axis=1)
        df['Noshow_No'] = df.apply((lambda x: 1 if x['No_Show_Month'] != 'Unknown' else 0 ),axis=1)
    except Exception as e:
        print('Exception:', file,e)
    # print(df.columns)
    return df.fillna('')


def out_data(df_data,df_writer,workbook,chart2,sheet_name,title, xtitle,ytitle):
    df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
    worksheet = df_writer.sheets[sheet_name]
    x, y = df_data.shape
    print('df_datagroup', x, y)
    # chart2 = workbook.add_chart({'type': 'column','subtype': 'stacked'})        #'subtype': 'percent_stacked'
    chart2 = workbook.add_chart({'type': 'column'})  # 'subtype': 'percent_stacked'
    for i in range(1, y):
        chart2.add_series({
            'name': [sheet_name, 0, i],
            'categories': [sheet_name, 1, 0, x, 0],
            'values': [sheet_name, 1, i, x, i],
        })
    # Add a chart title and some axis labels.
    chart2.set_title({'name': title})
    chart2.set_x_axis({'name': xtitle})
    chart2.set_y_axis({'name': 'Employee Number'})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart2.set_style(10)
    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart(10, 2, chart2, {'x_offset': 50, 'y_offset': 100})


def simple_analyze(df_data,file_out):
    df_simple = pd.DataFrame()
    df_simple_T = pd.DataFrame()
    df_datagroup = pd.DataFrame()
    line_rs = {}
    df_simple = df_data.groupby(['Country (ID)','Company (Label)','Employee Status (Label)','Division Short Text','BU Short Text','Reporting Unit (Reporting Unit ID)','Employment Type (Label)','Location Group (Name)','Job Classification (Job Code)',
                                 'Hire_Month','Leave_Month','No_Show_Month'])['EE Count','Hire_No','Leave_No','Noshow_No'].sum().reset_index()
    print('df_simple:','columns',df_simple.shape[0],'\n',df_simple.head(2))

    # get detail data
    for n_idx,n_row in df_simple.iterrows():
        line_rs={}
        # line_rs['Country'] = n_row['Country (ID)']
        # line_rs['Company'] = n_row['Company (Label)']
        # line_rs['Status'] = n_row['Employee Status (Label)']
        # line_rs['Division'] = n_row['Division Short Text']
        # line_rs['BU'] = n_row['BU Short Text']
        # line_rs['RU'] = n_row['Reporting Unit (Reporting Unit ID)']
        # line_rs['EMGroup'] = n_row['Employment Type (Label)']
        # line_rs['Location'] = n_row['Location Group (Name)']
        # line_rs['Job'] = n_row['Job Classification (Job Code)']
        # line_rs['Status'] = n_row['Employee Status (Label)']
        if n_row['Hire_Month'] !=  'Unknown':
            line_rs['Month'] = n_row['Hire_Month']
            line_rs['Hire_No'] = n_row['Hire_No']
            df_simple_T = df_simple_T.append(pd.Series(line_rs),ignore_index=True)
        if n_row['Leave_Month'] != 'Unknown':
            line_rs['Month'] = n_row['Leave_Month']
            line_rs['Leave_No'] = n_row['Leave_No']
            line_rs['Hire_No'] = 0
            line_rs['Noshow_No'] = 0
            df_simple_T = df_simple_T.append(pd.Series(line_rs),ignore_index=True)
        if n_row['No_Show_Month'] != 'Unknown':
            line_rs['Month'] = n_row['No_Show_Month']
            line_rs['Noshow_No'] = n_row['Noshow_No']
            line_rs['Leave_No'] = 0
            line_rs['Hire_No'] = 0
            df_simple_T = df_simple_T.append(pd.Series(line_rs), ignore_index=True)

    t_columns = ['Month','Hire_No','Leave_No','Noshow_No']
    df_simple_T = df_simple_T[t_columns]
    df_simple_T.sort_values('Month',inplace=True)

    df_dg_columns = df_simple_T.columns.tolist()
    print('df_columns',df_dg_columns)
    # for fld in get_columns():
    #     df_dg_columns.remove(fld)
    df_dg_columns.remove('Month')
    df_dg_columns.sort()
    print('data group columns sort',df_dg_columns)

    df_datagroup = df_simple_T.groupby('Month')[df_dg_columns].sum().reset_index()
    df_datagroup = df_datagroup.fillna(0).sort_values('Month')
    # df_datagroup = df_datagroup[df_datagroup['Month'] >= '2019-Ing' and df_datagroup['Month'] <= '2019-06']

    df_status = df_simple.groupby('Employee Status (Label)')['EE Count'].sum().reset_index()
    df_status = df_status.fillna(0)

    # df_company = df_simple_T.groupby('Company')[df_dg_columns].sum().reset_index().fillna(0).sort_values('EMP_NO',ascending=False)
    # df_chg = df_simple_T.groupby('Company')[df_dg_columns].sum().reset_index().fillna(0).sort_values('EMP_NO',ascending=False)


    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(get_root()+file_out,engine='xlsxwriter')
        workbook = df_writer.book

        # sheet_name = 'df_data'
        # df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'Simple'
        df_simple.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'Simple_T'
        df_simple_T.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '10-Changes'
        df_datagroup.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        x,y = df_datagroup.shape
        print('df_datagroup',x,y)
        # chart2 = workbook.add_chart({'type': 'column','subtype': 'stacked'})        #'subtype': 'percent_stacked'
        chart2 = workbook.add_chart({'type': 'line'})        #'subtype': 'percent_stacked'
        for i in range(1,y):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
            })
        # Add a chart title and some axis labels.
        chart2.set_title({'name': 'Employee Status Changes'})
        chart2.set_x_axis({'name': 'Month'})
        chart2.set_y_axis({'name': 'Employee Number'})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_style(10)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(10,2, chart2, {'x_offset': 50, 'y_offset': 100})

        sheet_name = '11-Status'
        chart2 = {}
        df_status.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        x,y = df_status.shape
        print('df_datagroup',x,y)
        chart2 = workbook.add_chart({'type': 'pie',})        #'subtype': 'percent_stacked'
        # chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'
        for i in range(1,y):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
            })
        # Add a chart title and some axis labels.
        chart2.set_title({'name': 'Employee status'})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_style(10)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(3,3, chart2, {'x_offset': 50, 'y_offset': 100})

        # # out_data(df_data, df_writer, workbook, chart2, sheet_name, title, xtitle,ytitle):
        # out_data(df_company, df_writer, workbook, chart2, '31-company_no', 'Employee No By Company', 'Company','Employee No')
        # out_data(df_ru, df_writer, workbook, chart2, '34-ru_no', 'Employee No By RUs', 'RUs','Employee No')

        workbook.close()
    except Exception as e:
        print('write file failed:', file_out)
        print('error log', e)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    file_in = 'AP_EmployeeBasicInfov2light-20190924.xlsx'
    file_out = 'Output_'+ now_date + '_t3_basic_info_all_status_chg.xlsx'
    time1 = time.time()

    df_data = pre_process(file_in)
    simple_analyze(df_data, file_out)

    print("Total running time", time.time() - time1)
