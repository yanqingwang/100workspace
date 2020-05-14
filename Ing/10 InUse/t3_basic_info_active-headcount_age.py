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
    return ['Country','Company','Board','Division','BU','RU','Department','Job','EMGroup','Location','SolidM','BP']


def get_age_range(l_date):
    try:
        l_dateDay = pd.to_datetime(l_date)
        today = date.today()
        age = today.year - l_dateDay.year - ((today.month, today.day) < (l_dateDay.month, l_dateDay.day))
        if age <= 20:
            return '0--20'
        elif age <= 30:
            return '25--30'
        elif age <=35:
            return '30--35'
        elif age <=40:
            return '35--40'
        elif age <=50:
            return '40--50'
        elif age <= 60:
            return '50--60'
        else:
            return '60--100'

    except Exception as e:
        print('Convert date failed:', l_date)
        print('error log', e)
        return 'Unkown'


def get_root():
    return 'C:/temp/analyze/'


def pre_process(input):
    df = pd.DataFrame()
    file = get_root() + input
    try:
        df = pd.DataFrame(pd.read_excel(io=file, sheet_name='Excel Output', header=0, skiprows=2))
        df = df[df['Employee Status (Label)'] == 'Active']
        # df = pd.DataFrame(pd.read_excel(io=input, sheet_name='Excel Output', header=0, skiprows=2))
        df['AgeRange'] = df.apply(lambda x:get_age_range(x['Date of Birth']),axis=1)
        # df['EventReason'] = df.apply((lambda x:'OTHERS' if (pd.isnull(x.EventReason)) else x.EventReason),axis=1)
    except Exception as e:
        print('Exception:', file,e)
    # print(df.columns)
    return df


def out_data(df_data,df_writer,workbook,chart2,sheet_name,title, xtitle,ytitle):
    df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
    worksheet = df_writer.sheets[sheet_name]
    x, y = df_data.shape
    print('df_datagroup', x, y)
    # chart2 = workbook.add_chart({'type': 'column','subtype': 'stacked'})        #'subtype': 'percent_stacked'
    chart2 = workbook.add_chart({'type': 'column'})  # 'subtype': 'percent_stacked'
    for i in range(1, y):
        chart2.add_series({
            # 'name': [sheet_name, 0, i],
            'categories': [sheet_name, 1, 0, x, 0],
            'values': [sheet_name, 1, i, x, i],
        })
    # Add a chart title and some axis labels.
    chart2.set_title({'name': title})
    chart2.set_x_axis({'name': xtitle})
    chart2.set_y_axis({'name': 'Employee Number'})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart2.set_style(16)
    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart(10, 2, chart2, {'x_offset': 50, 'y_offset': 100})


def simple_analyze(df_data,file_out):
    df_simple = pd.DataFrame()
    df_simple_T = pd.DataFrame()
    df_datagroup = pd.DataFrame()
    line_rs = {}
    df_simple = df_data.groupby(['Country (ID)','Company (Label)','Board Short Text','Division Short Text','BU Short Text','Department Short Text','Reporting Unit (Reporting Unit ID)','Employment Type (Label)','Location Group (Name)','Job Classification (Job Code)','Solid Line Manager Global ID','BP Global ID','AgeRange'])['10ZF Global ID'].count().reset_index()
    # df_simple = df_data.groupby(['Country (ID)','Company (Label)','Board Short Text','Division Short Text','BU Short Text','Employment Type (Label)','Job Classification (Job Code)','AgeRange'])['10ZF Global ID'].count().reset_index()
    print('df_simple:','columns',df_simple.shape[0],'\n',df_simple.head(2))

    # get detail data
    for n_idx,n_row in df_simple.iterrows():
        line_rs={}
        line_rs['Country'] = n_row['Country (ID)']
        line_rs['Company'] = n_row['Company (Label)']
        line_rs['Board'] = n_row['Board Short Text']
        line_rs['Division'] = n_row['Division Short Text']
        line_rs['BU'] = n_row['BU Short Text']
        line_rs['Department'] = n_row['Department Short Text']
        line_rs['RU'] = n_row['Reporting Unit (Reporting Unit ID)']
        line_rs['EMGroup'] = n_row['Employment Type (Label)']
        line_rs['Location'] = n_row['Location Group (Name)']
        line_rs['Job'] = n_row['Job Classification (Job Code)']
        line_rs['SolidM'] = n_row['Solid Line Manager Global ID']
        line_rs['BP'] = n_row['BP Global ID']
        line_rs[n_row['AgeRange']] = n_row['10ZF Global ID']
        line_rs['EMP_NO'] = n_row['10ZF Global ID']
        df_simple_T = df_simple_T.append(pd.Series(line_rs),ignore_index=True)

        # line_rs[n_row['EventReason']] = n_row['GlobalID']

    df_dg_columns = df_simple_T.columns.tolist()
    print('df_columns',df_dg_columns)
    for fld in get_columns():
        df_dg_columns.remove(fld)
    df_dg_columns.remove('EMP_NO')
    df_dg_columns.sort()
    print('data group columns sort',df_dg_columns)

    df_datagroup = df_simple_T.groupby('Country')[df_dg_columns].sum().reset_index()
    df_datagroup = df_datagroup.fillna(0)

    df_employee_group = df_simple_T.groupby('EMGroup')[df_dg_columns].sum().reset_index()
    df_employee_group = df_employee_group.fillna(0)

    df_country = df_simple_T.groupby('Country')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_company = df_simple_T.groupby('Company')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_company = df_company.sort_values("EMP_NO",ascending=False)
    df_Board = df_simple_T.groupby('Board')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_division = df_simple_T.groupby('Division')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_BU = df_simple_T.groupby('BU')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_Job = df_simple_T.groupby('Job')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_Location = df_simple_T.groupby('Location')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_ru = df_simple_T.groupby('RU')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    # df_department = df_simple_T.groupby('Department')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_solidM = df_simple_T.groupby('SolidM')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_solidM = df_solidM.sort_values('SolidM',ascending=False).sort_values("EMP_NO",ascending=False)
    df_BP = df_simple_T.groupby('BP')['EMP_NO'].sum().reset_index().fillna(0).sort_values("EMP_NO",ascending=False)
    df_BP = df_BP.sort_values('BP',ascending=False).sort_values("EMP_NO",ascending=False)


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

        sheet_name = '10-country_age'
        df_datagroup.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        x,y = df_datagroup.shape
        print('df_datagroup',x,y)
        # chart2 = workbook.add_chart({'type': 'column','subtype': 'stacked'})        #'subtype': 'percent_stacked'
        chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'
        for i in range(1,y):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
            })
        # Add a chart title and some axis labels.
        chart2.set_title({'name': 'Ages Factors By Country'})
        chart2.set_x_axis({'name': 'Company'})
        chart2.set_y_axis({'name': 'Employee Number'})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_style(16)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(10,2, chart2, {'x_offset': 50, 'y_offset': 100})

        sheet_name = '11-Emp_Group_age'
        chart2 = {}
        df_employee_group.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        x,y = df_employee_group.shape
        print('df_datagroup',x,y)
        chart2 = workbook.add_chart({'type': 'column','subtype': 'stacked'})        #'subtype': 'percent_stacked'
        # chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'
        for i in range(1,y):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
            })
        # Add a chart title and some axis labels.
        chart2.set_title({'name': 'Ages Factors By Employment Type'})
        chart2.set_x_axis({'name': 'Company'})
        chart2.set_y_axis({'name': 'Employee Number'})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_style(16)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(10,2, chart2, {'x_offset': 50, 'y_offset': 100})

        # out_data(df_data, df_writer, workbook, chart2, sheet_name, title, xtitle,ytitle):
        out_data(df_country, df_writer, workbook, chart2, '30-country_no', 'Employee No By Country', 'Country','Employee No')
        out_data(df_company, df_writer, workbook, chart2, '31-company_no', 'Employee No By Company', 'Company','Employee No')
        out_data(df_Board, df_writer, workbook, chart2, '31-board_no', 'Employee No By Board', 'Board','Employee No')
        out_data(df_division, df_writer, workbook, chart2, '32-division_no', 'Employee No By Division', 'Division','Employee No')
        out_data(df_BU, df_writer, workbook, chart2, '33-bu_no', 'Employee No By BU', 'BUs','Employee No')
        out_data(df_ru, df_writer, workbook, chart2, '34-ru_no', 'Employee No By RUs', 'RUs','Employee No')
        # out_data(df_department, df_writer, workbook, chart2, '35-department_no', 'Employee No By Department', 'Department','Employee No')
        out_data(df_Location, df_writer, workbook, chart2, '38-location_no', 'Employee No By Location', 'Locations','Employee No')
        out_data(df_Job, df_writer, workbook, chart2, '39-job_no', 'Employee No By Jobs', 'Jobs','Employee No')

        sheet_name = 'Solid_MGR'
        df_solidM.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'BP'
        df_BP.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        workbook.close()
    except Exception as e:
        print('write file failed:', file_out)
        print('error log', e)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    # file_in = 'EmployeeBasicInfov2all-Page1-20190802.xlsx'
    file_in = 'AP_EmployeeBasicInfov2light-20190924.xlsx'
    file_out = 'Output_' + now_date + 'All_basic_info_active_analyze_headcount.xlsx'
    time1 = time.time()

    df_data = pre_process(file_in)
    simple_analyze(df_data, file_out)

    print("Total running time", time.time() - time1)
