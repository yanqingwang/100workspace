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
import os

def get_direct(external, employee_type, employee_class,l_global_id):
    try:
        if external != 'Yes' and employee_class == 'Direct' and employee_type not in ["Intern/Students","Apprentice","Vacation Workers DE",""]:
            return 1
        else:
            return 0
    except Exception as e:
        print('get direct employee error:', l_global_id)
        print('error log', e)
        return 0


def get_indirect(external, employee_type, employee_class,l_global_id):
    try:
        if external != 'Yes' and employee_class == 'In-direct/Salaried' and employee_type not in ["Intern/Students","Apprentice","Vacation Workers DE",""]:
            return 1
        else:
            return 0
    except Exception as e:
        print('get direct employee error:', l_global_id)
        print('error log', e)
        return 0


def get_region(country):
    try:
        if country in ["ARE","AUS","CHN","JPN","KOR","MYS","PHL","SGP","THA","TWN","VNM","IDN","IND"]:
            return "AP"
        else:
            return "Unknown"
    except Exception as e:
        print('error log to get country', e)


def prepare_path():
    root = os.path.abspath('..') + '/testdata/10headcount/'
    if not os.path.exists(root):
        os.mkdir(root)
    tmp_path = root + 'rs'
    if not os.path.exists(tmp_path):
        os.mkdir(tmp_path)


def get_path():
    root = os.path.abspath('..') + '/testdata/10headcount/'
    # print(root)
    return  root


def out_chart(df_data, df_writer, workbook, chart2, sheet_name, title, xtitle, ytitle):
    df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
    worksheet = df_writer.sheets[sheet_name]
    x, y = df_data.shape
    print('df_datagroup', x, y)
    # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
    chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
    for i in range(1, y):
        chart2.add_series({
            'name': [sheet_name, 0, i],
            'categories': [sheet_name, 1, 0, x, 0],
            'values': [sheet_name, 1, i, x, i],
        })
    # Add a chart title and some axis labels.
    chart2.set_title({'name': title})
    chart2.set_x_axis({'name': xtitle})
    chart2.set_y_axis({'name': ytitle})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart2.set_style(10)
    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart(10, 2, chart2, {'x_offset': 50, 'y_offset': 100})


def pre_process(file_start):
    df = pd.DataFrame()

    file = ""
    for filename in os.listdir(get_path()):
        if filename.startswith(file_start):
            file = get_path() + filename

    try:
        df = pd.DataFrame(pd.read_excel(io=file, sheet_name='Excel Output', header=0, skiprows=2))
        print(df.head(2))
        df = df.rename(columns={"External Agency & Contingent Worker": "External_worker",
                                "Reporting Unit (Reporting Unit ID)": "RU",
                                "Division/Corporate Function/Region (Label)": "Division Text",
                                "Regular/Limited Employment (Label)": "EmploymentType_Text",
                                "Regular/Limited Employment (External Code)":"EmploymentType"})

        df['External'] = df.apply((lambda x: 1 if x['External_worker'] == 'Yes' else 0 ),axis=1)
        df['Intern'] = df.apply((lambda x: 1 if x['Employment Type (Label)'] == 'Intern/Students' else 0 ),axis=1)
        df['Apprentices'] = df.apply((lambda x: 1 if x['Employment Type (Label)'] == 'Apprentice' else 0 ),axis=1)
        df['VacationWorker'] = df.apply((lambda x: 1 if x['Employment Type (Label)'] == 'Vacation Workers DE' else 0 ),axis=1)

        df['AP'] = df.apply(lambda x:get_region(x['Country (ID)']),axis=1)
        df['EEDirect'] = df.apply(lambda x:get_direct(x['External_worker'],x['Employment Type (Label)'],x['Employee Class (Label)'],x['ZF Global ID']),axis=1)
        df['EEIndirect'] = df.apply(lambda x:get_indirect(x['External_worker'],x['Employment Type (Label)'],x['Employee Class (Label)'],x['ZF Global ID']),axis=1)
        # df['External'] = df.apply(lambda x:get_hire_chg(x['External Agency & Contingent Worker'],x['Event Reason Icode (Label)']),axis=1)
        print(df.columns)
    except Exception as e:
        print('Exception:', file,e)
    # print(df.columns)
    return df.fillna('')


def head_count_summary(df_data,file_out):
    df_simple = pd.DataFrame()
    df_country = pd.DataFrame()
    df_company = pd.DataFrame()
    df_ru = pd.DataFrame()

    df_2simple = pd.DataFrame()
    df_2country = pd.DataFrame()
    df_2company = pd.DataFrame()
    df_2ru = pd.DataFrame()

    df_data['EETotal'] = df_data['EEDirect'] + df_data['EEIndirect']

    # All
    df_simple = df_data.groupby(['Country (ID)','Company (Label)','RU'])['EETotal','EEDirect','EEIndirect','External','Intern','Apprentices','VacationWorker'].sum().reset_index()
    df_simple = df_simple.sort_values("EETotal", ascending=False)
    print('df_simple:','columns',df_simple.shape[0],'\n',df_simple.head(2))
    df_country = df_simple.groupby(['Country (ID)'])['EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker'].sum().reset_index().sort_values("EETotal", ascending=False)
    df_company = df_simple.groupby(['Company (Label)'])['EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker'].sum().reset_index().sort_values("EETotal", ascending=False)
    df_ru = df_simple.groupby(['RU'])['EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker'].sum().reset_index().sort_values("EETotal", ascending=False)
    # AP only
    df_2simple = df_data[df_data["AP"] == "AP"].groupby(['Country (ID)','Company (Label)','RU'])['EETotal','EEDirect','EEIndirect','External','Intern','Apprentices','VacationWorker'].sum().reset_index().sort_values(["Country (ID)","EETotal"], ascending=False)
    df_2country = df_2simple.groupby(['Country (ID)'])['EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker'].sum().reset_index().sort_values("EETotal", ascending=False)
    df_2company = df_2simple.groupby(['Company (Label)'])['EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker'].sum().reset_index().sort_values("EETotal", ascending=False)
    df_2ru = df_2simple.groupby(['RU'])['EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker'].sum().reset_index().sort_values("EETotal", ascending=False)

    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(get_path()+file_out,engine='xlsxwriter')
        workbook = df_writer.book
        # chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'

        sheet_name = '10_Initial'
        df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '10_overall'
        df_simple.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '10_country'
        chart2 = {}
        out_chart(df_country, df_writer, workbook, chart2,  sheet_name, 'Employee No By Country', 'Country', 'Employee No')
        # df_country.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '10_company'
        df_company.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '10_RU'
        df_ru.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '20_overall'
        df_2simple.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '20_country'
        chart2 = {}
        out_chart(df_2country, df_writer, workbook, chart2,  sheet_name, 'Employee No', 'Country', 'Employee No')
        # df_2country.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '20_company'
        chart2 = {}
        out_chart(df_2company, df_writer, workbook, chart2,  sheet_name, 'Employee No', 'Country', 'Employee No')
        # df_2company.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = '20_RU'
        chart2 = {}
        out_chart(df_2ru, df_writer, workbook, chart2,  sheet_name, 'Employee No', 'Country', 'Employee No')
        # df_2ru.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        workbook.close()
    except Exception as e:
        print('write file failed:', file_out)
        print('error log', e)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    file_start = 'Input'
    file_out = 'Output_'+ now_date+ '_headcout_output.xlsx'
    time1 = time.time()

    df_data = pre_process(file_start)
    head_count_summary(df_data, file_out)

    print("Total running time", time.time() - time1)
