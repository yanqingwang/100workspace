# -*- coding: utf-8 -*-
"""
This is a script to read files and updated as pdf files.
"""

from os import chdir, listdir
from datetime import date
import time
import datetime
import pandas as pd
import xlsxwriter
import os

from win32com.client import gencache, DispatchEx

def get_path():
    return 'c:/Users/z659190/Documents/10 Work/10 MyHRSuit/10 Project/131 learning/30_Workpackages/37 Data migration/1st Submission/'


def set_columns():
    return ['Global ID','Local ID', 'Learner Name', 'Item/ Program Name','Training Hours','Item Type','Start Date','End Date',
                           'Completion date','Expiration Date for Certifications','Vendor/ Instructor','Comments/ Remarks']

def get_str_date(l_date):
    if len(str(l_date)) > 10:
        return (str(l_date)[:10])
    else:
        return str(l_date)


def read_file():
    df_data_tmp = pd.DataFrame()
    df_data = pd.DataFrame()
    filenames = listdir(get_path())
    for fname in filenames:
        # if fname.startswith('p'):
        if fname.startswith('Con'):
            try:
                f_name = get_path() + fname
                sheet_to_df_map = pd.read_excel(f_name, sheet_name=None, skip_blank_lines=True, parse_dates=False)
                # sheets = pd.ExcelFile(fname)
                # # this will read the first sheet into df
                # for l_sheet in sheets:
                #     df_data = pd.DataFrame(pd.read_excel(io=fname, sheet_name=l_sheet, header=0, skiprows=0))
                # print(sheet_to_df_map.keys())
                for key in sheet_to_df_map.keys():
                    if not key in ['Notes', 'Format- various systems', 'Questions']:
                        df_data_tmp = sheet_to_df_map[key]
                        df_data_tmp.columns = set_columns()
                        df_data = df_data.append(df_data_tmp, sort=False)
                print(df_data.head())

            except Exception as e:
                print('read file failed:', fname)
                print('error log', e)
    df_data = df_data.sort_values('Global ID')
    print(df_data.columns)
    return df_data.fillna("")


def output_data(df_data, file_out):
    df_ru = pd.DataFrame()

    # df_data['EETotal'] = df_data['EEDirect'] + df_data['EEIndirect']

    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(get_path() + file_out, engine='xlsxwriter')
        workbook = df_writer.book
        # chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'
        # df_data = df_data.sort_values("Global ID", ascending=False)
        # df_data = df_data.sort_values('Global ID',inplace=True)
        sheet_name = 'Sheet1'
        df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        workbook.close()
    except Exception as e:
        print('write file failed:', file_out)
        print('error log', e)


def gen_pdf_from_xlsx(df_data):
    df_temp = pd.DataFrame()

    list_globalid = list(set(df_data['Global ID'].to_list()))
    print(list_globalid)

    df_data['Start Date'] = df_data.apply(lambda x:get_str_date(x['Start Date']),axis=1)
    df_data['End Date'] = df_data.apply(lambda x:get_str_date(x['End Date']),axis=1)

    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = 0
    for gid in list_globalid:

        df_temp = df_data[df_data['Global ID'] == gid]
        try:
            filename = get_path() + 'xlsx_tmp/p' + str(gid) + '.xlsx'
            # 创建一个excel
            df_writer = pd.ExcelWriter(filename, engine='xlsxwriter',date_format='yyyy-mm-dd')
            workbook = df_writer.book
            sheet_name = 'Sheet1'
            df_temp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False,header=False,startrow=1)
            worksheet = df_writer.sheets[sheet_name]
            worksheet.set_default_row(hide_unused_rows=True)
            # print(x,y)

            # Add a header format.
            header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC',
                                                 'font_size': 10, 'border': 1})

            common_format = workbook.add_format({'border': True,'font_size':10,'align': 'left', 'text_wrap': True})

            # Write the column headers with the defined format.
            for col_num, value in enumerate(df_temp.columns.values):
                worksheet.write(0, col_num, value)
            worksheet.set_row(0,height=None,cell_format=header_format)
            worksheet.set_column("C:C",18, cell_format=common_format)
            worksheet.set_column("A:A",10, cell_format=common_format)
            worksheet.set_column("B:B",10, cell_format=common_format)
            worksheet.set_column("D:D",32, cell_format=common_format)
            worksheet.set_column("E:E",6, cell_format=common_format)
            worksheet.set_column("F:F",10, cell_format=common_format)
            worksheet.set_column("G:H",12, cell_format=common_format)
            worksheet.set_column("I:J",10, cell_format=common_format)
            worksheet.set_column("K:K",16, cell_format=common_format)
            worksheet.set_column("L:L",16, cell_format=common_format)


            worksheet.set_paper(9)      #https://xlsxwriter.readthedocs.io/page_setup.html
            worksheet.set_margins(left=0.2,right=0.2, top=0.3, bottom=0.5)
            # (x,y) = df_temp.shape
            # worksheet.print_area(0,0,x-1,y-1)
            worksheet.fit_to_pages(1,1)     #
            worksheet.set_landscape()       #worksheet.set_portrait()
            workbook.close()
            df_writer.save()
            df_writer.close()

            path = 'C:/temp/pdf/'
            file_pdf = path + 'p' + str(gid) + '.pdf'
            if os.access(file_pdf, os.F_OK):
                os.remove(file_pdf)
            books = xlApp.Workbooks.Open(filename, False)
            books.ExportAsFixedFormat(0, file_pdf)  #0 is pdf,1 is xps
            books.Close(False)
            print('Save PDF Files：', file_pdf)

        except Exception as e:
            print('write file failed:', file_out)
            print('error log', e)

    xlApp.Quit()


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    file_out = 'Output_' + now_date + '_Learning_history_data.xlsx'
    time1 = time.time()

    chdir(get_path())

    df_data = read_file()
    output_data(df_data, file_out)
    gen_pdf_from_xlsx(df_data)

    print("Total running time", time.time() - time1)
