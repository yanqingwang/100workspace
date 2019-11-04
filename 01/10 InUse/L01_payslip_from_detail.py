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


def prepare_path():
    print(os.path.abspath('..'))
    root = os.path.abspath('..') + '/testdata/payslip/'
    if not os.path.exists(root):
        os.mkdir(root)
    tmp_path = root + 'rs'
    if not os.path.exists(tmp_path):
        os.mkdir(tmp_path)


def get_path():
    root = os.path.abspath('..') + '/testdata/payslip/'
    print(root)
    return  root


def set_columns():
    return ['Global ID','Local ID', 'Learner Name', 'Item/ Program Name','Training Hours','Item Type','Start Date','End Date',
                           'Completion date','Expiration Date for Certifications','Vendor/ Instructor','Comments/ Remarks']

def get_str_date(l_date):
    if len(str(l_date)) > 10:
        return (str(l_date)[:10])
    else:
        return str(l_date)


def read_file():
    df_conf = pd.DataFrame()
    df_data = pd.DataFrame()
    typedict = {3: str,4:str}
    for fname in listdir(get_path()):
        # if fname.startswith('p'):
        if fname.startswith('pay'):
            try:
                f_name = get_path() + fname
                df_data = pd.read_excel(f_name, sheet_name='Sheet1', skip_blank_lines=True, parse_dates=False,convert_float=False)
                df_conf = pd.read_excel(f_name, sheet_name='Format_conf', dtype=typedict, header=0)
                # print(sheet_to_df_map.keys())
                df_data = df_data.append(df_data,sort=False)
                print(df_data.head())

            except Exception as e:
                print('read file failed:', fname)
                print('error log', e)
    df_data = df_data.sort_values('ID')

    df_conf.dropna(axis=0, how='all', thresh=None, subset=None, inplace=True)
    print(df_data.columns)
    return df_data.fillna(""), df_conf.fillna("")


def gen_pdf_from_xlsx(df_data,df_conf):
    df_temp = pd.DataFrame()

    list_globalid = list(set(df_data['ID'].to_list()))
    print(list_globalid)

    file_out = get_path() + 'failed_overview.xlsx'

    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = 0

    fields_not_found = []

    for gid in list_globalid:

        df_temp = df_data[df_data['ID'] == gid]
        o_detail = pd.Series(df_temp.iloc[0])
        try:
            filename = get_path() + 'rs/xlsx_p' + str(gid) + '.xlsx'
            # 创建一个excel
            workbook = xlsxwriter.Workbook(filename)
            worksheet = workbook.add_worksheet()

            title_format = workbook.add_format({'bold': True, 'text_wrap': False, 'align':'center','valign': 'top', 'font_size': 30,})
            title_format.set_bottom(1)
            block_format = workbook.add_format({'bold': True, 'text_wrap': False, 'align':'left', 'valign': 'top', 'font_size': 16})
            block_format.set_bottom(1)
            label_format = workbook.add_format({'bold': True, 'text_wrap': True, 'align':'left', 'valign': 'top', 'font_size': 12})
            label_format.set_bottom(3)
            value_format = workbook.add_format({'text_wrap': False, 'align':'right', 'valign': 'top', 'font_size': 11,'num_format': '#,##0.00'})
            value_format.set_bottom(3)

            last_group = 0
            for n_idx,n_row in df_conf.iterrows():
                try:
                    l_format = block_format
                    row = int(n_row['StartLine'])
                    column = int(n_row['StartColumn'])-1

                    if n_row['Group'] == 'Payslip':
                        worksheet.merge_range(row,column,row,column+5,data=n_row['Labels'],cell_format=title_format)
                    elif  n_row['Group'] == 'BasicInfo':
                        value = ""
                        if str(n_row['Fields']).strip() != "":
                            try:
                                value = o_detail[n_row['Fields']]
                                l_format = label_format
                            except Exception as e:
                                print('No field found', n_row['Fields'], 'error', e)
                                fields_not_found.append(n_row['Fields'])


                        worksheet.write(row,column,n_row['Labels'],l_format)

                        if n_row['Type'] != 'AMT':
                            worksheet.write(row,column+1,value,l_format)
                        else:
                            worksheet.write_number(row,column+1, value, value_format)

                    else:
                        if last_group != n_row['Group']:
                            base_row = int(n_row['StartLine'])
                        # else:

                        value = ""
                        if str(n_row['Fields']).strip() != "":
                            try:
                                value = o_detail[n_row['Fields']]
                                l_format = label_format
                            except Exception as e:
                                print('No field found', n_row['Fields'], 'error', e)
                                fields_not_found.append(n_row['Fields'])


                        if n_row['Type'] != 'AMT':
                            worksheet.write(base_row,column+1,value,l_format)
                            worksheet.write(base_row,column,n_row['Labels'],l_format)
                            base_row = base_row + 1
                        else:
                            if isinstance(value,float):
                                worksheet.write_number(base_row, column+1, value, value_format)
                                worksheet.write(base_row,column,n_row['Labels'],l_format)
                                base_row = base_row + 1

                except Exception as e:
                    print('write excel file failed with fields:', n_row['Group'],n_row['Labels'])
                    print('error log', e)

                last_group = n_row['Group']

            # print(x,y)

            worksheet.set_column("A:A",6)
            worksheet.set_column("B:B",18)
            worksheet.set_column("C:C",12)
            worksheet.set_column("D:D",5)
            worksheet.set_column("E:E",18)
            worksheet.set_column("F:F",12)
            worksheet.set_column("G:G",6)

            worksheet.set_paper(9)      #https://xlsxwriter.readthedocs.io/page_setup.html
            worksheet.set_margins(left=0.7,right=0.7, top=0.3, bottom=0.5)
            # (x,y) = df_temp.shape
            # worksheet.print_area(0,0,x-1,y-1)
            worksheet.fit_to_pages(1,1)     #
            # worksheet.set_landscape()       #worksheet.set_portrait()

            worksheet.set_default_row(hide_unused_rows=True)    #隐藏无效值
            workbook.close()
        except Exception as e:
            print('write excel file failed:', file_out)
            print('error log', e)

        try:
            file_pdf = get_path() + 'rs/pdf_p' + str(gid) + '.pdf'
            if os.access(file_pdf, os.F_OK):
                os.remove(file_pdf)
            books = xlApp.Workbooks.Open(filename, False)
            books.ExportAsFixedFormat(0, file_pdf)  #0 is pdf,1 is xps
            books.Close(False)
            print('Save PDF Files：', file_pdf)
        except Exception as e:
            print('Get PDF file failed:', file_pdf)
            print('error log', e)


    xlApp.Quit()
    print('Not found items:',set(fields_not_found))


if __name__ == '__main__':
    df_data = pd.DataFrame()
    df_conf = pd.DataFrame()
    time1 = time.time()

    prepare_path()
    # chdir(get_path())

    df_data,df_conf = read_file()
    gen_pdf_from_xlsx(df_data,df_conf)

    print("Total running time", time.time() - time1)
