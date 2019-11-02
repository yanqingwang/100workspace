# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""


from os import chdir,listdir
from datetime import date
import time
import datetime
import pandas as pd
import xlsxwriter
import sys

from PyQt5.QtWidgets import QApplication
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets
from PyQt5.QtCore import QMarginsF
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QPageLayout, QPageSize

from IPython.display import HTML

from win32com import client

def get_path():
    return 'c:/Users/z659190/Documents/10 Work/10 MyHRSuit/10 Project/131 learning/30_Workpackages/37 Data migration/1st Submission/'


def read_file():
    df_data = pd.DataFrame()
    filenames = listdir(get_path())
    for fname in filenames:
        if fname.startswith('Con'):
            try:
                f_name = get_path() + fname
                sheet_to_df_map = pd.read_excel(f_name, sheet_name=None,skip_blank_lines=True,parse_dates=False)
                # sheets = pd.ExcelFile(fname)
                # # this will read the first sheet into df
                # for l_sheet in sheets:
                #     df_data = pd.DataFrame(pd.read_excel(io=fname, sheet_name=l_sheet, header=0, skiprows=0))
                # print(sheet_to_df_map.keys())
                for key in sheet_to_df_map.keys():
                    if not key in ['Notes','Format- various systems', 'Questions']:
                        df_data = df_data.append(sheet_to_df_map[key],sort=False)
                print(df_data.head())

            except Exception as e:
                print('read file failed:', fname)
                print('error log', e)
    df_data = df_data.sort_values('Global ID\n*')
    print(df_data.columns)
    return  df_data.fillna("")

def output_data(df_data,file_out):
    df_ru = pd.DataFrame()

    # df_data['EETotal'] = df_data['EEDirect'] + df_data['EEIndirect']

    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(get_path()+file_out,engine='xlsxwriter')
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


def DateFormat(x):
    try:
        return str(time.strftime('%YYYY-%mm-%dd', x))
    except Exception as e:
        print("Exception", e)


def generate_pdf(df_data):
    app = QtWidgets.QApplication(sys.argv)

    list_globalid = list(set(df_data['Global ID\n*'].to_list()))
    print(list_globalid)

    for gid in list_globalid:
        html_file = 'P'+str(gid)+'.html'
        df_temp = df_data[df_data['Global ID\n*'] == gid]
        th_props = [
            ('font-size', '8px'),
            ('text-align', 'center'),
            ('font-weight', 'bold'),
            ('color', '#6d6d6d'),
            ('background-color', '#f7f7f9')
        ]
        td_props = [
            ('font-size', '8px')
        ]

        # Set table styles
        styles = [
            dict(selector="th", props=th_props),
            dict(selector="td", props=td_props)
        ]
        # df_temp.style.set_table_styles(styles)

        pd.set_option('colheader_justify', 'center')  # FOR TABLE <th>
        html = (df_temp.style.set_table_styles(styles)
                .set_caption("Hover to highlight."))
        # df_temp.to_html(html_file,index=False,justify='left',border="1",formatters={'Start Date\n(yyyy-mm-dd)\n':DateFormat})
        df_temp.to_html(html_file,index=False,formatters={'Start Date\n(yyyy-mm-dd)\n':DateFormat})
        # with open(html_file, 'w') as f:
        #     f.write(table=df_temp.to_html(classes="styles"))

        loader = QtWebEngineWidgets.QWebEngineView()
        loader.load(QtCore.QUrl(get_path()+html_file))

        layout = QPageLayout(
            QPageSize(QPageSize.A4),
            QPageLayout.Portrait, QMarginsF(8, 12, 8, 12)
        )

        def printFinished():
            page = loader.page()
            print("%s Printing Finished!" % page.title())
            app.exit()

        def printToPDF(finished):
            loader.show()
            page = loader.page()
            page.printToPdf("%s.pdf" % page.title()[:9], layout)

        loader.page().pdfPrintingFinished.connect(printFinished)
        loader.loadFinished.connect(printToPDF)

        app.exec_()
        loader.close()

    app.closeAllWindows()
    

def gen_pdf_from_xlsx(df_data):
    df_temp = pd.DataFrame()
    
    list_globalid = list(set(df_data['Global ID\n*'].to_list()))
    print(list_globalid)

    for gid in list_globalid:
           
        df_temp = df_data[df_data['Global ID\n*'] == gid]
        try:
            filename = get_path() + 'p' + str(gid) + '.xlsx'
            # 创建一个excel
            df_writer = pd.ExcelWriter(filename,engine='xlsxwriter')
            workbook = df_writer.book
            # chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'
            # df_data = df_data.sort_values("Global ID", ascending=False)
            # df_data = df_data.sort_values('Global ID',inplace=True)
            sheet_name = 'Sheet1'
            df_temp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
    
            workbook.close()
            
            xlApp = client.Dispatch("Excel.Application")
            books = xlApp.Workbooks.Open(filename)
            file_pdf = get_path() + 'p' + str(gid) + '.pdf'
            books.ExportAsFixedFormat(0, file_pdf)
            xlApp.Quit()
            
        except Exception as e:
            print('write file failed:', file_out)
            print('error log', e)



if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    file_out = 'Output_'+ now_date+ '_Learning_history_data.xlsx'
    time1 = time.time()

    chdir(get_path())

    df_data = read_file()
    output_data(df_data, file_out)
#    generate_pdf(df_data)
    gen_pdf_from_xlsx(df_data)

    print("Total running time", time.time() - time1)
