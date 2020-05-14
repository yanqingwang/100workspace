# -*- coding: utf-8 -*-
"""
This is a script to read files and updated as pdf files.
use QT to generate pdf files in same folder, but need to set format
"""

from os import chdir, listdir
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
    df_data = pd.DataFrame()
    df_data_tmp = pd.DataFrame()
    filenames = listdir(get_path())
    for fname in filenames:
        if fname.startswith('p'):
        # if fname.startswith('Con'):
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


def generate_pdf(df_data):
    app = QtWidgets.QApplication(sys.argv)

    list_globalid = list(set(df_data['Global ID'].to_list()))
    print(list_globalid)

    for gid in list_globalid:
        html_file = 'P' + str(gid) + '.html'
        df_temp = df_data[df_data['Global ID'] == gid]
        df_temp.columns = set_columns()
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

        pd.set_option('colheader_justify', 'center')  # FOR TABLE <th>
        df_temp.to_html(html_file,index=False,justify='left',border="1")

        loader = QtWebEngineWidgets.QWebEngineView()
        loader.load(QtCore.QUrl(get_path() + html_file))

        layout = QPageLayout(
            QPageSize(QPageSize.A4),
            QPageLayout.Portrait, QMarginsF(8, 12, 8, 12)
        )

        def printFinished():
            page = loader.page()
            # print("%s Printing Finished!" % page.title())
            app.exit()

        def printToPDF(finished):
            loader.show()
            page = loader.page()
            page.printToPdf("%s.pdf" % page.title()[:9], layout)
            print("Generated pdf file %s.pdf" % page.title()[:9])

        loader.page().pdfPrintingFinished.connect(printFinished)
        loader.loadFinished.connect(printToPDF)

        app.exec_()
        loader.close()
    app.closeAllWindows()


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    file_out = 'Output_' + now_date + '_Learning_history_data.xlsx'
    time1 = time.time()

    chdir(get_path())

    df_data = read_file()
    output_data(df_data, file_out)
    generate_pdf(df_data)

    print("Total running time", time.time() - time1)
