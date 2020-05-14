# -*- coding: utf-8 -*-
"""
This is a script to read files and updated as pdf files.
"""

from datetime import date
import time
import datetime
import pandas as pd
import xlsxwriter
import os
import PyPDF4

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


def get_string(l_value):
    return str(l_value)


def read_file():
    df_data_tmp = pd.DataFrame()
    df_data = pd.DataFrame()
    file_path = get_path() + '2nd Submission/'
    filenames = os.listdir(file_path )
    for fname in filenames:
        if fname.startswith('Learning'):
        # if fname.startswith('test'):
            try:
                f_name = file_path + fname
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
                        print(fname,df_data_tmp.shape)
                        df_data_tmp['FileName'] = fname
                        df_data= df_data.append(df_data_tmp, sort=False)
                # print(df_data_tmp.head(1))

            except Exception as e:
                print('read file failed:', fname)
                print('error log', e)
    df_data['Global ID'] = df_data['Global ID'].apply(str)
    print(df_data.columns)
    print(df_data.shape)
    return df_data.fillna("")


def read_employee_file():
    df_ee = pd.DataFrame()
    df_ee_list = pd.DataFrame()
    try:

        f_name2 = get_path() + '\EmployeeHeadcount-Page1-20191231.xlsx'
        df_ee_list = pd.read_excel(f_name2, sheet_name='Excel Output', dtype = {'ZF Global ID':str}, parse_dates=False,skiprows=2)
        # print(df_ee_list.columns)
        df_ee_list['ZF Global ID'] = df_ee_list['ZF Global ID'].apply(str)

        df_ee = df_ee_list[['ZF Global ID','First Name','Last Name','Company (Legal Entity ID)','Company (Label)','Reporting Unit (Reporting Unit ID)','Employment Type (Label)','External Agency & Contingent Worker']]
        # print(df_ee.columns)

        return  df_ee.fillna("")

    except Exception as e:
        print('read status / employee file failed:')
        print('error log', e)


def output_data(df_data, file_out):
    df_ru = pd.DataFrame()

    # df_data['EETotal'] = df_data['EEDirect'] + df_data['EEIndirect']

    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(file_out, engine='xlsxwriter')
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


def gen_pdf_from_xlsx(df_data,df_employee):
    df_temp = pd.DataFrame()
    lf_ee = pd.DataFrame()

    df_data.drop(["FileName"],axis = 1,inplace=True)

    path = 'C:/temp/pdf/'
    list_columns = ['^UserId','documents','DocName','attachment','lastMod']
    df_list = pd.DataFrame(columns=list_columns)
    df_error_list = pd.DataFrame()

    list_globalid = list(set(df_data['Global ID'].to_list()))
    print(list_globalid)

    df_data['Start Date'] = df_data.apply(lambda x:get_str_date(x['Start Date']),axis=1)
    df_data['End Date'] = df_data.apply(lambda x:get_str_date(x['End Date']),axis=1)
    df_data['Completion date'] = df_data.apply(lambda x:get_str_date(x['Completion date']),axis=1)
    df_data['Expiration Date for Certifications'] = df_data.apply(lambda x:get_str_date(x['Expiration Date for Certifications']),axis=1)

    xlApp = DispatchEx("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = 0

    # df_data = df_data.sort_values('Global ID',inplace=True)
    df_data.sort_values(['Global ID','Start Date'],inplace=True,ascending=False)
    for gid in list_globalid:
        lf_ee = df_employee[df_employee['ZF Global ID'] == gid]
        if not lf_ee.empty:
            l_employee = pd.Series(lf_ee.iloc[0])
            df_temp = df_data[df_data['Global ID'] == gid]
            if not df_temp.empty:       # this will never be false, since the GID is from the table
                # try:
                #     df_temp.sort_values('Start Date',inplace=True,ascending=False)
                # except Exception as e:
                #     print('Sort table error for employee:', gid, e)
                #     emp_line = ['Global ID', gid, 'Sort']
                #     df_error_list = df_error_list.append (pd.Series(emp_line),ignore_index = True)

                try:
                    # file_name = 'LearningHistory_' + str(l_employee['Company (Legal Entity ID)']) + '_' + str(l_employee['Reporting Unit (Reporting Unit ID)'])  + '_' + str(gid)
                    file_name = 'LearningHistory_' + str(gid)
                    filename = path + 'xlsx/' + file_name + '.xlsx'
                    # 创建一个excel
                    df_writer = pd.ExcelWriter(filename, engine='xlsxwriter',date_format='yyyy-mm-dd')
                    workbook = df_writer.book
                    sheet_name = 'Sheet1'
                    df_temp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False,header=False,startrow=1)
                    worksheet = df_writer.sheets[sheet_name]
                    worksheet.set_default_row(hide_unused_rows=True)
                    # print(x,y)

                    # Add a header format.
                    h_format = workbook.add_format({'border': False,'font_size':3,'align': 'left', 'text_wrap': False,'fg_color':'#808080'})
                    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC',
                                                         'font_size': 10, 'border': 1})

                    common_format = workbook.add_format({'border': True,'font_size':10,'align': 'left', 'valign':'vcenter','text_wrap': True})

                    # Write the column headers with the defined format.
                    for col_num, value in enumerate(df_temp.columns.values):
                        worksheet.write(0, col_num, value)
                    worksheet.set_row(0,height=None,cell_format=header_format)
                    worksheet.set_column("C:C",15, cell_format=common_format)
                    worksheet.set_column("A:A",9, cell_format=common_format)
                    worksheet.set_column("B:B",10, cell_format=common_format)
                    worksheet.set_column("D:D",32, cell_format=common_format)
                    worksheet.set_column("E:E",6, cell_format=common_format)
                    worksheet.set_column("F:F",9, cell_format=common_format)
                    worksheet.set_column("G:H",11, cell_format=common_format)
                    worksheet.set_column("I:J",10, cell_format=common_format)
                    worksheet.set_column("K:K",16, cell_format=common_format)
                    worksheet.set_column("L:L",16, cell_format=common_format)

                    worksheet = df_writer.sheets[sheet_name]

                    # header = '&RInternal  '
                    # # header = '&RInternal  \n   Page:  &P     '
                    # worksheet.set_header(header)
                    header2 = '&C&20Learning History  \n\r &RInternal      &G      \r\n   '
                    # header2 = '&RInternal      &G          '
                    worksheet.set_header(header2, {'image_right': 'ZF_LOGO.png'})
                    worksheet.set_footer('&RPage  &P   of   &N    \t ')

                    worksheet.set_paper(9)      #https://xlsxwriter.readthedocs.io/page_setup.html
                    worksheet.set_margins(left=0.42,right=0.42, top=0.75, bottom=0.6)
                    # (x,y) = df_temp.shape
                    # worksheet.print_area(0,0,x-1,y-1)
                    worksheet.fit_to_pages(1,0)     #width, hight
                    worksheet.set_landscape()       #worksheet.set_portrait()
                    workbook.close()
                    df_writer.save()
                    df_writer.close()
                except Exception as e:
                    print('write excel file failed:', filename)
                    print('error log', e)
                    emp_line = ['Global ID', gid, 'Write Excel']
                    df_error_list = df_error_list.append (pd.Series(emp_line),ignore_index = True)

                try:
                    file_pdf_name =  file_name + '.pdf'
                    file_pdf = path + 'pdfs/' + file_pdf_name
                    if os.access(file_pdf, os.F_OK):
                        os.remove(file_pdf)
                    books = xlApp.Workbooks.Open(filename, False)
                    books.ExportAsFixedFormat(0, file_pdf)  #0 is pdf,1 is xps
                    books.Close(False)
                    print('Save PDF Files：', file_pdf)

                    # emp_line = [gid,'documents',l_employee['First Name'] + '_' + l_employee['Last Name']+'_'+str(gid),file_pdf_name,'1/1/2019']
                    emp_line = [gid,'documents','LearningHistory_'+l_employee['First Name'] + '_' + l_employee['Last Name']+'_'+str(gid),file_pdf_name,'1/1/2020']
                    df_list = df_list.append(pd.Series(emp_line,index=list_columns),ignore_index = True)

                except Exception as e:
                    print('Get PDF file failed:', file_pdf_name)
                    print('error log', e)
                    emp_line = ['Global ID', gid, 'Convert to PDF']
                    df_error_list = df_error_list.append (pd.Series(emp_line),ignore_index = True)

                # # just test, encrypt the pdf files
                # try:
                #     input_pdf = PyPDF4.PdfFileReader(file_pdf)
                #     pdf_writer = PyPDF4.PdfFileWriter()
                #     output = PyPDF4.PdfFileWriter()
                #     file_pdf2 = path + 'pdfe/' + file_pdf_name
                #     pdf_writer.appendPagesFromReader(input_pdf)
                #     pdf_writer.addMetadata({'/Title': 'Learning History','/Author':'ZF China','/Subject': 'ZF Learning History'})
                #     pdf_writer.encrypt('test',owner_pwd=None)
                #     # pdf_writer.encrypt('',owner_pwd='Ross')
                #     pdf_writer.write(open(file_pdf2, 'wb'))
                #     print("write encrypt pdf",file_pdf2)
                # except Exception as e:
                #     print('Encrypt PDF file failed:', file_pdf_name)
                #     print('error log', e)
                #     emp_line = ['Global ID', gid, 'Encrypt PDF']
                #     df_error_list = df_error_list.append (pd.Series(emp_line),ignore_index = True)

        # if lf_ee.empty:
        else:
            emp_line = ['Global ID', gid, 'Not Found']
            df_error_list = df_error_list.append (pd.Series(emp_line),ignore_index = True)

    xlApp.Quit()
    file_list = path + 'Output_' +  '_backgroundtemplate_zffriedric.xlsx'
    output_data(df_list, file_list)
    file_csv = path + 'Output_' +  '_backgroundtemplate_zffriedric.csv'
    df_list.to_csv(file_csv, sep = ',',encoding='UTF-8',index=False)

    file_csv = path + 'Output_' +  'Error_list.csv'
    df_error_list.to_csv(file_csv, sep = ',',encoding='UTF-8',index=False)



if __name__ == '__main__':
    df_data = pd.DataFrame()
    df_employee = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")
    file_out_with_path = get_path()  + '2nd Submission/' + 'Output_' + now_date + '_Learning_history_data.xlsx'

    time1 = time.time()

    os.chdir(get_path())

    df_data = read_file()
    output_data(df_data, file_out_with_path)

    df_employee = read_employee_file()
    gen_pdf_from_xlsx(df_data, df_employee)

    print("Total running time", time.time() - time1)
