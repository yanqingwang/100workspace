# -*- coding: utf-8 -*-
""" new file, attention to fields as below:
national Id, Date of Birth,Original Start Date-->Start Period of Employment,Cost Center (Label)-->Cost Centre (Label)
Created on Apr 25 11:05:58 2019
@author: Z659190
"""

import re
import pandas as pd
import time

# read file
def read_data(file_name):
    try:
        df = pd.DataFrame(pd.read_excel(io=file_name, sheet_name='Sheet1', header=0, skiprows=0))
        # df = pd.DataFrame(pd.read_excel(io=file_name, sheet_name='Sheet1', header=0, skiprows=0,index_col='Person Id'))
        return df
    except Exception as e:
        print('read file failed:', file_name)
        print('error log', e)


def write_data(df_initial,df_detail,df_summary,p_output_root):

    out_file = p_output_root + 't3_output_'+time.strftime("%Y-%m-%d",time.localtime())+'res.xlsx'

    try:
        # 创建一个excel
        df_writer = pd.ExcelWriter(out_file,engine='xlsxwriter')
        workbook = df_writer.book

        sheet_name = 'initial'
        df_initial.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'detail'
        df_detail.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name = 'Summary'
        df_summary.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        workbook.close()
        print("Successfully write file",)
    except Exception as e:
        print('write file failed:', out_file)
        print('error log', e)


def compare_detail(root_dir, line_columns,filename_old,filename_new):
    df_old = pd.DataFrame()
    df_new = pd.DataFrame()
    df_handling = pd.DataFrame()
    df_res = pd.DataFrame(columns=(line_columns))

    # filename_old = "file1.xlsx"
    # filename_new = "filenew.xlsx"
    file1 = root_dir + filename_old
    file2 = root_dir + filename_new

    df_old = read_data(file1)
    df_new = read_data(file2)
    df_old = df_old.fillna(0)
    df_new = df_new.fillna(0)

    # df_handling = df_new.head(0)

    personlist_old = df_old[u'Person Id'].tolist()
    personlist_new = df_new[u'ZF Global ID'].tolist()
    # print(df_old['Person Id'].tolist())
    # df_old.set_index(['Person Id','Reporting Unit ID'])

    print('data comparing',time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time())))
    # Handling new hires and changes, for changes, we need to compare the special fields one by one
    for n_idx,n_row in df_new.iterrows():
        changed_username = 0
        change_jobinformation = 0
        change_email = 0
        change_employeedata = 0
        change_jobrelationship = 0
        change_sm = 0
        change_icm = 0
        change_hrbp = 0
        change_localid = 0
        change_external = 0
        change_costcenter = 0
        change_department = 0
        change_startdate = 0
        change_ee = 0
        change_ru = 0

        if n_idx % 200 == 0:
            print('new file', n_idx, n_row['ZF Global ID'])


        if not n_row['ZF Global ID'] in personlist_old: # New hire
            n_row['Status'] = 'New Hire'
            df_handling = df_handling.append(pd.Series(n_row),sort=False)

            line_rs = {'RU':n_row['Reporting Unit (Reporting Unit ID)'],'Person Id':n_row['ZF Global ID'],'New':1}
            # Global variant
            line_rs['Headcount'] = 1
            if n_row['External Agency Worker'] == 'Yes':
                line_rs['External_Agency'] = 1
            line_rs['First Name'] = n_row['First Name']
            line_rs['Last Name'] = n_row['Last Name']
            df_res = df_res.append(pd.Series(line_rs),ignore_index=True)

        elif n_row['ZF Global ID'] in personlist_old:   # change
            n_row['Status'] = 'Change'
            df_handling = df_handling.append(pd.Series(n_row),sort=False)
            line_rs = {'RU':n_row['Reporting Unit (Reporting Unit ID)'],'Person Id':n_row['ZF Global ID']}
            # Global variant
            line_rs['Headcount'] = 1
            if n_row['External Agency Worker'] == 'Yes':
                line_rs['External_Agency'] = 1

            # print(n_row['Person Id'],'change comparing')
            df_o=df_old[df_old['Person Id']==n_row['ZF Global ID']]       #拿到符合条件的记录,dataFrame
            # print(type(df_o))
            # print(df_o)
            o_row = pd.Series(df_o.iloc[0])                   #获取行数据
            # o_row = pd.Series(df_o.iloc[0]).to_dict()                   #获取行数据

            # print(o_row)
            # print(type(o_row))
            # print(o_row['Reporting Unit ID'])


            if n_row['ZID'] != o_row['Username']:
                changed_username = changed_username + 1

            re_nlocaid = re.sub(r"\b0*([0-z][0-z]*|0)", r"\1", str(n_row['Local ID']))
            re_olocaid = re.sub(r"\b0*([0-z][0-z]*|0)", r"\1", str(o_row['China  Global Information Local ID']))
            if re_nlocaid != re_olocaid:
                change_localid = change_localid + 1

            if str(n_row['National Id']) != str(o_row['National ID']):
                change_employeedata = change_employeedata + 1
            if str(n_row['First Name']).upper() != str(o_row['First Name']).upper():
                change_employeedata = change_employeedata + 1
            if str(n_row['Last Name']).upper() != str(o_row['Last Name']).upper():
                change_employeedata = change_employeedata + 1
            if pd.to_datetime(n_row['Date Of Birth']) != pd.to_datetime(o_row['Date Of Birth']):
                change_employeedata = change_employeedata + 1
            if n_row['Alternate First Name'] != o_row['Alternate First Name']:
                change_employeedata = change_employeedata + 1
            if n_row['Alternate Last Name'] != o_row['Alternate Last Name']:
                change_employeedata = change_employeedata + 1
            # first line M2, 2nd line, M4
            # if n_row['Home Address'] != o_row['Home Address1 | Address Line 1 | Name of Addressee | Care Of | Street | Street and House Number | Street and House No. | Detailed Address | Building Number and Street | Extra Address Line | House Number and Street | Kanji Address Line 1 | City/District/County | Contact Name | Addressee | Care of | First Address Line | Village | Street Name and Number'] \
            #         and n_row['Home Address'] != o_row['Home Address1 | Address Line 1 | Name of Addressee | Care Of | Street | Street and House Number | Street and House No. | Detailed Address | House Number | Building Number and Street | Extra Address Line | House Number and Street | Kanji Address Line 1 | City/District/County | Contact Name | Addressee | Care of | First Address Line | Village | Street Name and Number']:
            #     change_employeedata = change_employeedata + 1

            if str(n_row['Business email address']).upper() != str(o_row['Business  Email Information Email Address']).upper():
                change_email = change_email + 1


            if n_row['Position Code'] != o_row['Position']:
                change_jobinformation = change_jobinformation + 1
            if n_row['Position Title'] != o_row['Position Title'] and n_row['Position Code'] == o_row['Position']:
                change_jobinformation = change_jobinformation + 1
            # job_title = (n_row['Job Classification (Job Code)'] + '-' + n_row['Job Classification (Label)']).upper()
            # if job_title != o_row['Job Classification'].upper():
            #     change_jobinformation = change_jobinformation + 1

            if float(n_row['Standard Weekly Hours']) != float(o_row['Standard Weekly Hours']):
                change_jobinformation = change_jobinformation + 1
            if (n_row['Department Short Text'] != o_row['Department Short Text'] and o_row['Department Short Text']) != 0:
                change_jobinformation = change_jobinformation + 1
                change_department = change_department + 1
            if n_row['Local Employment Type (Label)'] != o_row['Local Employment Type']:
                change_jobinformation = change_jobinformation + 1
            if n_row['Local Employee Class (Label)'] != o_row['Local Employee Class']:
                change_jobinformation = change_jobinformation + 1
            if int(n_row['Reporting Unit (Reporting Unit ID)']) != int(o_row['Reporting Unit ID']):
                change_jobinformation = change_jobinformation + 1
                change_ru = change_ru  + 1
            if pd.to_datetime(n_row['Contract End Date']) != pd.to_datetime(o_row['Contract End Date']):
                change_jobinformation = change_jobinformation + 1
            if pd.to_datetime(n_row['Start Period of Employment']) != pd.to_datetime(o_row['Employment Details Original Start Date']) and (not pd.isnull(o_row['Employment Details Original Start Date']) or not pd.isnull(n_row['Original Start Date'])):
                change_jobinformation = change_jobinformation + 1
                change_startdate = change_startdate + 1
            try:
                list_location = o_row['Location'].split('-')
                if str(n_row['Location (Name)']).upper() != str(list_location[1]).upper():
                    change_jobinformation = change_jobinformation + 1
            except Exception as e:
                print('split location error','Global ID',n_row['ZF Global ID'],'Location',o_row['Location'])
            re_ncc = re.sub(r"\b0*([0-z][0-z]*|0)", r"\1", str(n_row['Cost Centre (Label)']))
            re_occ = re.sub(r"\b0*([0-z][0-z]*|0)", r"\1", str(o_row['Cost Center Cost Center']))
            if re_ncc != re_occ:        #需要去掉前置0, cost center
                change_jobinformation = change_jobinformation + 1
                change_costcenter = change_costcenter + 1

            if int(n_row['Solid Line Manager']) != int(o_row['Solid Line Manager User Sys ID']):
                change_sm = change_sm + 1
            if int(n_row['In-country Manager ID']) != int(o_row['In-Country Manager  Job Relationships User Id']):
                change_icm = change_icm + 1
            if int(n_row['BP ID']) != int(o_row['HR Business Partner  Job Relationships User Id']):
                change_hrbp = change_hrbp + 1
            if n_row['External Agency Worker'] != o_row['Contingent Worker']:
                change_external = change_external + 1

            change_jobrelationship = change_sm + change_icm + change_hrbp

            line_rs['Change_Username'] = changed_username
            line_rs['Change_Solid_Manager'] = change_sm
            line_rs['Change_In-Country_Manager'] = change_icm
            line_rs['Change_HRBP'] = change_hrbp
            line_rs['Change_LocalID'] = change_localid
            # line_rs['Change_EmployeeData'] = change_employeedata
            if  change_employeedata > 0:
                line_rs['Chg_Emp_No'] = 1
            line_rs['Change_Email'] = change_email
            # line_rs['Change_JobInformation'] = change_jobinformation
            if  change_jobinformation > 0:
                line_rs['Chg_JobInfo_No'] = 1
            line_rs['Change_JobRelationship'] = change_jobrelationship
            if change_jobrelationship > 0:
                line_rs['Chg_Mgr_No'] = 1
            line_rs['Change_ExternalAgency'] = change_external
            line_rs['Change_CostCenter'] = change_costcenter
            line_rs['Change_Department'] = change_department
            line_rs['Change_StartDate'] = change_startdate
            line_rs['Change_RU'] = change_ru

            change_ee = changed_username + change_jobinformation + change_email + change_employeedata + change_jobrelationship + \
                        change_localid + change_external + change_costcenter + change_department + change_startdate
            if change_ee > 0:
                change_ee = 1
            line_rs['Change'] = change_ee
            if change_ee < 1:
                line_rs['Unchange'] = 1

            line_rs['First Name'] = n_row['First Name']
            line_rs['Last Name'] = n_row['Last Name']

            df_res = df_res.append(pd.Series(line_rs),ignore_index=True)

    for o_idx, o_row in df_old.iterrows():
        if o_idx % 500 == 0:
            print('old file', o_idx, o_row['Person Id'])
        if not o_row['Person Id'] in personlist_new:  # Termination
            n_row['Status'] = 'Termination'
            n_row['Reporting Unit (Reporting Unit ID)'] = o_row['Reporting Unit ID']
            n_row['ZF Global ID'] = o_row['Person Id']
            df_handling = df_handling.append(pd.Series(n_row),sort=False)
            line_rs = {'RU':o_row['Reporting Unit ID'],'Person Id':o_row['Person Id'],'Headcount':0,'Termination':1}
            line_rs['First Name'] = o_row['First Name']
            line_rs['Last Name'] = o_row['Last Name']
            df_res = df_res.append(pd.Series(line_rs),ignore_index=True)

    return df_handling,df_res


def summary_data(p_df_detail):
    print('Summary data',time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time())))
    p_df_detail.drop('Person Id',axis=1, inplace=True)
    df_sum = p_df_detail.groupby('RU')['Headcount','External_Agency', 'New', 'Termination', 'Change', 'Unchange','Change_Username','Change_LocalID','Change_CostCenter','Chg_JobInfo_No',
    'Change_Email', 'Chg_Emp_No', 'Chg_Mgr_No','Change_Department','Change_StartDate','Change_ExternalAgency','Change_RU','Change_JobRelationship','Change_Solid_Manager','Change_In-Country_Manager', 'Change_HRBP'].sum().reset_index()
    print(df_sum.head())
    return df_sum


if __name__ == '__main__':

    # 获取目标文件夹的路径
    root_dir = "C:/temp/compare/"
    output_root = 'C:/temp/compare/'
    time1 = time.time()

    file_old = '03_t3_data_validation_20190429.xlsx'
    file_new = "03_ChinaEmployeeBasicInfov2-Page1-20190924.xlsx"

    # line_columns_detail = (
    # 'RU', 'Person Id', 'Headcount', 'New', 'Termination', 'Change', 'Unchange','Change_Username','Change_LocalID','Change_JobInformation','Change_CostCenter','Chg_JobInfo_No',
    # 'Change_Email', 'Change_EmployeeData','Chg_Emp_No', 'Change_JobRelationship','Chg_Mgr_No','Change_Department','Change_StartDate','Change_ExternalAgency','External_Agency' )
    line_columns_detail = (
    'RU', 'Person Id', 'Headcount','External_Agency', 'New', 'Termination', 'Change', 'Unchange','Change_Username','Change_LocalID','Change_CostCenter','Chg_JobInfo_No',
    'Change_Email', 'Chg_Emp_No', 'Chg_Mgr_No','Change_Department','Change_StartDate','Change_ExternalAgency','Change_RU','Change_JobRelationship','Change_Solid_Manager','Change_In-Country_Manager', 'Change_HRBP')

    df_initial, df_detail = compare_detail(root_dir,line_columns_detail,file_old,file_new)

    df_sum = summary_data(df_detail)

    write_data(df_initial,df_detail,df_sum,output_root)

    time2 = time.time()

    print("Total running time", time2 - time1 )