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
import numpy as np
import re

class handling_data(object):

    def __init__(self):
        self.country = {"AU":"AUS","CN":"CHN",'ID':'IDN','JP':'JPN','KR':"KOR",'MY':'MYS',
                        'PH':'PHL','SG':'SGP','TW':'TWN','TH':'THA','AE':'ARE','VN':'VNM'}
        self.root = 'c:/temp/orgchart/'
        self.data = 'AP_EmployeeBasicInfov2light-20200131.xlsx'
        self.bcs_data = 'DataQualityReport_01_2020_original.xlsm'


    def get_div(self,division_short):
        if division_short in ['A','B','C','E','I','P','R''T','U']:
            return division_short
        else:    #['-','D','F','G','H','M','MK','O','Q','S','V']:
            return 'Z'


    def get_direct(self,external, employee_type, employee_class,l_global_id):
        try:
            if external != 'Yes' and employee_class == 'Direct' and employee_type not in ["Intern/Students","Apprentice","Vacation Workers DE"]:
                return 1
            else:
                return 0
        except Exception as e:
            print('get direct employee error:', l_global_id)
            print('error log', e)
            return 0


    def get_indirect(self,external, employee_type, employee_class,l_global_id):
        try:
            if external != 'Yes' and employee_class == 'In-direct/Salaried' and employee_type not in ["Intern/Students","Apprentice","Vacation Workers DE",""]:
                return 1
            else:
                return 0
        except Exception as e:
            print('get direct employee error:', l_global_id)
            print('error log', e)
            return 0


    def check_ZID(self,indirect,ZID):
        try:
            if str(ZID).upper().startswith('NA_') or  str(ZID).upper().startswith('MIG') or str(ZID)[:1].isnumeric(): #invalid ZID
                return 1
            else:
                return 0
        except Exception as e:
            print('Check email failed:')
            print('error log', e)
            return 0


    def check_indirect(self,indirect, res):
        try:
            if int(indirect) > 0 and int(res) > 0:
                return 1
            else:
                return 0
        except Exception as e:
            print('Check indirect related data ')
            print('error log', e)
            return 0


    def check_date(self,hire_date, grp_date):
        try:

            l_hiredate = pd.to_datetime(hire_date)
            l_grpdate = pd.to_datetime(grp_date)
            if l_hiredate < l_grpdate :    #hire date earlier than group date,  return error
                return 1
            else:
                return 0
        except Exception as e:
            print('Comparing hire date vs group date', 'Hire date:',hire_date, 'Group Date',grp_date)
            print('error log', e)
            return 0



    def check_email_zid(self,email, zid):
        try:
            if email > 0 and zid > 0:    #hire date earlier than group date,  return error
                return 1
            else:
                return 0
        except Exception as e:
            print('With email, but without valid ZID', e)
            return 0


    def check_cnt(self,country_id):
        if country_id == 'CHN':
            return 'CHN & HK'
        else:
            return 'Others'


    def check_email(self,email):
        if len(str(email)) > 0:
            lv_mail = str(email).lower().strip()
        else:
            return 0
        try:
            if re.match(r'^([a-zA-Z0-9.]+)@zf.com$',lv_mail) or re.match(r'^([a-zA-Z0-9.]+)@trw.com$',lv_mail):
                return 1
            else:
                return 0
        except Exception as e:
            print('Check email failed:')
            print('error log', e)
            return 0


    def get_gender(self,gender, female):
        if female == 'X' and gender == 'F':
            return 1
        elif female != 'X' and gender == 'M':
            return 1


    def get_status(self,status):
        if status in ['Active','Unpaid Leave']:
            return 'A'
        else: #Retired, temrinate,Dormant Discarded
            return 'I'


    def build_hierarchy(self, df_raw_data, list_dept, level):
        df = pd.DataFrame()
        new_list_dept = []
        new_df_raw_data = pd.DataFrame()
        for n_idx, n_row in df_raw_data.iterrows():
            try:
                line_data = n_row
                if line_data['Parent Department (ID)'] in list_dept:
                    new_list_dept.append(line_data['Department ID'])

                    line_data['Department ID-' + level] =  line_data['Department ID']
                    line_data['Department Long-' + level] =  line_data['Department Long']
                    line_data['Department Short-' + level] =  line_data['Department Short']

                new_df_raw_data = new_df_raw_data.append(pd.Series(line_data), ignore_index=True)

            except Exception as e:
                print('Exception:', line_data['Department ID'],e)
        return new_df_raw_data.fillna(''), list(set(list_dept))


    def get_data(self):
        df = pd.DataFrame()

        emp_file = self.root + self.data

        try:
            df = pd.DataFrame(pd.read_excel(io=emp_file, sheet_name='Excel Output', header=0, skiprows=2))
            # print(df.columns)
            print(df.head(2))

        except Exception as e:
            print('Exception:', emp_file,e)
        # print(df.columns)
        return df.fillna('')


    def final_summary(self):
        list_parent = []
        df_proc = pd.DataFrame()

        df_data = self.get_data()
        print('employee data', df_data.columns)
        file_out = 'Output_' + now_date + '_orgchart.xlsx'

        df_proc,list_parent = self.build_hierarchy(df_data, list_parent, 0)


        df_simple = pd.DataFrame()

        df_res = pd.DataFrame()

        df_data['EETotal'] = df_data['EEDirect'] + df_data['EEIndirect']

        df_ap_division = df_data.groupby(['Division_rs'])[
            'Total_EE','EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker', 'email_check','ZID_check','SM_Chck','IM_Chck',
            'BP_Chck','DATE_CHK','EMAIL_ZID_CHK'].sum().reset_index()
        df_division = df_ap_division.sort_values("EETotal", ascending=False)

        # All
        df_simple = df_data.groupby(['Country (ID)', 'Company (Label)', 'RU'])[
            'Total_EE','EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker', 'email_check','ZID_check','SM_Chck','IM_Chck',
            'BP_Chck','DATE_CHK','EMAIL_ZID_CHK','ID_email_check','ID_SM_Chck','ID_IM_Chck','ID_BP_Chck'].sum().reset_index()

        df_simple['Country'] = df_simple.apply(lambda x: self.check_cnt(x['Country (ID)']), axis=1)

        df_simple = df_simple.sort_values("EETotal", ascending=False)
        print('df_simple:', 'columns', df_simple.shape[0], '\n', df_simple.head(2))


        # combine with BCS data

        for n_idx, n_row in df_simple.iterrows():
            try:
                line_rs = n_row
                bcs_line = pd.Series(df_bcs_data[df_bcs_data['RU'] == str(line_rs['RU'])].iloc[0])
                line_rs['DIV'] = bcs_line['DIV']
                line_rs['BCS_EMPs'] = bcs_line['BCS_EMPs']
                line_rs['SF_EMPs'] = bcs_line['SF_EMPs']
                line_rs['BCS_EMPs_DT'] = bcs_line['BCS_EMPs_DT']
                line_rs['SF_EMPs_DT'] = bcs_line['SF_EMPs_DT']
                line_rs['BCS_EMPs_ID'] = bcs_line['BCS_EMPs_ID']
                line_rs['SF_EMPs_ID'] = bcs_line['SF_EMPs_ID']
                line_rs['BCS_Others'] = bcs_line['BCS_Others']
                line_rs['SF_Others'] = bcs_line['SF_Others']

                line_rs['GAPS_EMP'] = bcs_line['GAPS_EMP']
                line_rs['ABS_GAPS_EMP'] = bcs_line['ABS_GAPS_EMP']
                line_rs['GAPS_OTHERS'] = bcs_line['GAPS_OTHERS']
                line_rs['ABS_GAPS_OTHERS'] = bcs_line['ABS_GAPS_OTHERS']

                df_res = df_res.append(pd.Series(line_rs), ignore_index=True)
            except Exception as e:
                print('Exception:', n_row['Country (ID)'],n_row['RU'],e)

        df_res2 = df_res.groupby(['Country', 'Country (ID)', 'DIV' , 'RU' ])[
            'Total_EE','EETotal','EEIndirect','email_check','ZID_check','SM_Chck','IM_Chck','BP_Chck','ID_email_check','ID_SM_Chck','ID_IM_Chck','ID_BP_Chck',
             'DATE_CHK','EMAIL_ZID_CHK','BCS_EMPs','SF_EMPs','BCS_EMPs_DT','SF_EMPs_DT','BCS_EMPs_ID','SF_EMPs_ID','BCS_Others','SF_Others',
            'GAPS_EMP','GAPS_OTHERS','ABS_GAPS_EMP','ABS_GAPS_OTHERS'].sum().reset_index()

        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(self.root+file_out,engine='xlsxwriter')
            workbook = df_writer.book
            # chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'

            sheet_name = '10_Initial'
            df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '10_bcs_initial'
            df_bcs_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)


            sheet_name = 'AP_20_division'
            df_division.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '30_simple'
            df_simple.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            # pvt_tmp = pd.pivot_table(df_data, index=['Country (ID)', 'Company (Label)'], values=['EETotal'],
            #                          columns=['JobFamily'],aggfunc={np.sum}, fill_value=0)
            #                          # columns=['JobFamily'],aggfunc={np.sum}, fill_value=0,margins=True)
            # pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")
            sheet_name = '20_res'
            df_res.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '20_res_simple'
            df_res2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            workbook.close()
        except Exception as e:
            print('write file failed:', file_out)
            print('error log', e)


if __name__ == '__main__':
    df_data = pd.DataFrame()
    now_date = date.today().strftime("%Y%m%d")

    time1 = time.time()
    obj_factor = handling_data()
    obj_factor.final_summary()

    # head_count_summary(df_data, file_out

    print("Total running time", time.time() - time1)
