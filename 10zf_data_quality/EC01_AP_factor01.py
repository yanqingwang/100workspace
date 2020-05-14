# -*- coding: utf-8 -*-
"""
age range / hire / termination /
@author: Z659190
Headcount
Emails & ZID
Managers & HR BP
Entry Date
Under 18 Employees

"""

from datetime import date
import time
import pandas as pd
import xlsxwriter
import numpy as np
import re
from collections import defaultdict


class ApFactorData(object):
    def __init__(self):
        self.country = {"AU":"AUS","CN":"CHN",'ID':'IDN','JP':'JPN','KR':"KOR",'MY':'MYS',
                        'PH':'PHL','SG':'SGP','TW':'TWN','TH':'THA','AE':'ARE','VN':'VNM'}
        self.division = ['A','B','C','E','I','P','R','T','U','Z']
        self.root = 'c:/temp/Quality/'
        self.empfile = 'AP_EmployeeBasicInfov2light-20200430.xlsx'
        self.bcs_data = 'DataQualityReport_2020_03.xlsm'
        self.now_date = date.today().strftime("%Y%m%d")
        self.result = defaultdict(dict)

    def get_div(self,division_short):
        if division_short in ['A','B','C','E','I','P','R','T','U']:
            return division_short
        else:    #['-','D','F','G','H','M','MK','O','Q','S','V']:
            return 'Z'

    def get_div_from_bcs(self,emp_ru,emp_div):
        new_div = ""
        new_div = self.result[str(emp_ru)]
        if new_div == {}:
            new_div = self.get_div(emp_div)
        return new_div
        # try:
        #     new_div = self.result.get([str(emp_ru)])
        # except Exception as e:
        #     print('get divsion from bcs failed: RU', emp_ru, "use old div:", emp_div)
        #     print('error log', e)
        #     new_div = self.get_div(emp_div)
        # return new_div


    def get_age_range(self,l_date):
        try:
            l_dateDay = pd.to_datetime(l_date)
            today = date.today()
            age = today.year - l_dateDay.year - ((today.month, today.day) < (l_dateDay.month, l_dateDay.day))
            if pd.isna(l_dateDay):
                return 'NoBirthDate'
            elif age < 18:
                return '0--18'
            elif age < 25:
                return '18--25'
            elif age < 30:
                return '25--30'
            elif age < 35:
                return '30--35'
            elif age < 40:
                return '35--40'
            elif age < 50:
                return '40--50'
            elif age < 60:
                return '50--60'
            elif age < 70:
                return '60--70'
            elif age < 80:
                return '70--80'
            else:
                return '80+'

        except Exception as e:
            print('Convert date failed:', l_date)
            print('error log', e)
            return 'Unkown'


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
            print('get indirect employee error:', l_global_id)
            print('error log', e)
            return 0


    def check_ZID(self,indirect,ZID):
        try:
            if str(ZID).upper().startswith('NA_') or  str(ZID).upper().startswith('NA-') or str(ZID).upper().startswith('Z_') or \
               str(ZID).upper().startswith('MIG') or str(ZID)[:1].isnumeric(): #invalid ZID
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
            if pd.isna(grp_date) and pd.isna(hire_date):
                return 0
            elif  pd.isna(grp_date) and not pd.isna(hire_date):
                return 1
            else:
                l_hiredate = pd.to_datetime(hire_date)
                l_grpdate = pd.to_datetime(grp_date)
                if l_grpdate > l_hiredate :    #hire date earlier than group date,  return error
                    return 1
                else:
                    return 0
        except Exception as e:
            print('Comparing hire date vs group date', 'Hire date:',hire_date, 'Group Date',grp_date)
            print('error log', e)
            return 1        # error occurs

    def check_seniority_date(self,hire_date, grp_date,seniority_date):
        try:

            if pd.isna(seniority_date) and not pd.isna(hire_date):
                return 1
            if pd.isna(seniority_date) and pd.isna(hire_date):
                return 0
            else:
                l_hiredate = pd.to_datetime(hire_date)
                l_grpdate = pd.to_datetime(grp_date)
                l_seniority_date = pd.to_datetime(seniority_date)
                if l_seniority_date > l_grpdate or l_seniority_date > l_hiredate :    #hire date earlier than group date,  return error
                    return 1
                else:
                    return 0
        except Exception as e:
            print('Comparing Seniority Date: vs hire date & group date', 'Hire date:',hire_date, 'Group Date',grp_date)
            print('error log', e)
            return 1        # error occurs

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
            return 'CHN&HK'
        else:
            return 'Others'

    def check_email(self,email):
        if len(str(email)) > 0:
            lv_mail = str(email).lower().strip()
        else:
            return 0
        try:
            # if re.match(r'^([a-zA-Z.-]+)@zf.com$',lv_mail) or re.match(r'^([a-zA-Z.-]+)@trw.com$',lv_mail)\
            #         or re.match(r'^([a-zA-Z.-]+)@zf[0-9a-zA-Z-]{0,20}\.com$',lv_mail)\
            #         or re.match(r'^([a-zA-Z.-]+)@aac.co.th$',lv_mail)\
            #         or re.match(r'^([a-zA-Z.-]+)@fmg[0-9a-zA-Z-]{0,20}\.com$',lv_mail)\
            #         :
            # match zf.com/trw.com/zf*.com/aac.co.th.com/fmg*.com
            if re.match(r'(^([a-zA-Z]+)+(.\w+)+(.)+(.\w+)*)@(zf|trw|zf+(.-\w+)|aac.co.th|fmg+([-.\w]+))+(.com)$',
                        lv_mail):
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


    def get_bcs_data(self):
        df = pd.DataFrame()

        df_dict = pd.DataFrame()

        file = self.root +  self.bcs_data
        try:
            df = pd.DataFrame(pd.read_excel(io=file, sheet_name='DataQuality', header=0, skiprows=7))


            print('bcs data',df.head(2))
            # print(df.columns)

            df['RU'] = df['A02']

            # Employee number, not FTE
            df['BCS_EMPs'] = df['A16']
            df['BCS_EMPs'] = df.apply((lambda x: round(x['BCS_EMPs'],0)),axis=1)

            df['SF_EMPs'] = df['A17']

            df['BCS_EMPs_DT'] = df['A23']
            df['SF_EMPs_DT'] = df['A24']

            df['BCS_EMPs_ID'] = df['A30']
            df['SF_EMPs_ID'] = df['A31']

            # External Workers (regulars + fixed-term) + Apprentices + Intern / Grad. Student + Vacation Worker / Temp.
            df['BCS_Others'] = df['A38'] + df['A52'] + df['A58'] + df['A67']
            df['SF_Others'] = df['A39'] + df['A53'] + df['A59'] + df['A68']

            df['GAPS_EMP'] = df['BCS_EMPs'] - df['SF_EMPs']
            df['GAPS_OTHERS'] = df['BCS_Others'] - df['SF_Others']

            df['ABS_GAPS_EMP'] = abs(df['GAPS_EMP'])
            df['ABS_GAPS_OTHERS'] = abs(df['GAPS_OTHERS'])

            df['DIV'] = df.apply((lambda x: self.get_div(x['A05'])),axis=1)

            df_x = df.loc[:,
                   ['DIV','RU', 'BCS_EMPs', 'SF_EMPs', 'BCS_EMPs_DT', 'SF_EMPs_DT', 'BCS_EMPs_ID', 'SF_EMPs_ID', 'BCS_Others',
                    'SF_Others','GAPS_EMP','ABS_GAPS_EMP','GAPS_OTHERS','ABS_GAPS_OTHERS']]

            print(df_x.columns)
            print(df_x.head(2))
        except Exception as e:
            print('Exception:', file, e)
        # print(df.columns)

        df_dict = df_x[['RU','DIV']]
        for ru, div in df_dict.itertuples(index=False):
            self.result[ru] = div
            # print(ru, div)
        # print(self.result['200704'])

        # print(result['300018'])
        return df_x.fillna('')



    def get_emp_data(self):
        df_raw = pd.DataFrame()
        df = pd.DataFrame()

        emp_file = self.root + self.empfile

        try:
            df_raw = pd.DataFrame(pd.read_excel(io=emp_file, sheet_name='Excel Output', dtype = {'ZF Global ID':str,'Admin Group (ID)':str},header=0, skiprows=2))
            # print(df.columns)
            df = df_raw[['ZF Global ID','ZID',	'Local ID',	'Employee Status (Label)',	'First Name',	'Last Name',	'Alternate First Name',
                    'Alternate Last Name',	'Date Of Birth',	'Hire Date'	,'Original Start Date',	'Seniority Start Date',	'Termination Date'	,
                    'Solid Line Manager ID',	'Solid Line Manager (Last Name)',	'Solid Line Manager Position',	'Solid Line Manager (First Name)',
                    'In-country Manager Global ID',	'In-country Manager Position',	'In-country Manager First Name',	'In-country Manager Last Name'	,
                    'BP Global ID',	'BP Position',	'BP First Name',	'BP Last Name',
                    'Matrix Manage Global ID',	'Matrix Manager Position',	'Matrix Manage First Name',	'Matrix Manager Last Name',
                    'Event (Label)',	'Event Reason Icode (Label)',	'Position Code',	'Position Title',	'Job Classification (Job Code)',
                    'Job Classification (Label)',	'Company (Label)',	'Board short text',	'Division Short Text',	'BU, Divisional Function/GDF (Label)',	'BU Short Text',
                    'Department (Label)',	'Department Short Text',	'Reporting Unit (Reporting Unit ID)',
                    'Employee Class (Label)',	'Employment Type (Label)','Regular/Limited Employment (Label)',
                    'Location (Name)',	'Location Group (Name)',	'Business email address',
                    'Country (ID)', 'Admin Group (ID)',
                    'External Agency Worker ID',	'External Agency Worker',	'Contingent Worker (External Code)',
                    'Employee Status (External Code)',	'Gender']]
            df = df.rename(columns={"Reporting Unit (Reporting Unit ID)": "RU",
                                    "Regular/Limited Employment (Label)": "Regular_Limited Employment (Label)",
                                    'BU, Divisional Function/GDF (Label)':'BU_Divisional Function_GDF (Label)'
                                    })

            df['chk_status'] = df.apply(lambda x:self.get_status(x['Employee Status (Label)']),axis=1)
            df['External'] = df.apply((lambda x: 1 if x['External Agency Worker'] == 'Yes' else 0 ),axis=1)
            df['Intern'] = df.apply((lambda x: 1 if x['Employment Type (Label)'] == 'Intern/Students' else 0 ),axis=1)
            df['Apprentices'] = df.apply((lambda x: 1 if x['Employment Type (Label)'] == 'Apprentice' else 0 ),axis=1)
            df['VacationWorker'] = df.apply((lambda x: 1 if x['Employment Type (Label)'] == 'Vacation Workers DE' else 0 ),axis=1)
            df['Total_EMP'] = df.apply((lambda x: 1),axis=1)

            # df['AP'] = df.apply(lambda x:get_region(x['Country (ID)']),axis=1)
            df['EEDirect'] = df.apply(lambda x:self.get_direct(x['External Agency Worker'],x['Employment Type (Label)'],x['Employee Class (Label)'],x['ZF Global ID']),axis=1)
            df['EEIndirect'] = df.apply(lambda x:self.get_indirect(x['External Agency Worker'],x['Employment Type (Label)'],x['Employee Class (Label)'],x['ZF Global ID']),axis=1)
            # df['Female'] = df.apply(lambda x:get_gender(x['Gender'], "X"),axis=1)


            df['Male'] = df.apply(lambda x:self.get_gender(x['Gender'], ""),axis=1)

            df['Division_rs'] = df.apply(lambda x:self.get_div(x['Division Short Text']),axis=1)

            df['email_check'] = df.apply(lambda x:self.check_email(x['Business email address']),axis=1)
            df['ZID_check'] = df.apply(lambda x:self.check_ZID(x['EEIndirect'],x['ZID']),axis=1)

            df['SM_Chck'] = df.apply((lambda x: 1 if not pd.isnull(x['Solid Line Manager (Last Name)'])  else 0), axis=1)
            df['IM_Chck'] = df.apply((lambda x: 1 if not pd.isnull(x['In-country Manager Last Name']) else 0), axis=1)
            df['BP_Chck'] = df.apply((lambda x: 1 if not pd.isnull(x['BP Last Name'])  else 0), axis=1)

            df['ID_email_check'] = df.apply(lambda x:self.check_indirect(x['EEIndirect'],x['email_check']),axis=1)

            df['DATE_CHK'] = df.apply(lambda x: self.check_date(x['Hire Date'],x['Original Start Date']), axis=1)
            df['Seniority_DATE_CHK'] = df.apply(lambda x: self.check_seniority_date(x['Hire Date'],x['Original Start Date'],x['Seniority Start Date']), axis=1)
            df['AGE_RANGE'] = df.apply(lambda x: self.get_age_range(x['Date Of Birth'],), axis=1)
            df['AgeLess18'] = df.apply(lambda x: 1 if (x['AGE_RANGE'] == '0--18' or x['AGE_RANGE'] == 'Unkown' or x['AGE_RANGE'] == '80+') else 0, axis=1)
            df['EMAIL_ZID_CHK'] = df.apply(lambda x: self.check_email_zid(x['email_check'],x['ZID_check']), axis=1)

            df = df[df['chk_status'] == 'A']

            print(df.head(2))

            # df['External'] = df.apply(lambda x:get_hire_chg(x['External Agency & Contingent Worker'],x['Event Reason Icode (Label)']),axis=1)
            # print(df.columns)
        except Exception as e:
            print('Exception:', emp_file,e)
        # print(df.columns)

        return df.fillna('')



    def out_mails(self, df_data, df_writer, workbook,sheet_name,common_format):
        emails_ap = pd.DataFrame()

        emails_ap = df_data.groupby(['DIV'])[
            'Total_EMP', 'email_check','EMAIL_ZID_CHK', 'EEIndirect', 'ID_email_check' ].sum().reset_index()

        emails_ap = emails_ap.rename(columns={"email_check": "Employee_with_Emails",
                                })

        emails_ap['Rate'] = emails_ap['Employee_with_Emails'] / emails_ap['Total_EMP']
        emails_ap['IndirectRate'] = emails_ap['ID_email_check'] / emails_ap['EEIndirect']

        print(emails_ap.head())

        emails_ap = emails_ap[['DIV', 'Total_EMP', 'Employee_with_Emails', 'Rate','EEIndirect','ID_email_check','IndirectRate']]


        x, y = emails_ap.shape
        print('df_datagroup', x, y)
        emails_ap.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]

        worksheet.set_column("D:D", cell_format=common_format)
        worksheet.set_column("G:G", cell_format=common_format)
        # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
        chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
        for i in range(1, 3):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
            })

        line_chart = workbook.add_chart({'type': 'line'})
        # Configure the data series for the secondary chart.
        line_chart.add_series({
            'name': [sheet_name, 0, 3],
            'categories': [sheet_name, 1, 0, x, 0],
            'values': [sheet_name, 1, 3, x, 3],
            # 'data_labels': {'value': True,'num_format': "9"},
            'data_labels': {'value': True, 'num_format': "0.0%"},
            'y2_axis': True,
        })
        # Combine the charts.
        chart2.combine(line_chart)

        # chart2.set_size({'x_scale': 1.5, 'y_scale': 1})
        chart2.set_size({'width': 540, 'height': 300})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(10)
        # chart2.set_style(48)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(8, 2, chart2, {'x_offset': 50, 'y_offset': 100})



    def out_chart(self, df_data, df_writer, workbook, chart2, sheet_name, title, xtitle, ytitle):
        df_headcount_ana = df_data.groupby(['DIV'])[
            'BCS_EMPs', 'SF_EMPs', 'GAPS_EMP', 'ABS_GAPS_EMP'].sum().reset_index()
        df_headcount_ana.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        x, y = df_headcount_ana.shape
        print('df_datagroup', x, y)
        # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
        chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
        for i in range(1,3):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                # 'data_labels': {'series_name': True, 'position': 'below'},
            })

        line_chart = workbook.add_chart({'type': 'line'})
        # Configure the data series for the secondary chart.
        line_chart.add_series({
            'name': [sheet_name, 0, 3],
            'categories': [sheet_name, 1, 0, x, 0],
            'values': [sheet_name, 1, 3, x, 3],
            'data_labels': {'value': False},
            'y2_axis':    True,
        })
        line_chart.add_series({
            'name': [sheet_name, 0, 4],
            'categories': [sheet_name, 1, 0, x, 0],
            'values': [sheet_name, 1, 4, x, 4],
            'data_labels': {'value': True},
            'y2_axis':    True,
        })
        # Combine the charts.
        chart2.combine(line_chart)

        # Add a chart title and some axis labels.
        # chart2.set_title({'name': title})
        # chart2.set_x_axis({'name': xtitle})
        # chart2.set_y_axis({'name': ytitle})

        # chart2.set_size({'x_scale': 1.5, 'y_scale': 1})
        chart2.set_size({'width': 540, 'height': 300})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(10)
        # chart2.set_style(48)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(8, 2, chart2, {'x_offset': 50, 'y_offset': 100})


    def overall_data(self,df_data, df_writer, workbook,sheet_name):
        df_emp_rs = df_data.groupby(['Country (ID)','DIV','RU'])['Total_EMP','email_check','ZID_check','SM_Chck','IM_Chck','BP_Chck','EMAIL_ZID_CHK','DATE_CHK',
                                    'EEIndirect','ID_email_check'].sum().reset_index()
        df_emp_rs.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        sheet_name_new = sheet_name + '_div'
        df_emp_rs_div = df_data.groupby(['DIV'])['Total_EMP','email_check','ZID_check','SM_Chck','IM_Chck','BP_Chck','EMAIL_ZID_CHK','DATE_CHK',
                                    'EEIndirect','ID_email_check'].sum().reset_index()
        df_emp_rs_div['Email_Rate'] = df_emp_rs_div.apply(lambda x: x['email_check'] / x['Total_EMP'], axis=1)
        df_emp_rs_div['ID_Email_Rate'] = df_emp_rs_div.apply(lambda x: x['ID_email_check'] / x['EEIndirect'], axis=1)
        df_emp_rs_div['SM_GAP'] = df_emp_rs_div.apply(lambda x: x['Total_EMP'] - x['SM_Chck'], axis=1)
        df_emp_rs_div['IM_GAP'] = df_emp_rs_div.apply(lambda x: x['Total_EMP'] - x['IM_Chck'], axis=1)
        df_emp_rs_div['BP_GAP'] = df_emp_rs_div.apply(lambda x: x['Total_EMP'] - x['BP_Chck'], axis=1)
        df_emp_rs_div['SM_RATE'] = df_emp_rs_div.apply(lambda x: x['SM_Chck']/x['Total_EMP'], axis=1)
        df_emp_rs_div['IM_RATE'] = df_emp_rs_div.apply(lambda x: x['IM_Chck']/x['Total_EMP'], axis=1)
        df_emp_rs_div['BP_RATE'] = df_emp_rs_div.apply(lambda x: x['BP_Chck']/x['Total_EMP'], axis=1)
        df_emp_rs_div['Date_MISS_MATCH'] = df_emp_rs_div.apply(lambda x: x['DATE_CHK']/x['Total_EMP'], axis=1)
        df_emp_rs_div.to_excel(df_writer, sheet_name=sheet_name_new, encoding="utf-8", index=False)


    def chk_zid_mail(self, df_data, df_writer, workbook,sheet_name):

        df_res_tmp = df_data[df_data['Country'] == 'CHN&HK']

        data_ap = df_data.groupby(['DIV'])['EMAIL_ZID_CHK'].sum().reset_index()
        data_ap = data_ap.rename(columns={"EMAIL_ZID_CHK": "Data_AP"})

        data_chn = df_res_tmp.groupby(['DIV'])['EMAIL_ZID_CHK'].sum().reset_index()
        data_chn = data_chn.rename(columns={"EMAIL_ZID_CHK": "Data_CHN"})
        # print(data_res.head())
        data_res = pd.merge(data_ap,data_chn)
        print("pd_merge",data_res)

        x, y = data_res.shape
        print('df_datagroup', x, y)
        data_res.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
        chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
        for i in range(1, 3):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                'data_labels': {'value': True},
            })

        # chart2.set_size({'x_scale': 1.5, 'y_scale': 1})
        chart2.set_size({'width': 540, 'height': 300})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(10)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(8, 2, chart2, {'x_offset': 50, 'y_offset': 100})


    def get_gap(self,total, value):
        return total - value


    def chk_mgr(self, df_data, df_writer, workbook,sheet_name,common_format):
        data_res = pd.DataFrame()

        data_res = df_data.groupby(['DIV'])['Total_EMP', 'SM_Chck','IM_Chck', 'BP_Chck'].sum().reset_index()

        data_res['SM_GAP'] = data_res.apply((lambda x: self.get_gap(x['Total_EMP'],x['SM_Chck'])), axis=1)
        data_res['IM_GAP'] = data_res.apply((lambda x: self.get_gap(x['Total_EMP'],x['IM_Chck'])), axis=1)
        data_res['BP_GAP'] = data_res.apply((lambda x: self.get_gap(x['Total_EMP'],x['BP_Chck'])), axis=1)

        data_res['SM_RATE'] = data_res.apply((lambda x: (x['SM_Chck']/x['Total_EMP'])), axis=1)
        data_res['IM_RATE'] = data_res.apply((lambda x: (x['IM_Chck']/x['Total_EMP'])), axis=1)
        data_res['BP_RATE'] = data_res.apply((lambda x: (x['BP_Chck']/x['Total_EMP'])), axis=1)

        data_res2 = data_res[['DIV','SM_GAP','IM_GAP','BP_GAP','SM_RATE','IM_RATE','BP_RATE']]

        print(data_res.head(1))

        x, y = data_res2.shape
        print('df_datagroup', x, y)
        data_res2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]

        worksheet.set_column("E:G", cell_format=common_format)
        # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
        chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
        for i in range(1, 4):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
            })

        line_chart = workbook.add_chart({'type': 'line'})
        # Configure the data series for the secondary chart.

        for i in range(4, 7):
            line_chart.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                # 'data_labels': {'value': True,'num_format': "9"},
                'data_labels': {'value': False, 'num_format': "0%"},
                'y2_axis': True,
            })
        # Combine the charts.
        chart2.combine(line_chart)

        # chart2.set_size({'x_scale': 1.5, 'y_scale': 1})
        chart2.set_size({'width': 540, 'height': 300})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(10)
        # chart2.set_style(48)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(8, 2, chart2, {'x_offset': 50, 'y_offset': 100})


    def check_date_match(self, df_data, df_writer, workbook,sheet_name,common_format):
        data_res = pd.DataFrame()

        data_res = df_data.groupby(['DIV'])['Total_EMP', 'DATE_CHK'].sum().reset_index()

        data_res['MISS_MATCH_RATE'] = data_res.apply((lambda x: (x['DATE_CHK']/x['Total_EMP'])), axis=1)

        data_res2 = data_res[['DIV','DATE_CHK','MISS_MATCH_RATE']]

        print(data_res.head(1))

        x, y = data_res2.shape
        print('df_datagroup', x, y)
        data_res2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        worksheet.set_column("C:C", cell_format=common_format)
        # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
        chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
        for i in range(1, 2):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                'data_labels': {'value': True},
            })

        line_chart = workbook.add_chart({'type': 'line'})
        # Configure the data series for the secondary chart.

        for i in range(2, 3):
            line_chart.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                # 'data_labels': {'value': True,'num_format': "9"},
                'data_labels': {'value': False, 'num_format': "0%"},
                'y2_axis': True,
            })
        # Combine the charts.
        chart2.combine(line_chart)

        # chart2.set_size({'x_scale': 1.5, 'y_scale': 1})
        chart2.set_size({'width': 500, 'height': 300})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(10)
        # chart2.set_style(48)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(8, 2, chart2, {'x_offset': 50, 'y_offset': 100})

    def check_s_date_match(self, df_data, df_writer, workbook,sheet_name,common_format):
        data_res = pd.DataFrame()

        data_res = df_data.groupby(['DIV'])['Total_EMP', 'Seniority_DATE_CHK'].sum().reset_index()

        data_res['MISS_MATCH_RATE'] = data_res.apply((lambda x: (x['Seniority_DATE_CHK']/x['Total_EMP'])), axis=1)

        data_res2 = data_res[['DIV','Seniority_DATE_CHK','MISS_MATCH_RATE']]

        print(data_res.head(1))

        x, y = data_res2.shape
        print('df_datagroup', x, y)
        data_res2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        worksheet.set_column("C:C", cell_format=common_format)
        # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
        chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
        for i in range(1, 2):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                'data_labels': {'value': True},
            })

        line_chart = workbook.add_chart({'type': 'line'})
        # Configure the data series for the secondary chart.

        for i in range(2, 3):
            line_chart.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                # 'data_labels': {'value': True,'num_format': "9"},
                'data_labels': {'value': False, 'num_format': "0%"},
                'y2_axis': True,
            })
        # Combine the charts.
        chart2.combine(line_chart)

        # chart2.set_size({'x_scale': 1.5, 'y_scale': 1})
        chart2.set_size({'width': 500, 'height': 300})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(10)
        # chart2.set_style(48)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(8, 2, chart2, {'x_offset': 50, 'y_offset': 100})


    def check_age(self, df_data, df_writer, workbook,sheet_name):
        data_res = pd.DataFrame()

        data_res = df_data.groupby(['DIV'])['AgeLess18'].sum().reset_index()
        x, y = data_res.shape
        print('df_datagroup', x, y)
        data_res.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        worksheet = df_writer.sheets[sheet_name]
        # chart2 = workbook.add_chart({'type': 'bar'})        #'subtype': 'stacked'
        chart2 = workbook.add_chart({'type': 'column'})  # ''subtype': 'percent_stacked'
        for i in range(1, 2):
            chart2.add_series({
                'name': [sheet_name, 0, i],
                'categories': [sheet_name, 1, 0, x, 0],
                'values': [sheet_name, 1, i, x, i],
                'data_labels': {'value': True},
            })

        line_chart = workbook.add_chart({'type': 'line'})
        # Configure the data series for the secondary chart.

        # chart2.set_size({'x_scale': 1.5, 'y_scale': 1})
        chart2.set_size({'width': 500, 'height': 300})
        # Set an Excel chart style. Colors with white outline and shadow.
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(10)
        # chart2.set_style(48)
        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart(8, 2, chart2, {'x_offset': 50, 'y_offset': 100})

    def init_format(self,workbk):
        self.common_format = workbk.add_format(
            {'align': 'right', 'valign': 'vcenter', 'text_wrap': True, 'num_format': "0.0%"})
        self.one_format = workbk.add_format(
            {'align': 'right', 'valign': 'vcenter', 'text_wrap': True, 'num_format': "0%"})
        self.date_value_format = workbk.add_format({'num_format':'yyyy/mm/dd', })

    def out_raw_data(self,folder,field,range,df_headcount,df_emails,df_mgr,df_age18,df_entry_date,df_s_entry_date):

        df_div_ru = pd.DataFrame()
        df_div_email = pd.DataFrame()
        df_div_mgr = pd.DataFrame()
        df_div_age = pd.DataFrame()

        for lv_value in range:
            lv_value = str(lv_value)
            try:
                file_name = 'Output_' + lv_value + '_' + self.now_date + '_raw.xlsx'
                # 创建一个excel
                df_writer = pd.ExcelWriter(self.root + folder + file_name, engine='xlsxwriter')
                workbook = df_writer.book
                self.init_format(workbook)

                # df_div_ru = df_RU_res[df_RU_res['RU'] == division]
                # sheet_name = '10_RU'
                # df_div_ru.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

                df_hd_ru = df_headcount[df_headcount[field] == lv_value]
                sheet_name = '120_HeadCount'
                df_hd_ru.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
                worksheet = df_writer.sheets[sheet_name]
                worksheet.set_column("J:J", cell_format=self.common_format)

                df_div_emails = df_emails[df_emails[field] == lv_value]
                if not df_div_emails.empty:
                    sheet_name = '130_Emails_InvalidZID'
                    df_div_emails.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

                df_div_mgr = df_mgr[df_mgr['DIV'] == lv_value]
                sheet_name = '140_MGR_BP'
                if not df_div_mgr.empty:
                    df_div_mgr.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

                df_div_entry_date = df_entry_date[df_entry_date[field] == lv_value]
                if not df_div_entry_date.empty:
                    sheet_name = '150_EntryDate'
                    df_div_entry_date.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
                    worksheet_h = df_writer.sheets[sheet_name]
                    worksheet_h.set_column("J:K", width=18, cell_format=self.date_value_format)

                # df_s_div_entry_date = df_s_entry_date[df_s_entry_date[field] == lv_value]
                # if not df_s_div_entry_date.empty:
                #     sheet_name = '160_SeniorityDate'
                #     df_s_div_entry_date.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
                #     # worksheet_h = df_writer.sheets[sheet_name]
                #     # worksheet_h.set_column("J:K", width=18, cell_format=self.date_value_format)

                df_div_age = df_age18[df_age18[field] == lv_value]
                if not df_div_age.empty:
                    sheet_name = '170_Age18'
                    df_div_age.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
                    worksheet = df_writer.sheets[sheet_name]
                    worksheet.set_column("J:J", width=18, cell_format=self.date_value_format)

                workbook.close()
                df_writer.close()

            except Exception as e:
                print('write file failed:', file_name, sheet_name, lv_value)
                print('error log', e)
                df_writer.close()

    def quality_check_summary(self):
        df_bcs_data = self.get_bcs_data()
        print('----------------------------------------------------')
        df_data = self.get_emp_data()
        print('employee data', df_data.columns)
        print('employee data', df_data.head())
        file_out = 'Output_' + self.now_date + '_ap_factors.xlsx'
        file_sum = 'Output_' + self.now_date + '_ap_summary.xlsx'
        file_raw = 'Output_' + self.now_date + '_ap_raw.xlsx'

        df_simple = pd.DataFrame()

        df_res = pd.DataFrame()
        df_res_chn = pd.DataFrame()
        df_res2 = pd.DataFrame()

        df_RU_res = pd.DataFrame()
        df_headcount = pd.DataFrame()
        df_emails = pd.DataFrame()
        df_mgr = pd.DataFrame()
        df_entry_date = pd.DataFrame()
        df_s_entry_date = pd.DataFrame()
        df_age18 = pd.DataFrame()

        df_data['EETotal'] = df_data['EEDirect'] + df_data['EEIndirect']
        df_data['DIV'] = df_data.apply(lambda x: self.get_div_from_bcs(x['RU'],x['Division Short Text']), axis=1)

        list_admin = list(set(df_data['Admin Group (ID)'].tolist()))

        # All
        df_simple = df_data.groupby(['Country (ID)', 'Company (Label)', 'RU'])[
            'Total_EMP','EETotal', 'EEDirect', 'EEIndirect', 'External', 'Intern', 'Apprentices', 'VacationWorker', 'email_check','ZID_check','SM_Chck','IM_Chck',
            'BP_Chck','DATE_CHK','Seniority_DATE_CHK','EMAIL_ZID_CHK','ID_email_check','AgeLess18'].sum().reset_index()

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

        df_res['SM_GAP'] = df_res.apply(lambda x: x['Total_EMP'] - x['SM_Chck'], axis=1)
        df_res['IM_GAP'] = df_res.apply(lambda x: x['Total_EMP'] - x['IM_Chck'], axis=1)
        df_res['BP_GAP'] = df_res.apply(lambda x: x['Total_EMP'] - x['BP_Chck'], axis=1)

        df_res_chn = df_res[df_res['Country'] == 'CHN&HK']
        df_res_other = df_res[df_res['Country'] != 'CHN&HK']

        list_cc = list(set(df_res['Company (Label)'].tolist()))
        # print(list_cc)

        df_res2 = df_res.groupby(['Country', 'Country (ID)', 'DIV' , 'RU' ])[
            'Total_EMP','EETotal','EEIndirect','email_check','ZID_check','SM_Chck','IM_Chck','BP_Chck','ID_email_check',
            'SM_GAP','IM_GAP','BP_GAP', 'AgeLess18',
             'DATE_CHK','Seniority_DATE_CHK','EMAIL_ZID_CHK','BCS_EMPs','SF_EMPs','BCS_Others','SF_Others',
            'GAPS_EMP','ABS_GAPS_EMP','GAPS_OTHERS','ABS_GAPS_OTHERS'].sum().reset_index()


        # df_res2_group = df_res2.getgroup
        df_RU_res = df_res.groupby(['Country (ID)', 'Company (Label)','DIV' , 'RU'])['Total_EMP'].sum().reset_index()
        df_RU_res.sort_values(by = ['Country (ID)','Company (Label)','DIV',"Total_EMP"], ascending=[True,True,True,False],inplace=True)

        df_headcount = df_res.groupby(['Country (ID)', 'Company (Label)','DIV' , 'RU'])['Total_EMP','BCS_EMPs','SF_EMPs','GAPS_EMP','ABS_GAPS_EMP'].sum().reset_index()
        df_headcount['HD_GAP_RATE'] = df_headcount.apply(lambda x: 0 if x['BCS_EMPs'] < 1 else round(x['ABS_GAPS_EMP'] / x['BCS_EMPs'],3), axis=1)
        df_headcount.sort_values(by = ["ABS_GAPS_EMP"], ascending=[False], inplace=True)

        df_emails = df_data[['Country (ID)','Company (Label)','DIV','RU','Admin Group (ID)','ZF Global ID', 'ZID', 'Employee Status (Label)','First Name','Last Name',
                             'Business email address', 'email_check','ZID_check','EMAIL_ZID_CHK']]
        df_emails = df_emails.loc[(df_emails['EMAIL_ZID_CHK'] > 0) ]

        df_mgr = df_data[['Country (ID)','Company (Label)','DIV','RU','Admin Group (ID)','ZF Global ID', 'ZID', 'Employee Status (Label)','First Name','Last Name',
                          'Solid Line Manager ID','Solid Line Manager Position','Solid Line Manager (Last Name)', 'Solid Line Manager (First Name)',
                          'In-country Manager Global ID','In-country Manager Position','In-country Manager First Name','In-country Manager Last Name',
                          'BP Global ID','BP Position','BP First Name','BP Last Name',
                          'SM_Chck','IM_Chck','BP_Chck'#
        ]]
        df_mgr = df_mgr.loc[(df_mgr['SM_Chck'] == 0) | (df_mgr['IM_Chck'] == 0) | (df_mgr['BP_Chck'] == 0) ]

        df_entry_date = df_data[['Country (ID)','Company (Label)','DIV','RU','Admin Group (ID)','ZF Global ID', 'ZID', 'Employee Status (Label)','First Name','Last Name',
                            'Hire Date', 'Original Start Date','DATE_CHK']]
        df_entry_date = df_entry_date.loc[(df_entry_date['DATE_CHK'] > 0) ]

        df_s_entry_date = df_data[['Country (ID)','Company (Label)','DIV','RU','Admin Group (ID)','ZF Global ID', 'ZID', 'Employee Status (Label)','First Name','Last Name',
                            'Hire Date', 'Original Start Date','Seniority Start Date','Seniority_DATE_CHK']]
        df_s_entry_date = df_s_entry_date.loc[(df_s_entry_date['Seniority_DATE_CHK'] > 0) ]

        df_age18 = df_data[['Country (ID)','Company (Label)','DIV','RU','Admin Group (ID)','ZF Global ID', 'ZID', 'Employee Status (Label)','First Name','Last Name',
                            'Date Of Birth','AgeLess18']]
        df_age18 = df_age18.loc[(df_age18['AgeLess18'] > 0) ]

        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(self.root+file_out,engine='xlsxwriter')
            workbook = df_writer.book
            self.init_format(workbook)

            # chart2 = workbook.add_chart({'type': 'column'})        #'subtype': 'percent_stacked'

            # sheet_name = '20_DIV_C'
            # pvt_tmp = pd.pivot_table(df_res2, index=['Country (ID)'], values=["RU"],
            #                          aggfunc={np.count_nonzero}, fill_value=0)
            # pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")

            sheet_name = '10_Initial'
            df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '12_EMP_AP_All'
            self.overall_data(df_data, df_writer, workbook,sheet_name)

            sheet_name = '14_EMP_CHN_All'
            self.overall_data(df_data[df_data['Country (ID)']=='CHN'], df_writer, workbook,sheet_name)

            sheet_name = '15_DIV_O'
            pvt_tmp = pd.pivot_table(df_res2, index=['DIV'], values=["RU"],
                                     aggfunc={np.count_nonzero}, fill_value=0)
            pvt_tmp.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8")


            sheet_name = '20_bcs_initial'
            df_bcs_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            # sheet_name = 'AP_20_division'
            # df_division.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '20_EMP_SUM'
            df_simple.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            # sheet_name = '20_res'
            # df_res.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '30_res_simple'
            df_res2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            # sheet_name = '100_RU'
            # df_RU_res.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            #
            # sheet_name = '110_Headcount'
            # df_headcount.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            #
            # sheet_name = '120 Emails&ZID'
            # df_emails.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            #
            # sheet_name = '130MGR&BP'
            # df_mgr.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            # based on simple data to have analyze chart
            sheet_name = '210_Headcount_AP'
            chart2 = {}
            self.out_chart(df_res, df_writer, workbook, chart2,  sheet_name, 'Headcount AP-Employees', 'By Division', 'Employee No')

            sheet_name = '212_Headcount_CHN'
            chart2 = {}
            self.out_chart(df_res_chn, df_writer, workbook, chart2,  sheet_name, 'Headcount AP-Employees', 'By Division', 'Employee No')

            sheet_name = '214_Headcount_Other'
            chart2 = {}
            self.out_chart(df_res_other, df_writer, workbook, chart2,  sheet_name, 'Headcount AP-Employees', 'By Division', 'Employee No')

            sheet_name = '220_Email_ap'
            self.out_mails(df_data, df_writer, workbook,sheet_name,self.common_format)
            sheet_name = '212_Email_CHN'
            self.out_mails(df_res_chn, df_writer, workbook,sheet_name,self.common_format)

            sheet_name = '220_Email_ZID'
            self.chk_zid_mail(df_res, df_writer, workbook,sheet_name)

            sheet_name = '230_MGR_AP'
            self.chk_mgr(df_res, df_writer, workbook,sheet_name,self.one_format)
            sheet_name = '232_MGR_CHN'
            self.chk_mgr(df_res_chn, df_writer, workbook,sheet_name,self.one_format)

            sheet_name = '240_DATE_AP'
            self.check_date_match(df_res, df_writer, workbook,sheet_name,self.common_format)
            sheet_name = '242_DATE_CHN'
            self.check_date_match(df_res_chn, df_writer, workbook,sheet_name,self.common_format)

            sheet_name = '244_Seniority_DATE_AP'
            self.check_s_date_match(df_res, df_writer, workbook,sheet_name,self.common_format)
            sheet_name = '246_Seniority_DATE_CHN'
            self.check_s_date_match(df_res_chn, df_writer, workbook,sheet_name,self.common_format)

            sheet_name = '250_AGE_AP'
            self.check_age(df_res, df_writer, workbook,sheet_name)

            workbook.close()
            df_writer.close()
        except Exception as e:
            print('write file failed:', file_out, sheet_name)
            print('error log', e)
            workbook.close()
        finally:
            df_writer.close()

        print('Prepare summary Data in one  ----------------------------------------------------')
        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(self.root+file_sum,engine='xlsxwriter')
            sheet_name = '10_Initial'
            df_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '20_bcs_initial'
            df_bcs_data.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            sheet_name = '30_res_simple'
            df_res2.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

        except Exception as e:
            print('write file failed:', file_sum, sheet_name)
            print('error log', e)
        finally:
            workbook.close()
            df_writer.close()

        print('Prepare RAW Data in one  ----------------------------------------------------')
        # RAW Data output
        try:
            # 创建一个excel
            df_writer = pd.ExcelWriter(self.root+file_raw,engine='xlsxwriter',date_format='yyyy-mm-dd')
            workbook = df_writer.book
            self.init_format(workbook)
            sheet_name = '110_RU'
            df_RU_res.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '120_Headcount'
            df_headcount.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            worksheet = df_writer.sheets[sheet_name]
            worksheet.set_column("J:J", cell_format=self.common_format)

            sheet_name = '130_Emails&ZID'
            df_emails.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '140MGR&BP'
            df_mgr.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            sheet_name = '150_EntryDate'
            df_entry_date.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False, header=False, startrow=1)
            worksheet = df_writer.sheets[sheet_name]
            for col_num, value in enumerate(df_entry_date.columns.values):
                worksheet.write(0, col_num, value)
            worksheet.set_column("J:K", cell_format=self.date_value_format)

            sheet_name = '160_Seniority_Date'
            df_s_entry_date.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False, header=False, startrow=1)
            worksheet = df_writer.sheets[sheet_name]
            for col_num, value in enumerate(df_s_entry_date.columns.values):
                worksheet.write(0, col_num, value)
            worksheet.set_column("J:K", cell_format=self.date_value_format)

            sheet_name = '170_BirthDate'
            df_age18.to_excel(df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)
            worksheet = df_writer.sheets[sheet_name]
            worksheet.set_column("J:J", cell_format=self.date_value_format)

            workbook.close()
            df_writer.close()
        except Exception as e:
            print('write file failed:', file_raw, sheet_name)
            print('error log', e)
        finally:
            workbook.close()
            df_writer.close()

    #     RAW Data By Division
        print('Prepare RAW Data by Division ----------------------------------------------------')

        df_headcount.sort_values(by=['DIV','Country (ID)','RU'], ascending=True, inplace=True)
        df_emails.sort_values(by=['DIV','Country (ID)','RU'], ascending=True, inplace=True)
        df_mgr.sort_values(by=['DIV','Country (ID)','RU'], ascending=True, inplace=True)
        df_age18.sort_values(by=['DIV','Country (ID)','RU'], ascending=True, inplace=True)
        df_entry_date.sort_values(by=['DIV','Country (ID)','RU'], ascending=True, inplace=True)
        df_s_entry_date.sort_values(by=['DIV','Country (ID)','RU'], ascending=True, inplace=True)
        self.out_raw_data('/DIV/','DIV',self.division,df_headcount,df_emails,df_mgr,df_age18,df_entry_date,df_s_entry_date)

        #     RAW Data By Legal Entity
        print('Prepare RAW Data by Legal Entity ----------------------------------------------------')

        df_headcount.sort_values(by=['Company (Label)','RU'], ascending=True, inplace=True)
        df_emails.sort_values(by=['Company (Label)','RU'], ascending=True, inplace=True)
        df_mgr.sort_values(by=['Company (Label)','RU'], ascending=True, inplace=True)
        df_age18.sort_values(by=['Company (Label)','RU'], ascending=True, inplace=True)
        self.out_raw_data('/LE/','Company (Label)',list_cc,df_headcount,df_emails,df_mgr,df_age18,df_entry_date,df_s_entry_date)

        #     RAW Data By Admin Group
        # print('Prepare RAW Data by Admin Group ----------------------------------------------------')
        # self.out_raw_data('/Admin/','Admin Group (ID)',list_admin,df_headcount,df_emails,df_mgr,df_age18,df_entry_date)


if __name__ == '__main__':
    df_data = pd.DataFrame()

    time1 = time.time()
    obj_factor = ApFactorData()
    obj_factor.quality_check_summary()

    # head_count_summary(df_data, file_out

    print("Total running time", time.time() - time1)

    time.sleep(1)

    print("done")