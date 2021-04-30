from mailmerge import MailMerge
from docx2pdf import convert
import pandas as pd
import time
import os
# from datetime import datetime

class handle_merge():
    def __init__(self):
        self.df_data = pd.DataFrame()
        self.path = 'C:/temp/mm/'
        self.prefix = '劳动关系转移三方协议_'
        self.template = '劳动关系转移三方协议_20201221 transfer JT to CHJ2.docx'
        self.template_fix = '劳动关系转移三方协议_20201221 transfer JT to CHJ2 - fix.docx'
        # self.datafile = 'Moving name list_JT&CHJ_ 202102.xlsx'
        self.datafile = 'Moving name list_JT&CHJ_ 202102 - Copy.xlsx'

    def read_files(self):
        # df_data = pd.DataFrame()
        try:
            self.df_data = pd.DataFrame(pd.read_excel(io=self.path + self.datafile,
                                                      sheet_name='Raw data',
                                                      dtype={'National ID': 'str'},
                                                      header=0,
                                                      # skiprows=2))
                                                      ))
            self.df_data.drop_duplicates(inplace=True)
            self.df_data.fillna("",inplace=True)
            print(self.df_data.head())
            print(self.df_data.shape)
        except Exception as e:
            print('Handle Employee Data Exception:', self.datafile, e)

    def main(self):
        list_columns = ['Global ID','Employee Name', 'Status','Notes']
        df_list = pd.DataFrame(columns=list_columns)
        self.read_files()

        for n_idx, n_row in self.df_data.iterrows():
            if len(str(n_row['Target Move Date'])) > 0:
                if n_row['Template'] == 'fix':
                    lv_template = self.template_fix
                else:
                    lv_template = self.template
                document_1 = MailMerge(self.path + lv_template)
                # print(document_1.get_merge_fields())
                # isinstance(['Target Move Date'], datetime):

                try:
                    lv_target_cn = n_row['Target Move Date'].strftime("%Y年%m月%d日")
                    lv_target_en = n_row['Target Move Date'].strftime("%d/%m/%Y")
                    lv_start_en = n_row['Start date of ZF Service'].strftime("%d/%m/%Y")
                    lv_start_cn = n_row['Start date of ZF Service'].strftime("%Y年%m月%d日")
                    if n_row['Template'] == 'fix':
                        lv_contract_end = ''
                        lv_contract_end_cn = ''
                    else:
                        lv_contract_end = n_row['Latest Contract ending date'].strftime("%d/%m/%Y")
                        lv_contract_end_cn = n_row['Latest Contract ending date'].strftime("%Y年%m月%d日")
                        print(lv_contract_end_cn)
                        # lv_contract_end = n_row['Latest Contract ending date'].to_timestamp().strftime("%d/%m/%Y")
                        # lv_date = time.localtime(time.mktime(time.strptime(n_row['Latest Contract ending date'],'%Y-%m-%d')))
                        # # print(lv_date)
                        # # lv_contract_end = time.strptime("%d/%m/%Y %H:%M:%S",lv_date)
                        # # print(lv_contract_end)
                        # # lv_contract_end_cn = n_row['Latest Contract ending date'].to_timestamp().strftime("%Y年%m月%d日")
                        # # print(lv_contract_end_cn)
                        # lv_contract_end_cn = '至' + str(lv_date.tm_year) +'年' + str(lv_date.tm_mon) + '月' + str(lv_date.tm_mday) + '日止'
                        # lv_contract_end = 'lasts until ' + str(lv_date.tm_mday) + '/' + str(lv_date.tm_mon) + '/' + str(lv_date.tm_year)

                    dict = {
                        "Chinese_Name": str(n_row['Chinese Name']),
                        "Employee_Name": str(n_row['Employee Name']),
                        "National_ID": str(n_row['National ID']),
                        #                    "Start_date_of_ZF_Service": str(n_row['Start date of ZF Service']),
                        #                    "Target_Move_Date": str(n_row['Target Move Date'])
                        "Target_Move_Date": lv_target_en,
                        "Target_Move_Date_CN": lv_target_cn,
                        "Start_date_of_ZF_Service": lv_start_en,
                        "Start_date_of_ZF_Service_CN": lv_start_cn,
                        "Contract_End_CN": lv_contract_end_cn,
                        "Contract_End": lv_contract_end,
                    }
                    document_1.merge(**dict)

                    file_path = self.path + 'xlsx/'
                    if not os.path.exists(file_path):
                        os.makedirs(file_path)

                    lv_name = self.prefix + n_row['Division for list'] + '_' + n_row['Employee Name']
                    file_docx = self.path + 'xlsx/' + lv_name + '.docx'
                    document_1.write(file_docx)
                    # time.sleep(1)

                    file_path = self.path + 'pdf' + '/' + n_row['HR BP2'] + '/'
                    if not os.path.exists(file_path):
                        os.makedirs(file_path)

                    if n_row['Mark'] != "":
                        file_path = file_path + n_row['Mark'] + '/'
                        if not os.path.exists(file_path):
                            os.makedirs(file_path)

                    #                file_pdf = file_path + 'pdf/' + lv_name + '.pdf'
                    file_pdf = file_path + lv_name + '.pdf'
                    # file_pdf = self.path + 'pdfs/' + n_row['姓名'] + '.pdf'
                    # self.generate_pdf(file_docx,file_pdf)
                    convert(file_docx, file_pdf, keep_active=True)
                    print('Success:',n_row['Employee Name'],n_idx+1)
                    document_1.close()

                    emp_line = [n_row['Globle ID'], n_row['Employee Name'], 'S', '']
                    df_list = df_list.append(pd.Series(emp_line, index=list_columns), ignore_index=True)

                except Exception as e:
                    print('Failed to generate file for line ', n_idx+1, e)
                    emp_line = [n_row['Globle ID'], n_row['Employee Name'], 'E', '']
                    df_list = df_list.append(pd.Series(emp_line, index=list_columns), ignore_index=True)
        df_list.to_csv(self.path + 'log.csv')


if __name__ == '__main__':
    time1 = time.time()
    mail_merge = handle_merge()
    mail_merge.main()
    print("Done, Total running time", time.time() - time1)
