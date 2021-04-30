# -*- coding: utf-8 -*-
from mailmerge import MailMerge
from docx2pdf import convert
import pandas as pd
import time
import os


class HandleMerge():
    def __init__(self):
        self.df_data = pd.DataFrame()
        self.path = 'C:/temp/b/'
        self.prefix = 'Bonus Communication Letter_'
        self.template = 'Bonus Communication Letter TemplateV2.docx'
        # self.template_fix = '劳动关系转移三方协议_20201221 transfer JT to CHJ2 - fix.docx'
        # self.datafile = 'Moving name list_JT&CHJ_ 202102.xlsx'
        self.datafile = 'Employee_List Div.xlsx'

    def read_files(self):
        # df_data = pd.DataFrame()
        try:
            self.df_data = pd.DataFrame(pd.read_excel(io=self.path + self.datafile,
                                                      sheet_name='Sheet1',
                                                      dtype={
                                                            'Global ID': 'str',
                                                            # 'Performance Rating': 'str',
                                                            # 'Individual Modifier': 'float',
                                                            # 'Business Modifier': 'float',
                                                            # 'Payout Rate': 'float',
                                                            # 'Bonus_Amount': 'float'
                                                             },
                                                      header=0,
                                                      # skiprows=2))
                                                      ))
            self.df_data.drop_duplicates(inplace=True)
            self.df_data.fillna("",inplace=True)
            print(self.df_data.head())
            # print(self.df_data.info())
        except Exception as e:
            print('Handle Employee Data Exception:', self.datafile, e)

    def main(self):
        list_columns = ['Global ID','Local Name', 'Status','Notes']
        df_list = pd.DataFrame(columns=list_columns)
        self.read_files()

        for n_idx, n_row in self.df_data.iterrows():

            lv_template = self.template
            document_1 = MailMerge(self.path + lv_template)

            try:

                dict = {
                    "Global_ID": str(n_row['Global ID']),
                    "English_Name": str(n_row['English Name']),
                    "Local_Name": str(n_row['Local Name']),
                    "Department": str(n_row['Department']),

                    "Performance_Rating": n_row['Performance Rating'],
                    "Bonus_Amount": '{:.2f}'.format(n_row['Bonus_Amount']),
                    "Individual_Modifier": '{:.2f}'.format(n_row['Individual Modifier']),
                    # "Individual_Modifier": '{:.2f}%'.format(n_row['Individual Modifier']),
                    "Business_Modifier": '{:.2f}'.format(n_row['Business Modifier']),
                    "Payout_Rate": '{:.2f}'.format(n_row['Payout Rate']),

                }
                document_1.merge(**dict)

                file_path = self.path + 'xlsx/'
                if not os.path.exists(file_path):
                    os.makedirs(file_path)
                # filename
                lv_name = self.prefix + str(n_row['Global ID']) + '_' + n_row['Local Name']
                file_docx = self.path + 'xlsx/' + lv_name + '.docx'
                document_1.write(file_docx)
                # time.sleep(1)

                file_path = self.path + 'pdf' + '/' + n_row['HRBP'] + '/'
                if not os.path.exists(file_path):
                    os.makedirs(file_path)

                if n_row['Division'] != "":
                    file_path = file_path + n_row['Division'] + '/'
                    if not os.path.exists(file_path):
                        os.makedirs(file_path)

                if n_row['Division2'] != "":
                    file_path = file_path + n_row['Division2'] + '/'
                    if not os.path.exists(file_path):
                        os.makedirs(file_path)

                if n_row['Division3'] != "":
                    file_path = file_path + n_row['Division3'] + '/'
                    if not os.path.exists(file_path):
                        os.makedirs(file_path)

                file_pdf = file_path + lv_name + '.pdf'

                convert(file_docx, file_pdf, keep_active=True)
                print('Success:',n_row['Local Name'],n_idx+1)
                document_1.close()

                emp_line = [n_row['Global ID'], n_row['Local Name'], 'S', '']
                df_list = df_list.append(pd.Series(emp_line, index=list_columns), ignore_index=True)

            except Exception as e:
                print('Failed to generate file for line ', n_idx+1, e)
                emp_line = [n_row['Global ID'], n_row['Local Name'], 'E', '']
                df_list = df_list.append(pd.Series(emp_line, index=list_columns), ignore_index=True)
        df_list.to_csv(self.path + 'log.csv')


if __name__ == '__main__':
    time1 = time.time()
    mail_merge = HandleMerge()
    mail_merge.main()
    print("Done, Total running time", time.time() - time1)


