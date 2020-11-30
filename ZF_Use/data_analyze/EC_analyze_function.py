# -*- coding: utf-8 -*-
"""
@author: Z659190
静态数据展示显示处理 --- 男女比例/级别分布/JF分布/国家分布/Division分布/工作地分布/直接间接情况
"""
import pandas as pd


class AnalyzeObj(object):
    def __init__(self,file_res):
        try:
            self.df_writer = pd.ExcelWriter(file_res, engine='xlsxwriter')
            self.workbook = self.df_writer.book
            self.row_num = 1
        except Exception as e:
            print('Create file failed:', file_res)
            print('error log', e)

    def create_pi_chart(self, df1, sheet_name,v_title, start_col):
        chart_type = 'pie'
        df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8", index=False, startrow=self.row_num, startcol=start_col)
        x, y = df1.shape
        # print('df_datagroup', sheet_name, x, y)
        worksheet = self.df_writer.sheets[sheet_name]
        chart = self.workbook.add_chart({'type': chart_type})  # ''subtype': 'percent_stacked'
        # for i in range(start_col, start_col+y):
        for i in range(1, y):
            # print("again",i)
            chart.add_series({
                # 'name': [sheet_name, start_col, start_row],
                'name': v_title,
                # 开始行，开始列，结束行，结束列
                'categories': [sheet_name, self.row_num+1, start_col, self.row_num+x, start_col],
                'values': [sheet_name, self.row_num+1, start_col+i, self.row_num+x, start_col+i],
                'data_labels': {'value': True, 'percentage': True},
            })
        # Add a title.
        chart.set_title({'name': v_title})
        worksheet.insert_chart(self.row_num, 5 , chart, {'x_offset': 10, 'y_offset': 20})
        self.row_num = self.row_num + x + 10

    def create_column_chart(self, df1, sheet_name,v_title, start_col):
        chart_type = 'column'
        df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8", index=False,
                     startrow=self.row_num, startcol=start_col)
        x, y = df1.shape
        # print('df_datagroup', sheet_name, x, y)
        worksheet = self.df_writer.sheets[sheet_name]
        chart = self.workbook.add_chart({'type': chart_type})  # ''subtype': 'percent_stacked'
        # for i in range(start_col, start_col+y):
        for i in range(1, y):
            # print("again",i)
            chart.add_series({
                'name': [sheet_name,self.row_num, i+1],
                # 'name': v_title,
                # 开始行，开始列，结束行，结束列
                'categories': [sheet_name, self.row_num+1, start_col, self.row_num+x, start_col],
                'values': [sheet_name, self.row_num+1, start_col+i, self.row_num+x, start_col+i],
                'data_labels': {'value': False},
            })
        # Add a title.
        chart.set_title({'name': v_title})
        worksheet.insert_chart(self.row_num, 5 , chart, {'x_offset': 10, 'y_offset': 20})
        self.row_num = self.row_num + x + 10

    def out_static_analyze(self,df_data):
        df1 = pd.DataFrame()
        lv_fieldname = 'Count'
        try:
            # 创建一个excel
            # gender
            sheet_name = '00_raw'
            df_data.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            df_data = df_data.rename(columns={'ZF Global ID':lv_fieldname})

            # gender
            sheet_name = 'StaticRes'
            df1 = df_data.groupby(['Gender'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=[lv_fieldname], ascending=False).reset_index(drop = True)
            self.create_pi_chart(df1,sheet_name,'Gender',1)
            # Division
            df1 = df_data.groupby(['Division'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=lv_fieldname, ascending=False).reset_index(drop = True)
            self.create_pi_chart(df1,sheet_name,'Division',1)
            #Country
            df1 = df_data.groupby(['Country'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=lv_fieldname, ascending=False).reset_index(drop = True)
            self.create_pi_chart(df1,sheet_name,'Country',1)
            #Grade Level
            df1 = df_data.groupby(['EmploymentType'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=lv_fieldname, ascending=False).reset_index(drop = True)
            self.create_pi_chart(df1,sheet_name,'Employment Type',1)
            #Job Family Level
            df1 = df_data[df_data['JF'] != 'N.'].groupby(['JF'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=lv_fieldname, ascending=False).reset_index(drop = True)
            self.create_pi_chart(df1,sheet_name,'Job Family',1)
            #Location Group
            df1 = df_data.groupby(['Location Group (Name)'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=lv_fieldname, ascending=False).reset_index(drop = True)
            self.create_pi_chart(df1,sheet_name,'Location Group (Name)',1)
            #Employee Class
            df1 = df_data.groupby(['Employee Class (Label)'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=lv_fieldname, ascending=False).reset_index(drop = True)
            self.create_pi_chart(df1,sheet_name,'Employee Class (Label)',1)

        #     Year Date, split China and AP
            self.row_num = 1
            sheet_name = 'DateRes'
            # Service Year
            df1 = df_data[df_data['ServiceYear']!='Unkown'].groupby(['ServiceYear'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=['ServiceYear'], ascending=True).reset_index(drop = True)
            # self.create_column_chart(df1,sheet_name,'ServiceYear',1)
            df1 = df1.rename(columns={'Count':'AllCount'})

            df2 = df_data[(df_data['ServiceYear']!='Unkown')&(df_data['Country'] == 'CN')].groupby(['ServiceYear'])[lv_fieldname].count().reset_index()
            df2 = df2.sort_values(by=['ServiceYear'], ascending=True).reset_index(drop = True)
            df2 = df2.rename(columns={'Count':'CHNCount'})
            merge_data = pd.merge(df1, df2, how='left', on='ServiceYear', left_index=True,indicator=False)
            self.create_column_chart(merge_data,sheet_name,' ServiceYear Data',1)

            # Age
            df1 = df_data[(df_data['Age']!='Unkown')].groupby(['Age'])[lv_fieldname].count().reset_index()
            df1 = df1.sort_values(by=['Age'], ascending=True).reset_index(drop = True)
            # self.create_column_chart(df1,sheet_name,'China Age',1)
            df1 = df1.rename(columns={'Count':'AllCount'})

            df2 = df_data[(df_data['Age']!='Unkown')&(df_data['Country'] == 'CN')].groupby(['Age'])[lv_fieldname].count().reset_index()
            df2 = df2.sort_values(by=['Age'], ascending=True).reset_index(drop = True)
            df2 = df2.rename(columns={'Count':'CHNCount'})
            merge_data = pd.merge(df1, df2, how='left', on='Age', left_index=True,indicator=False)
            self.create_column_chart(merge_data,sheet_name,'Age Data',1)

        except Exception as e:
            print('write file failed:', sheet_name)
            print('error log', e)
        finally:
            pass

    def main(self,df_data):
        df_active = df_data.loc[(df_data['Employee Status (Label)'] == 'Active') | (df_data['Employee Status (Label)'] == 'Unpaid Leave') ]
        self.out_static_analyze(df_active)
        print("finished static analyzing")
        try:
            self.workbook.close()
            self.df_writer.close()
        except Exception as e:
            print('Close file failed:', e )

