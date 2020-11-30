# -*- coding: utf-8 -*-
"""
@author: Z659190
静态数据展示显示处理 --- 过去一段时间在职人数、新入职人数、离职人数
"""
import pandas as pd
# import numpy as np

low_date = '2020-01-01'
high_date = '2020-12-31'
date_range = pd.period_range(low_date, high_date, freq='M')

class AnalyzeTimeObj(object):
    def __init__(self,file_res):
        try:
            self.df_writer = pd.ExcelWriter(file_res, engine='xlsxwriter')
            self.workbook = self.df_writer.book
            self.row_num = 1
            self.d_data = pd.DataFrame
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
                'name': v_title,
                'name': [sheet_name,self.row_num, i+1],
                # 开始行，开始列，结束行，结束列
                'categories': [sheet_name, self.row_num+1, start_col, self.row_num+x, start_col],
                'values': [sheet_name, self.row_num+1, start_col+i, self.row_num+x, start_col+i],
                'data_labels': {'value': True},
            })
        # Add a title.
        chart.set_title({'name': v_title})
        worksheet.insert_chart(self.row_num, 5 , chart, {'x_offset': 10, 'y_offset': 20})
        self.row_num = self.row_num + x + 10

    def create_column_chart(self, df1, sheet_name,v_title, start_col):
        chart_type = 'line'
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
                'name': [sheet_name,self.row_num, start_col+i],
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

    def out_yearly_analyze(self,df_data,prefix):
        df1 = pd.DataFrame()
        lv_fieldname = 'Count'
        lv_row = 0
        try:
            # 创建一个excel
            # gender

            sheet_name = prefix+'00_traw'
            df_data.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            df_data = df_data.rename(columns={'ZF Global ID':lv_fieldname})
            df1 = df_data.loc[df_data['NewHP'].isin(date_range)].reset_index(drop=True)

            sheet_name = prefix+'10_trends_data'
            df11 = df1.groupby(['NewHP'])[lv_fieldname].count().reset_index()
            df11 = df11.sort_values(by=['NewHP'], ascending=True).reset_index(drop = True)
            df11 = df11.rename(columns={'NewHP':'Month'})
            df11 = df11.rename(columns={'Count':'NewHire_No'})
            df11.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8", index=False,startrow=lv_row)

            lv_row = lv_row + 20
            pivot_df1 = pd.pivot_table(df1,index=['Country'],columns=['NewHP'],
                                       values=[lv_fieldname],aggfunc='count',margins=True).reset_index()
            # print(pivot_df1.head(2))
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 20
            pivot_df1 = pd.pivot_table(df1,index=['Division'],columns=['NewHP'],
                                       values=[lv_fieldname],aggfunc='count').reset_index()
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 20
            pivot_df1 = pd.pivot_table(df1,index=['Employee Class (Label)'],columns=['NewHP'],
                                       values=[lv_fieldname],aggfunc='count').reset_index()
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 10
            pivot_df1 = pd.pivot_table(df1,index=['Employment Type (Label)'],columns=['NewHP'],
                                       values=[lv_fieldname],aggfunc='count',margins=True).reset_index()
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 20
            df2 = df_data.loc[df_data['NewTP'].isin(date_range)].reset_index(drop=True)
            df21 = df2.groupby(['NewTP'])[lv_fieldname].count().reset_index()
            df21 = df21.sort_values(by=['NewTP'], ascending=True).reset_index(drop = True)
            df21 = df21.rename(columns={'NewTP':'Month'})
            df21 = df21.rename(columns={'Count':'Termination No.'})
            df21.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 20
            pivot_df1 = pd.pivot_table(df2,index=['Country'],columns=['NewTP'],
                                       values=[lv_fieldname],aggfunc='count').reset_index()
            # print(pivot_df1.head(2))
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 20
            pivot_df1 = pd.pivot_table(df2,index=['Division'],columns=['NewTP'],
                                       values=[lv_fieldname],aggfunc='count').reset_index()
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 20
            pivot_df1 = pd.pivot_table(df2,index=['Employee Class (Label)'],columns=['NewTP'],
                                       values=[lv_fieldname],aggfunc='count').reset_index()
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            lv_row = lv_row + 10
            pivot_df1 = pd.pivot_table(df2,index=['Employment Type (Label)'],columns=['NewTP'],
                                       values=[lv_fieldname],aggfunc='count').reset_index()
            pivot_df1.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8",startrow=lv_row)

            sheet_name = prefix+'30_combine'
            merge_data = pd.merge(df11, df21, how='outer', on=['Month'], indicator=False).reset_index(drop=True)
            self.create_column_chart(merge_data,sheet_name,'New Hire / Termination Employee No.',1)
            self.d_data = merge_data

        except Exception as e:
            print('write period data file failed:', sheet_name)
            print('error log', e)
        finally:
            pass

    def out_period_res(self,df_data,prefix):

        df_res_local = pd.DataFrame()
        for month in date_range:
            line = {}
            try:
                line['Month'] = month
                # last_day = month.to_timestamp(how='end')
                # print(last_day)
                line['Active_total'] = len(df_data.loc[(df_data['NewHP'] <= month) &
                                                             ((df_data['NewTP'] > month) | (df_data['NewTP'].isna()))].index)
                                                             # (df_data['NewTP'].isna())].index)
                line['NewHire'] = len(df_data.loc[(df_data['NewHP'] == month)].index)
                line['Termination'] = len(df_data.loc[(df_data['NewTP'] == month)].index)
                df_res_local = df_res_local.append(pd.Series(line), ignore_index=True)
                df_res_local = df_res_local[['Month','Active_total','NewHire','Termination']]

            except Exception as e:
                print('Count Data Exception:', e)

        try:
            sheet_name = prefix+'30_combine'
            df_res_local['Month'] = df_res_local['Month'].apply(pd.Period)
            # self.d_data = self.d_data.infer_objects()
            frames = [df_res_local,self.d_data]

            df_res_new = pd.merge(df_res_local, self.d_data, how='outer', on=['Month'], indicator=False).reset_index(drop=True)
            # df_res_new = pd.concat(frames,keys=['Month'], axis=1).reset_index(drop=True)
            # df_res_new = pd.concat(frames, axis=1, ignore_index=True)
            # df_res_new = pd.concat(frames, axis=1)
            self.create_column_chart(df_res_new, sheet_name, 'Combined Res.', 1)
            # df_res_new.to_excel(self.df_writer, sheet_name=sheet_name, encoding="utf-8", index=False)

            df_res_local['Active_CHG'] = df_res_local.Active_total.diff()
            df_res_local['New_CHG'] = df_res_local.NewHire.diff()
            df_res_local['Term_CHG'] = df_res_local.Termination.diff()

            self.create_column_chart(df_res_local, sheet_name, 'Change Trends.', 1)

        except Exception as e:
            print('Get active period data file failed:', sheet_name)
            print('error log', e)
        finally:
            pass

    def main(self,df_data):
        df_res = df_data.loc[(df_data['Employee Status (Label)'] != 'Dormant') & (df_data['Employee Status (Label)'] != 'Discarded') ]
        self.out_yearly_analyze(df_res,'AP')
        self.out_period_res(df_res,'AP')

        self.row_num = 1
        df_res = df_res.loc[df_res['Country'] == 'CN']
        self.out_yearly_analyze(df_res,'CHN')
        self.out_period_res(df_res,'CHN')

        print("finished time period analyzing")
        try:
            self.workbook.close()
            self.df_writer.close()
        except Exception as e:
            print('Close file failed:', e )

