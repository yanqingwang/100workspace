# -*- coding: utf-8 -*-
"""
Created on Apr 25 11:05:58 2019
@author: Z659190
"""

import re
import pandas as pd
import time
import configparser


# read file
def read_data(file_name, sheet_name,skip_line):
    skip_line = int(skip_line)
    try:
        df = pd.DataFrame(pd.read_excel(io=file_name, sheet_name=sheet_name, header=0, skiprows=skip_line))
        df = df[df['Employee Status (Label)'] == 'Active']
        return df
    except Exception as e:
        print('read file failed:', file_name)
        print('error log', e)


# read configuration
def read_config_data(file_name):
    try:
        df = pd.DataFrame(pd.read_excel(io=file_name, sheet_name='Sheet1', header=0, skiprows=0))
        df_seq = pd.DataFrame(pd.read_excel(io=file_name, sheet_name='Sequence', header=0, skiprows=0))

        return df,df_seq
    except Exception as e:
        print('read file failed:', file_name)
        print('error log', e)


def write_data(df_initial, df_detail, df_summary,p_paras,out_summary):

    out_file = p_paras.get('Path') + p_paras.get('Prefix') + time.strftime("%Y-%m-%d", time.localtime()) + out_summary

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


def get_formula(input_value, formula):
    if formula == "Remove_Zero":
        return re.sub(r"\b0*([0-z][0-z]*|0)", r"\1", str(input_value))
    elif formula == 'Upper':
        return str(input_value).upper()

    elif formula == 'Float':
        if pd.isnull(input_value):
            return 0
        else:
            return float(input_value)
    elif formula == 'Int':
        if pd.isnull(input_value):
            return 0
        else:
            return int(input_value)

    elif formula == 'Date':
        if pd.isnull(input_value):
            return 0
        else:
            return pd.to_datetime(input_value)

    elif formula == 'Count':
        if not pd.isnull(input_value):
            return 1
    else:
        return input_value


def compare_value(n_value, o_value, flag):
    num_diff = 1
    if flag == 'CNT':
        if n_value != o_value:
            return num_diff
    elif flag == 'AMT':
        return n_value - o_value
    elif flag == 'RAT':
        if not ( o_value < 0 or o_value >0 ):
            return  float(n_value/o_value)
        else:
            return 999999999


def compare_detail(paras, out_file, ):
    unique_key = paras["unique_key"]
    filename_old = paras["file_old"]
    filename_new = paras["file_new"]
    conf_file = paras["file_conf"]
    cmp_all = paras["compare_all"]
    skip_line = paras["skip_line"]

    df_old = pd.DataFrame()
    df_new = pd.DataFrame()
    df_conf = pd.DataFrame()
    df_seq = pd.DataFrame()

    file1 = paras.get('Path') + filename_old
    file2 = paras.get('Path') + filename_new
    file_config = paras.get('Path') + conf_file

    df_old = read_data(file1,'Excel Output',skip_line)
    df_new = read_data(file2,'Excel Output',skip_line)
    df_old = df_old.fillna(0)
    df_new = df_new.fillna(0)
    df_conf,df_seq = read_config_data(file_config)
    # df_conf = df_conf.fillna("")
    # df_handling = df_new.head(0)

    personlist_old = df_old[unique_key].tolist()
    personlist_new = df_new[unique_key].tolist()

    # output table
    fields_sequence = df_seq[u'OutFields'].tolist()
    df_res = pd.DataFrame(columns = fields_sequence)    # output detail results
    df_handling = pd.DataFrame()    # output original data

    print('data comparing',time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time())))
    # Handling new hires and changes, for changes, we need to compare the special fields one by one
    for n_idx,n_row in df_new.iterrows():

        if n_idx % int(paras['chk_num']) == 0:
            print('new file', n_idx, n_row[unique_key])

        if not n_row[unique_key] in personlist_old: # New hire
            df_handling = df_handling.append(pd.Series(n_row),sort=False)
            line_rs = {'New':1}

            # 根据配置表格自动比较
            for c_idx, line_conf in df_conf.iterrows():
                v_field = line_conf['NewFields']
                # fixed value
                if not pd.isnull(line_conf['OutFields']):
                    if not pd.isnull(line_conf['Formula']):
                        n_value = get_formula(n_row[v_field],line_conf['Formula'])
                    else:
                        n_value = n_row[v_field]
                    #
                    if not pd.isnull(line_conf['BoolValue']):
                        if line_conf['BoolValue'] == n_row[v_field]:
                            line_rs[line_conf['OutFields']] = 1
                    else:
                        line_rs[line_conf['OutFields']] = n_value

                #         new add
                if cmp_all == 'Y':
                    if not pd.isnull(line_conf['CompareDetail']):
                        n_value = get_formula(n_row[v_field], line_conf['Formula'])

                        diff_value = compare_value(n_value, 0, line_conf['CompareModel'])
                        if not pd.isnull(diff_value) and diff_value != 0:
                            line_rs[line_conf['CompareDetail']] = diff_value
                            if not pd.isnull(line_conf['CompareGroup']):
                                line_rs[line_conf['CompareGroup']] = line_rs[line_conf['CompareDetail']]

            df_res = df_res.append(pd.Series(line_rs),ignore_index=True)

        elif n_row[unique_key] in personlist_old:   # change
            v_change_flag = 0
            df_handling = df_handling.append(pd.Series(n_row),sort=False)
            line_rs = {}

            df_o=df_old[df_old[unique_key]==n_row[unique_key]]       #拿到符合条件的记录,dataFrame
            # print(type(df_o))
            # print(df_o)
            o_row = pd.Series(df_o.iloc[0])                   #获取行数据
            # o_row = pd.Series(df_o.iloc[0]).to_dict()                   #获取行数据

            for c_idx, line_conf in df_conf.iterrows():
                v_field = line_conf['NewFields']
                # New fixed variant
                n_value = get_formula(n_row[v_field],line_conf['Formula'])
                o_value = get_formula(o_row[line_conf['OldFields']],line_conf['Formula'])

                if not pd.isnull(line_conf['OutFields']):
                    if not pd.isnull(line_conf['BoolValue']):
                        if n_row[v_field] == line_conf['BoolValue']:
                            line_rs[line_conf['OutFields']] = 1
                    else:
                        line_rs[line_conf['OutFields']] = n_value

                # compare value
                if not pd.isnull(line_conf['CompareDetail']):
                    diff_value = compare_value(n_value,o_value,line_conf['CompareModel'])
                    if not pd.isnull(diff_value) and diff_value != 0:
                        line_rs[line_conf['CompareDetail']] = diff_value
                        v_change_flag = v_change_flag + diff_value
                        if not pd.isnull(line_conf['CompareGroup']):
                            line_rs[line_conf['CompareGroup']] = line_rs[line_conf['CompareDetail']]

            if v_change_flag > 0:
                line_rs['Change'] = 1
            else:
                line_rs['Unchange'] = 1

            df_res = df_res.append(pd.Series(line_rs),ignore_index=True)

    for o_idx, o_row in df_old.iterrows():
        if o_idx % 500 == 0:
            print('old file', o_idx, o_row[unique_key])
        if not o_row[unique_key] in personlist_new:  # Termination
            v_change_flag = 0
            df_handling = df_handling.append(pd.Series(o_row),sort=False)
            line_rs = {'Leave':1}

            for c_idx, line_conf in df_conf.iterrows():
                v_field = line_conf['OldFields']
                o_value = get_formula(o_row[v_field], line_conf['Formula'])
                if not pd.isnull(line_conf['OutputOldFields']):
                    # print(line_conf,type(line_conf))
                    if not pd.isnull(line_conf['BoolValue']):
                        if line_conf['BoolValue'] == o_row[v_field]:
                            line_rs[line_conf['OutputOldFields']] = 1
                    else:
                        line_rs[line_conf['OutputOldFields']] = o_value

                if cmp_all == 'Y':
                    if not pd.isnull(line_conf['CompareDetail']):
                        diff_value = compare_value(0, o_value , line_conf['CompareModel'])
                        if not pd.isnull(diff_value) and diff_value != 0:
                            line_rs[line_conf['CompareDetail']] = diff_value
                            if not pd.isnull(line_conf['CompareGroup']):
                                line_rs[line_conf['CompareGroup']] = line_rs[line_conf['CompareDetail']]

            df_res = df_res.append(pd.Series(line_rs),ignore_index=True)

    return df_handling, df_res,df_seq


def summary_data(p_df_detail,list_fields,rule):
    print('Summary data',time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time())))
    # p_df_detail.drop('unique_key',axis=1, inplace=True)
    print("df_detail",p_df_detail.columns.tolist())

    df_group = list_fields[list_fields['Cate'] == 'Sum']
    df_groupflds = df_group['OutFields'].tolist()

    df_sort = list_fields[list_fields['Sort'] == 'Sort']
    df_sort_flds = df_sort['OutFields'].tolist()

    sub_fields = list_fields[list_fields['Cate'] == 'Num']
    sub_fields = sub_fields.fillna(0)

    df_subfields = sub_fields['OutFields'].tolist()
    df_subfields = [fld for fld in df_subfields if str(fld) != 'nan' ]
    df_title = p_df_detail.columns.tolist()
    agg_table = {}
    for fld in df_subfields:
        if fld in df_title:
            agg_table[fld] = rule

    # df_sum = p_df_detail.groupby(['RU'],as_index=False).agg(agg_table)
    # df_sum = p_df_detail.groupby(df_groupflds,as_index=False).agg(agg_table)
    df_sum = p_df_detail.groupby(['Country','Company','RU']).agg(agg_table)
    print("df sum",df_sum.columns.tolist())
    df_sum.sort_values(df_sort_flds,inplace=True,ascending=False)

    return df_sum



class prepare_initial:
    paras = {}
    def __init__(self):

        config = configparser.ConfigParser()
        config.read('t4_auto_group.conf')
        lists_header = config.sections()  #'
        print(lists_header)

        try:
            self.paras["file_new"] = config['File']['NewFile']
            self.paras["file_old"] = config['File']['OldFile']
            self.paras["file_conf"] = config['File']['ConfFile']
            self.paras["skip_line"] = config['File']['SkipLine']
            self.paras["Path"] = config['File']['Path']
            self.paras["Prefix"] = config['File']['Prefix']
            self.paras["chk_num"] = config['File']['check_number']
            self.paras["unique_key"] = config['Parameters']['unique_key']
            self.paras["compare_all"] = config['Parameters']['compare_all']
        except Exception as e:
            print('read configuration table error',e)

    def get_paras(self):
        return self.paras

    def get_para(self,key):
        return self.para[key]


if __name__ == '__main__':

    # 获取目标文件夹的路径
    time1 = time.time()

    # file_new, file_old, skip_line, file_conf, unique_key, compare_all = init_paras()
    init_obj = prepare_initial()
    paras = init_obj.get_paras()

    out_detail = '_details.xlsx'
    out_summary = '_summary.xlsx'
    grp_agg = 'sum'

    df_initial, df_detail, df_conf_seq = compare_detail(paras, out_detail)

    df_sum = summary_data(df_detail, df_conf_seq, grp_agg)

    write_data(df_initial,df_detail,df_sum,paras,out_summary)

    time2 = time.time()

    print("Total running time", time2 - time1 )