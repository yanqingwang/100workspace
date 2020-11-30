# -*- coding: utf-8 -*-
"""
@author: Z659190
提供数据预处理的公共方法

"""

from datetime import date,datetime
import pandas as pd

mg_grade = ['Management Group 1', 'Management Group 2', 'Management Group 3', 'Management Group 4',
                 'ZF TRW OIP Pool', 'Executive Mgmt Group', 'Global Executive Team', 'Board of Management']


def conv_date(x):
    # print(type(x))
    # print(x)
    if not pd.isna(x):
        # print(x.date().strftime("%Y%m"))
        # print(pd.Period(x,'M'))
        return pd.Period(x, 'M')


# 获取有效日期。将assignment的日期写入到日期字段，将公司集团日期覆盖原入职日期
def update_date(date1, date2):
    try:
        if pd.isnull(date2):
            if isinstance(date1,datetime):
                return date1
            else:
                return pd.NaT
        elif pd.isnull(date1):
            if isinstance(date2, datetime):
                return date2
            else:
                return pd.NaT
        else:
            if isinstance(date1,datetime) and isinstance(date2, datetime):
                if date2 >= date1:
                    return date2
                else:
                    return date1
            elif isinstance(date1,datetime) and (not isinstance(date2, datetime)):
                return date1
            elif (not isinstance(date1,datetime)) and isinstance(date2, datetime):
                return date2
            # else:
            #     return ''
    except Exception as e:
        print('Convert Exception:', e, date1, date2)
        try:
            return pd.to_datetime(date1)
        except Exception as e:
            return pd.to_datetime(date2)


def get_year(l_date):
    try:
        if pd.isna(l_date):
            return 'NoDate'
        if isinstance(l_date, datetime):
            l_date_day = pd.to_datetime(l_date)
            today = date.today()
            year = today.year - l_date_day.year - ((today.month, today.day) < (l_date_day.month, l_date_day.day))
            return year
        else:
            return 0
    except Exception as e:
        print('Convert date failed:', l_date,e)
        return 'NoDate'


def get_age_range(l_date):
    try:
        l_year = get_year(l_date)
        if l_year == 'NoDate':
            return 'NoDate'
        elif l_year < 25:
            return '[00--25)'
        elif l_year < 35:
            return '[25--35)'
        elif l_year < 45:
            return '[35--45)'
        elif l_year < 55:
            return '[45--55)'
        elif l_year < 70:
            return '[55--70)'
        else: # age >= 70:
            return '70+'
    except Exception as e:
        print('error log', l_date, e)
        return 'Unknown'


def get_service_year(l_date):
    try:
        l_year = get_year(l_date)
        if l_year < 1:
            return '[00--01)'
        elif l_year < 3:
            return '[01--03)'
        elif l_year < 5:
            return '[03--06)'
        elif l_year < 10:
            return '[06--10)'
        elif l_year < 15:
            return '[10--15)'
        else:
            return '[15--50)'
    except Exception as e:
        print('error log', l_date,e)
        return 'Unkown'


def get_mgr(employment_type):
    if employment_type in mg_grade:
        return "Manager_Level"
    else:
        return "Employee_level"

def get_period(l_date):
    try:
        ld = datetime(l_date)

    except Exception as e:
        print('error log', l_date,e)
        return 'Unkown'
