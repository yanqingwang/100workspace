# # -*- coding: utf-8 -*-
# """
# @author: Z659190
# split data by company
#
# """
import pandas as pd
import datetime
from datetime import date
import time

def get_service_year(l_date, l_end):
    try:
        if pd.isna(l_date):
            return 'NoDate'
        if pd.isna(l_end):
            return 'NoDate'
        if isinstance ( (l_date, datetime) & (l_end, datetime)):
            l_date_day = pd.to_datetime(l_date)
            l_end_day = pd.to_datetime(l_end)
            # year = l_end.year - l_date_day.year - ((l_end.month, l_end.day) < (l_date_day.month, l_date_day.day))
            lv_months = (l_end_day.year - l_date_day.year) * 12 + \
                         l_end_day.month - l_date_day.month - (l_end_day.day < l_date_day.day)

            return lv_months
        else:
            return 0
    except Exception as e:
        print('Convert date failed:', l_date,e)
        return 'NoDate'


if __name__ == '__main__':
    l_date = datetime(time.strptime('2020-01-10', "%Y-%m-%d")))
    l_end = int(time.mktime(time.strptime('2020-09-23', "%Y-%m-%d")))
    print(get_service_year(l_date,l_end))