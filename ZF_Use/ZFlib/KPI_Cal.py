def headcnt_rate(bcs_num,sf_num):
    if bcs_num != 0:
        rate = round(1 - abs(sf_num / bcs_num - 1),4) # 返回距离1的差异
        return rate
    else:
        return 0

def nm_rate(num,total):
    try:
        if total != 0:
            rate = round(num / total,4) # 返回Rate
            return rate
        else:
            return 0
    except Exception as e:
        return 0

def overall_res(hd,sm,im,bp,job,cost,age,id):
    res = hd * 0.3 + sm * 0.1 + im * 0.1 + bp * 0.1 + job * 0.1 +  cost * 0.1 + age * 0.1 + id * 0.1
    return round(res,4)