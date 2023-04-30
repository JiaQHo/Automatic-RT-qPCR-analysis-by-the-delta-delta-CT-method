import xlrd
import pandas as pd
import numpy as np
from scipy import stats


def cq_analysis(df):
    df = df.sort_values(by=['target', 'sample'], ascending=True)
    # 内参选择
    actin = df[df['target'] == 'Actin']
    df = df[df['target'] != 'Actin']

    avg = df.groupby(['target', 'sample'])['cq'].mean()  # ts = df.groupby(['target','sample']).agg(np.mean)
    # target number
    target_num = df.groupby('target').count()

    # average of targets
    target_avg = []
    i = 1
    for a in list(avg):
        while i <= 3:
            target_avg.append(a)
            i = i + 1
        i = 1
    target_array = np.array(target_avg)
    df['AVG_target'] = target_array

    # Actin
    actin_cq = actin['cq'].tolist()
    actin_cq_array = []

    i = 1
    while i <= len(target_num):
        actin_cq_array.extend(actin_cq)
        i = i + 1
    actin_cq_array = np.array(actin_cq_array)
    df['Actin'] = actin_cq_array

    # △Cq
    difference_cq = actin_cq_array - target_array
    df['△Cq'] = difference_cq

    # 2^-(△Cq)
    difference_pow = []
    for a in difference_cq:
        difference_pow.append(pow(2, -a))
    df['2^-(△Cq)'] = difference_pow

    # avg(2^(-△CT))
    avg_pow = df.groupby(['target', 'sample'])['2^-(△Cq)'].mean()

    # 选第一项作为对照组
    yiwei = []
    for a in list(avg_pow):  # 转换为一维数组
        yiwei.append(a)

    reference = []
    i = 0
    for a in yiwei:
        if i % 6 == 0:
            reference.append(yiwei[i])
        i = i + 1

    avg_pows = []
    i = 1
    for a in reference:
        while i <= 18:
            avg_pows.append(a)
            i = i + 1
        i = 1
    avg_pows = np.array(avg_pows)
    divide_pow = difference_pow / avg_pows
    df['2^(-△CT) / avg(2^(-△CT))'] = divide_pow

    # expression
    expression = df.groupby(['target', 'sample'])['2^(-△CT) / avg(2^(-△CT))'].mean()
    expression_list = []
    i = 1
    for a in list(expression):
        while i <= 3:
            expression_list.append(a)
            i = i + 1
        i = 1
    df['Rel Exp'] = expression_list

    # stdev
    stdev = np.array(list(df.groupby(['target', 'sample'])['2^(-△CT) / avg(2^(-△CT))'].std()))  # 默认为stdev
    i = 1
    stdev_list = []
    for a in stdev:
        while i <= 3:  # 三个重复样
            stdev_list.append(a)
            i = i + 1
        i = 1
    df['stdev'] = stdev_list

    # T test
    # 分组
    df.loc[df['sample'].isin(['J1', 'J4', 'J5', 'H1', 'H4', 'H5', 'K1', 'K4', 'K5']), 'group'] = 'BT'
    df.loc[df['sample'].isin(['J2', 'J3', 'J6', 'H2', 'H3', 'H6', 'K2', 'K3', 'K6']), 'group'] = 'TC'
    target_name = df['target'].tolist()
    t_test = {}
    for a in set(target_name):
        tc = df[(df['group'] == 'TC') & (df['target'] == a)]['2^(-△CT) / avg(2^(-△CT))'].tolist()
        bt = df[(df['group'] == 'BT') & (df['target'] == a)]['2^(-△CT) / avg(2^(-△CT))'].tolist()
        re = stats.ttest_ind(tc, bt)
        t_test[a] = re
    re = pd.DataFrame(t_test).T
    re.columns = ['statistic', 'pvalue']
    re.loc[re['pvalue'] < 0.05, 'sig'] = '*'
    re.loc[re['pvalue'] < 0.01, 'sig'] = '**'
    re.loc[re['pvalue'] < 0.001, 'sig'] = '***'
    return df, re


# wb = xlrd.open_workbook(r"E:/qPCR/H reg3g zo1 occludin -  Quantification Cq Results.xlsx")
# # 通过索引的方式获取到某一个sheet，现在是获取的第一个sheet页
# sheet = wb.sheet_by_index(0)
# target = sheet.col_values(3)
# sample = sheet.col_values(5)
# cq = sheet.col_values(7)
# # print(cq)
# del target[0]
# del sample[0]
# del cq[0]
# df = pd.DataFrame({'target': target,
#                    'sample': sample,
#                    'cq': cq})
#
# print(cq_analysis(df))