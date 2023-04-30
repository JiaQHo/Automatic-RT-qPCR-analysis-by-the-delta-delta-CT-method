import xlrd
import pandas as pd
import os
from Cqanalysis import cq_analysis

path = "E:/qPCR"  # 也可采用 r" D:\Test_path" 或者是"D:/Test_path"
xlsx_list = []
for root,dirs,files in os.walk(path):
    for file in files:
        # 使用join函数将文件名称和文件所在根目录连接起来
        # print(os.path.join(root, file))
        if 'xlsx' in os.path.splitext(file)[1] and 'Quantification' in os.path.splitext(file)[0]:  # 筛选目录内xlsx文件
            xlsx_list.append(file)
            # wb = xlrd.open_workbook(file)     # 一个循环可以
# print(xlsx_list[0])
# wb = xlrd.open_workbook(xlsx_list[0])
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
# print(df)
# print(xlsx_list)
holding = {}
i = 1
for excel in xlsx_list:
    print(excel)
    wb = xlrd.open_workbook(excel)

    # 通过索引的方式获取到某一个sheet，现在是获取的第一个sheet页
    sheet = wb.sheet_by_index(0)

    target = sheet.col_values(3)
    sample = sheet.col_values(5)
    cq = sheet.col_values(7)
    # print(cq)
    del target[0]
    del sample[0]
    del cq[0]

    df = pd.DataFrame({'target': target,
                       'sample': sample,
                       'cq': cq})

    # df = df.sort_values(by=['target', 'sample'], ascending=True)
    # actin = df[df['target'] == 'Actin']
    # df = df[df['target'] != 'Actin']
    re = cq_analysis(df)
    holding[i] = re[0]
    i = i + 1
    print(re[0])    # cq result
    # re[0].to_csv('E:/qPCR/Cq_analysis_%s.csv' % excel, index=False, sep=',', encoding='utf_8_sig')
    # cqs[excel] = re[0]
    print(re[1])    # significance
    # re[1].to_csv('E:/qPCR/Significance_%s.csv' % excel, index=True, sep=',', encoding='utf_8_sig')
    # sign[excel] = re[1]
    print('-----------------------------------------------------------------------')
    print('\n')
re = pd.concat(list(holding.values()), ignore_index=True)[['sample','target','group','2^(-△CT) / avg(2^(-△CT))']]
# re = re.pivot(index='sample',columns='target',values='2^(-△CT) / avg(2^(-△CT))')
print(re)
# re.to_csv('E:/qPCR/合并Cq.csv', index=False, sep=',', encoding='utf_8_sig')

# re_k = re.loc[re['sample'].isin(['K1','K2','K3','K4','K5','K6'])]
# re_h = re.loc[re['sample'].isin(['H1','H2','H3','H4','H5','H6'])]
# re_j = re.loc[re['sample'].isin(['J1','J2','J3','J4','J5','J6'])]
#
# re_k = re_k.groupby(['target', 'sample', 'group']).mean()
# re_h = re_h.groupby(['target', 'sample', 'group']).mean()
# re_h = re_h.groupby(['target', 'sample', 'group']).mean()

# re_k.to_csv('E:/qPCR/空肠Cq.csv', index=False, sep=',', encoding='utf_8_sig')
# re_h.to_csv('E:/qPCR/回肠Cq.csv', index=False, sep=',', encoding='utf_8_sig')
# re_j.to_csv('E:/qPCR/结肠Cq.csv', index=False, sep=',', encoding='utf_8_sig')
# print(re_k)

