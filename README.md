# Automatic-RT-qPCR-analysis-by-the-delta-delta-CT-method
# A simple python script for calculating the results of multiple  RT-qPCR result sheets by 2–∆∆Ct method.

# Usage
import xlrd
import pandas as pd
import os
from Cqanalysis import cq_analysis
# Change the work directory and move the files there.
path = "E:/qPCR"  # 也可采用 r" D:\Test_path" 或者是"D:/Test_path"
xlsx_list = []
# Get the files in list
for root,dirs,files in os.walk(path):
    for file in files:
        # 使用join函数将文件名称和文件所在根目录连接起来
        if 'xlsx' in os.path.splitext(file)[1] and 'Quantification' in os.path.splitext(file)[0]:  # 筛选目录内xlsx文件
            xlsx_list.append(file)
# Read the files and use the Cqanalysis 
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
    del target[0]
    del sample[0]
    del cq[0]
    df = pd.DataFrame({'target': target,
                       'sample': sample,
                       'cq': cq})
    re = cq_analysis(df)
    holding[i] = re[0]
    i = i + 1
    print(re[0])    # cq result
    print(re[1])    # significance
    print('-----------------------------------------------------------------------')
    print('\n')
re = pd.concat(list(holding.values()), ignore_index=True)[['sample','target','group','2^(-△CT) / avg(2^(-△CT))']]
print(re)
