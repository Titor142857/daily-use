import pandas as pd
import numpy as np
import os
import openpyxl as opxl

root_dir = 'D:\\研究项目\\气质数据4月\\'
alkane_name = '正构烷烃20220408'
chem_name = '牛肉丸20220408'
alkane_dir = os.path.join(root_dir, alkane_name+'.xls')
chem_dir = os.path.join(root_dir, chem_name+'.xls')
result_dir = os.path.join(root_dir, chem_name+'.xlsx')
#定义操作路径，定义文件名称

alkanexls = pd.ExcelFile(alkane_dir)
ridata = {}
alkanedata = pd.read_excel(alkanexls)
alkanedata.to_excel(result_dir, sheet_name='alkane', index=False)
for index,row in alkanedata.iterrows():
    ridata[row['RI']] = row['RT']
#创建正构烷烃RT对应RI的字典，并为后面的数据导出创建文件

def ricalcu(rt):
    for i in dict.keys(ridata):
        if max(dict.values(ridata)) <= rt:
            ri = 9999
            break
        elif ridata[i] <= rt and ridata[i+100] >= rt:
            ri = (i + 100*(rt-ridata[i])/(ridata[i+100]-ridata[i]))/60
            break
        else:
            ri = 0
    return(ri)
#定义函数，根据正构烷烃RT对应RI的字典，计算该色谱条件下任意RT值对应的RI值
#0和9999为色谱峰RT值超出了正构烷烃范围，0为RT小于最小的正构烷烃，9999为 RT大于最大的正构烷烃

book = opxl.load_workbook(result_dir)
xls = pd.ExcelFile(chem_dir)
for name in xls.sheet_names:
    tadata = pd.read_excel(xls, sheet_name=name)
    tadata.insert(2, 'RI', 0)
    tadata.insert(3, '物质名称', 'null')
    RIwr = {}
    for index,row in tadata.iterrows():
        RIwr[row['RT']] = ricalcu(row['RT'])
        #print(index)
    #print(RIwr)
    tadata['RI']=tadata['RT'].map(RIwr)
    #print(tadata)
    #遍历循环色谱峰文件，使用ricalcu函数计算对应RI

    with pd.ExcelWriter(result_dir, engine='openpyxl') as writer:
        writer.book = book
        #writer.sheets = {sheet.title: sheet for sheet in book.worksheets}
        tadata.to_excel(writer, sheet_name=name, index=False)
    #以色谱峰文件每一页为一个dataframe，输出至新的结果文件