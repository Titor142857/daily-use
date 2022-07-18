import pandas as pd
import numpy as np
import os
import openpyxl as opxl

root_dir = 'D:\\研究项目\\气质数据4月\\'
# alkane_name = '正构烷烃20220408'
chem_name = '牛肉丸气质20220408'
# alkane_dir = os.path.join(root_dir, alkane_name+'.xls')
# chem_dir = os.path.join(root_dir, chem_name+'.xls')
result_dir = os.path.join(root_dir, chem_name+'.xlsx')
extract_dir = os.path.join(root_dir, chem_name+'-extract.xlsx')
#定义操作路径，定义文件名称

xls = pd.ExcelFile(result_dir)
extract = pd.ExcelFile(extract_dir)
oc2_scale = pd.read_excel(extract)
sq_scale = pd.read_excel(extract)
book = opxl.load_workbook(extract_dir)
i = 0
for name in xls.sheet_names:
    oc2_scale.insert(i, name, 'null', allow_duplicates=False)
    sq_scale.insert(i, name, 'null', allow_duplicates=False)
    tadata = pd.read_excel(xls, sheet_name=name)
    oc2_dic = {}
    sq_dic = {}
    for index, row in tadata.iterrows():

        oc2_dic[row['物质名称']] = row['仲辛醇面积比例系数']
        sq_dic[row['物质名称']] = row['比例']

    oc2_scale[name] = oc2_scale['物质名称'].map(oc2_dic, na_action=None)
    sq_scale[name] = sq_scale['物质名称'].map(sq_dic, na_action=None)
    i = i+1

with pd.ExcelWriter(extract_dir, engine='openpyxl') as writer:
    writer.book = book
    oc2_scale.to_excel(writer, sheet_name='仲辛醇面积比例系数', index=False)
    sq_scale.to_excel(writer, sheet_name='归一化法面积比例', index=False)