import pandas as pd
import numpy as np
import os
import openpyxl as opxl

root_dir = 'D:\\研究项目\\'
# alkane_name = '正构烷烃20220408'
or_name = '总表combine.intensity'
# alkane_dir = os.path.join(root_dir, alkane_name+'.xls')
# chem_dir = os.path.join(root_dir, chem_name+'.xls')
or_dir = os.path.join(root_dir, or_name+'.xlsx')
result_dir = os.path.join(root_dir, or_name+'-result.xlsx')

xls = pd.ExcelFile(or_dir)
ordata = pd.read_excel(xls, sheet_name='combine')
resultdata = pd.DataFrame()
for index, row in ordata.iterrows():
    MS2Metabolite = row['MS2Metabolite'].split(';')
    MS2Metabolite_reform = list(set(MS2Metabolite))
    MS1hmdbName = row['MS1hmdbName'].split(';')
    MS1hmdbName_reform = list(set(MS1hmdbName))
    MS1keggName = row['MS1keggName'].split(';')
    MS1keggName_reform = list(set(MS1keggName))
    namelist = MS2Metabolite_reform+MS1hmdbName_reform+MS1keggName_reform
    rowdata = pd.DataFrame(row).T
    #print(list)
    countlist = []
    for name in namelist:
        times = namelist.count(name)
        if times == 3 and name != '-':
            countlist.append(name)
    resultlist = list(set(countlist))
    if resultlist:
        resultnames = ';'.join(resultlist)
        rowdata.insert(3, 'resultnames',resultnames, allow_duplicates=False)
        resultdata = pd.concat([resultdata, rowdata], ignore_index=True)

with pd.ExcelWriter(result_dir, engine='openpyxl') as writer:
    resultdata.to_excel(writer, index=False)

