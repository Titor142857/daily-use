import pandas as pd
from scipy import stats
import os
import openpyxl as opxl

root_dir = 'E:\\研究项目\\'
data_name = '理化指标相关性.xlsx'
result_name = '理化指标相关性结果.xlsx'
data_dir = os.path.join(root_dir, data_name)
result_dir = os.path.join(root_dir, result_name)

data_frame = pd.read_excel(pd.ExcelFile(data_dir))
result_frame = pd.read_excel(pd.ExcelFile(result_dir))
book = opxl.load_workbook(result_dir)


heads = {}
for head in data_frame:
    heads[head] = []

pearson_correlations = pd.DataFrame(heads)
pearson_pvalues = pd.DataFrame(heads)
spearman_correlations = pd.DataFrame(heads)
spearman_pvalues = pd.DataFrame(heads)
kendall_correlations = pd.DataFrame(heads)
kendall_pvalues = pd.DataFrame(heads)
df = data_frame

for keys1 in df:
    correlation = {}
    pvalue = {}
    for keys2 in df:
        s,t = stats.pearsonr(df[keys1], df[keys2])
        correlation[keys2] = s
        pvalue[keys2] = t
    pearson_correlations.loc[keys1] = correlation
    pearson_pvalues.loc[keys1] = pvalue
pearson_correlations.insert(0,'指标',heads)
pearson_pvalues.insert(0,'指标',heads)

for keys1 in df:
    correlation = {}
    pvalue = {}
    for keys2 in df:
        s,t = stats.spearmanr(df[keys1], df[keys2])
        correlation[keys2] = s
        pvalue[keys2] = t
    spearman_correlations.loc[keys1] = correlation
    spearman_pvalues.loc[keys1] = pvalue
spearman_correlations.insert(0,'指标',heads)
spearman_pvalues.insert(0,'指标',heads)

for keys1 in df:
    correlation = {}
    pvalue = {}
    for keys2 in df:
        s,t = stats.kendalltau(df[keys1], df[keys2])
        correlation[keys2] = s
        pvalue[keys2] = t
    kendall_correlations.loc[keys1] = correlation
    kendall_pvalues.loc[keys1] = pvalue
kendall_correlations.insert(0,'指标',heads)
kendall_pvalues.insert(0,'指标',heads)


with pd.ExcelWriter(result_dir, engine='openpyxl') as writer:
    writer.book = book
    pearson_correlations.to_excel(writer, sheet_name='皮尔逊系数', index=False)
    pearson_pvalues.to_excel(writer, sheet_name='皮尔逊p值', index=False)
    spearman_correlations.to_excel(writer, sheet_name='斯皮尔曼系数', index=False)
    spearman_pvalues.to_excel(writer, sheet_name='斯皮尔曼p值', index=False)
    kendall_correlations.to_excel(writer, sheet_name='肯德尔系数', index=False)
    kendall_pvalues.to_excel(writer, sheet_name='肯德尔p值', index=False)
