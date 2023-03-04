#coding=utf-8
import pandas as pd

import cmodel
import os
from cmodel import *
import statsmodels.api as sm
import statsmodels.tsa.stattools as stools
import numpy as np
from sklearn.decomposition import PCA


# 准备数据文件
data_input_file = r'AgriInputs2208.xlsx'
data_output_file = r'Result_Biomass_2023_02_01_16_04_17.xlsx'

data_input_path = data_dir + '\\' + data_input_file
data_output_path = result_dir +'\\' +  data_output_file

result_file = r'Result_Conribution_' + time_tool.formatted_string() + '.xlsx'
result_path = result_dir + '\\' + result_file

info_file = r'Result_Info_' +  time_tool.formatted_string() + '.txt'
info_path = result_dir + '\\' + info_file

# 信息文件
f = open(info_path,"w")

# 导入回归数据
excel_tool.load_workbook(data_input_path)
data_input = excel_tool.cell_DataFrame()
#print(data_input)

excel_tool.load_workbook(data_output_path)
data_output = excel_tool.cell_DataFrame()
#print(data_output.shape)

# 回归变量准备
x = DataFrameClass(np.log(data_input.astype(float)))
y0 = DataFrameClass(np.log(data_output.colDataFrame(16).astype(float)))


X = x.subDataFrame(rstart=1)
y = y0.subDataFrame(rstart=1)

#print(X)
# 相关系数
var_corr = X.corr()
print('\n[变量相关系数]------------------', file=f)
print(var_corr, file=f)

# 主成分分析
model_PCA = PCA(n_components=0.95)
model_PCA.fit(X)
print('\n[主成分分析结果]------------------', file=f)
print(model_PCA.explained_variance_ratio_, file=f)
print(model_PCA.explained_variance_, file=f)
print(model_PCA.components_, file=f)

F = DataFrameClass(index=X.index)
i = 0
for c in model_PCA.components_:
    i = i + 1
    tmpF = DataFrameClass((X * c).sum(axis=1),columns=['F' + str(i)])
    F.concat(tmpF, inplace=True)

# 变量初步检验
# 平稳性检验 ADF 单位根
print('[ADF Test]-----------', file=f)
regre_types = {'nc','c','ct','ctt'}
adf_x =y.concat(F)
for i in range(0,adf_x.shape[1]):
    adf_var = adf_x.colDataFrame(i)
    print('[',i,']--',adf_var.columns[0],'--', file=f)
    for regre_type in regre_types:
        result_adf = stools.adfuller(adf_var,regression=regre_type, autolag='AIC')
        print(regre_type,result_adf, file=f)

# 添加常数项
F = DataFrameClass(sm.add_constant(F))
print(F)
X = DataFrameClass(sm.add_constant(X))


# 回归分析
model = sm.OLS(y,F)
result_ols = model.fit()
print('[回归模型结果]------------------', file=f)
print(result_ols.summary(), file=f)
print(result_ols.params, file=f)

ctb = model_PCA.components_
for i in range(np.shape(ctb)[0]):
    ctb[i,:]= ctb[i,:] * result_ols.params[i+1]

result_coeffs = DataFrameClass(ctb.sum(axis=0))
#result_coeffs = DataFrameClass(result_coeffs.T,columns = x.columns)
result_coeffs = result_coeffs.T
print('[原始变量系数]-----------------')
print(result_coeffs, file=f)

# 协整检验
print('--协整检验---')
#print(stools.coint(np.asarray(y0), np.asarray(x),trend ='ct', method ='aeg', \
#    maxlag = None, autolag ='aic', return_results = None))
print(result_ols.resid)

result_coint = stools.adfuller(result_ols.resid,regression='c',autolag='AIC')
print(result_coint)

# 计算贡献度
delta_x = DataFrameClass(data_input.diff())
delta_x = delta_x.subDataFrame(rstart=1)
x = data_input.subDataFrame(rstart=1)

delta_y = DataFrameClass(data_output.diff())
delta_y = delta_y.subDataFrame(rstart=1,cstart=16, cend=16)
y = data_output.subDataFrame(rstart=1,cstart=16, cend=16)

r1 = delta_x / x * result_coeffs
r2 = delta_y / y
result_contribution = r1 / (r2+0.0001)
print('\n[贡献度]-----------------------',file=f)
print(result_contribution, file=f)

excel_tool.close_workbook()
excel_tool.new_workbook(result_path)
excel_tool.delete_all_sheets()
excel_tool.add_sheet(r'Corr')
excel_tool.copy_DataFrame(var_corr,withtitles=True, withindex=True)
excel_tool.add_sheet(r'Contribution' )
excel_tool.copy_DataFrame(result_contribution,withtitles=True, withindex=True, indextitle='Year')
excel_tool.add_sheet(r'R1')
excel_tool.copy_DataFrame(r1,withtitles=True, withindex=True, indextitle='Year')
excel_tool.add_sheet(r'R2')
excel_tool.copy_DataFrame(r2,withtitles=True, withindex=True, indextitle='Year')

excel_tool.save_workbook()

# 关闭文件
f.close()