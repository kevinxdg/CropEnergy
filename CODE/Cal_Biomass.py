#coding=utf-8
from cmodel import *
import pandas as pd
#pd.options.display.float_format = '{:.1f}'.format

# 设定 计算 Biomass 的数据文件
data_file = r'AgriOutput_20230129.xlsx'
para_file = r'Parameters.xlsx'
result_file = 'Result_Biomass_' + time_tool.formatted_string() + '.xlsx'

data_path = data_dir + data_file
para_path = data_dir + para_file
result_path = result_dir + result_file

data_sheet_index = 0
para_sheet_index = 1


# 导入文件数据
exl = ExcelHelper()
exl.load_workbook(data_path)
data_agriOuput = exl.cell_DataFrame(index=data_sheet_index)
print(data_agriOuput.columns)

# 导入参数文件
exl.load_workbook(para_path)
para_agriOutput = exl.cell_DataFrame(index=para_sheet_index)
print(para_agriOutput)

# 计算 Biomass
biomass = data_agriOuput.copy_index()

for index in para_agriOutput.index:
    para_HI = para_agriOutput.loc[index,'HI']
    para_WI = para_agriOutput.loc[index,'WI']
    tmp_biomass = data_agriOuput[index] * (1-para_WI) / para_HI
    biomass.concat(tmp_biomass, inplace=True)
-
biomass.insert(biomass.col_count, 'TotalBiomass', biomass.sum(axis=1))
print(biomass)

# 保存数据结果

exl = ExcelHelper()
exl.copy_DataFrame(biomass,withtitles=True, withindex=True, indextitle='Year')
exl.save_workbook(result_path)

print('Done: Biomass result was saved in file ' + result_file)

