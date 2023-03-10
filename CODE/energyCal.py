from cmodel import *
import numpy as np

data_file = r'AgriInputs2208_For_energy.xlsx'
data_path = data_dir + data_file

para_file = r'Parameters.xlsx'
para_path = data_dir + para_file

save_file = r'Result_Energy_' + time_tool.formatted_string() + '.xlsx'
save_path = result_dir + save_file

exl = ExcelHelper()
exl.load_workbook(data_path )
data = exl.cell_DataFrame()

exl.load_workbook(para_path)
tmp_para = exl.cell_DataFrame(index=0)
print(tmp_para)

var_index = list(range(20))
var_index = var_index[:4] + var_index[6:]
para = tmp_para.T.colsDataFrame(var_index)
para_value = para.rowDataFrame(1)
print('------')
print(para_value)

(rd,cd)=data.shape
(rp,cp)=para_value.shape
para_value.resize(rd,cp,amend_value=False,columns=para_value.columns, index=data.index)
para_N_2000 = tmp_para.T.colDataFrame(4).iloc[1,0]
para_N_2019 = tmp_para.T.colDataFrame(5).iloc[1,0]
para_value.iloc[(1994-1949):(2000-1949),3] = para_N_2000
para_value.iloc[(2000-1949):,3] = para_N_2019

data = DataFrameClass(data.fillna(0))
result = data * para_value
result.columns = para_value.columns
print(result)

direct_energy = result.colsDataFrame([0,1,7])
fertilizer_energy = DataFrameClass(result.colsDataFrame([3,4,5,6]).sum(axis=1),columns=['Fertilizer'])
indirect_energy = fertilizer_energy.concat(result.colsDataFrame(list(range(8,18))),axis=1)
total_energy = direct_energy.concat(indirect_energy,axis=1)

exl = ExcelHelper()
exl.copy_DataFrame(total_energy,withtitles=True, withindex=True, indextitle='Year')
exl.save_workbook(save_path)


energy_sum = DataFrameClass()
energy_sum.concat(direct_energy.sum(axis=1), ignore_index=False,inplace=True)
energy_sum.concat(indirect_energy.sum(axis=1), ignore_index=False,inplace=True)
energy_sum.concat(total_energy.sum(axis=1), ignore_index=False,inplace=True)
energy_sum.columns = ['Direct','Indirect','Total']

exl.add_sheet('EnergyUse')
exl.active_sheet = exl.sheet_count - 1
exl.copy_DataFrame(energy_sum, indextitle='Year')
exl.save_workbook()
print(energy_sum)





