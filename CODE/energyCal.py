from cmodel import *
import numpy as np
data_path = 'H:\\Python\\Projects\\CropEnergy\\Data\\'
exl = ExcelHelper()
exl.load_workbook(data_path + 'AgriInputs2208.xlsx')
data = exl.cell_DataFrame()
exl.load_workbook(data_path + 'Parameters.xlsx')
tmp_para = exl.cell_DataFrame()
var_index = list(range(20))
var_index = var_index[:4] + var_index[6:]
para = tmp_para.T.colsDataFrame(var_index)
para_value = para.rowDataFrame(1)
(rd,cd)=data.shape
(rp,cp)=para_value.shape
para_value.resize(rd,cp,amend_value=False,columns=para_value.columns, index=data.index)
para_N_2000 = tmp_para.T.colDataFrame(4).iloc[1,0]
para_N_2019 = tmp_para.T.colDataFrame(5).iloc[1,0]
para_value.iloc[(1994-1949):(2000-1949),3] = para_N_2000
para_value.iloc[(2000-1949):,3] = para_N_2019

data = DataFrameClass(data.fillna(0))
result = data * para_value
print(result)

tm = timeTool()
fname = data_path + 'Result\\Result_Energy_' + tm.formatted_string() + '.xlsx'
exl = ExcelHelper()
exl.copy_DataFrame(result,withtitles=True, withindex=True, indextitle='Year')
exl.save_workbook(fname)

direct_energy = result.colsDataFrame([0,1,7]).sum(axis=1)
indirect_energy = result.colsDataFrame(list(range(2,7))+list(range(8,18))).sum(axis=1)
total_energy = direct_energy + indirect_energy

energy_sum = DataFrameClass()
energy_sum.concat(direct_energy, ignore_index=False,inplace=True)
energy_sum.concat(indirect_energy, ignore_index=False,inplace=True)
energy_sum.concat(total_energy, ignore_index=False,inplace=True)
energy_sum.columns = ['Direct','Indirect','Total']

exl.add_sheet('EnergyUse')
exl.active_sheet = exl.sheet_count - 1
exl.copy_DataFrame(energy_sum, indextitle='Year')
exl.save_workbook()
print(energy_sum)





