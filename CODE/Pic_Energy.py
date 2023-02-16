#coding=utf-8
import matplotlib.pyplot as plt
from cmodel import *

data_file = r'Result_Energy_2023_02_16_10_52_53.xlsx'
data_path = result_dir + data_file

print(data_path)

excel_tool.load_workbook(data_path)
data = excel_tool.cell_DataFrame(index=0)
print(data)

pic_data = data.colDataFrame(0).T
plt.bar(data.index.tolist(),pic_data.list)
for i in range(1,data.col_count):
    btm_pic_data = data.colsDataFrame(list(range(0,i))).sum(axis=1).T.array
    pic_data = data.colDataFrame(i).T
    plt.bar(data.index.tolist(),pic_data.list,bottom=btm_pic_data)

plt.show()