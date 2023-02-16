#coding=utf-8
import matplotlib.pyplot as plt
from cmodel import *

data_file = r'Result_Energy_2023_02_16_10_52_53.xlsx'
data_path = result_dir + data_file

print(data_path)

excel_tool.load_workbook(data_path)
data = excel_tool.cell_DataFrame(index=0)
print(data)

pic_data = data.colDataFrame(1).T

print(pic_data.array)
plt.bar(data.index.tolist(),pic_data.list)
print(data.index.tolist())
plt.show()