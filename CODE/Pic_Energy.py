#coding=utf-8
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from cmodel import *

data_file = r'Result_Energy_2023_02_16_23_33_23.xlsx'
data_path = result_dir + data_file

pic_file = r'Img_Energy.jpeg'
pic_path = result_dir + pic_file
print(data_path)

excel_tool.load_workbook(data_path)
data = excel_tool.cell_DataFrame(index=0)
print(data)



fig = plt.figure(figsize=(10,5),dpi=300)

plt.ylabel(r'Energy Consumption')
pic_data = data.colDataFrame(0).T
plt.bar(data.index.tolist(),pic_data.list,label=data.columns[0])
for i in range(1,data.col_count):
    btm_pic_data = data.colsDataFrame(list(range(0,i))).sum(axis=1).T.array
    pic_data = data.colDataFrame(i).T
    plt.bar(data.index.tolist(),pic_data.list,bottom=btm_pic_data, \
            label=data.columns[i])

ax = plt.gca()

ax.ticklabel_format(style='sci', scilimits=(-1, 1), axis='y',useMathText=True)

plt.legend(loc='best',ncol=2)
plt.savefig(pic_path)
plt.show()