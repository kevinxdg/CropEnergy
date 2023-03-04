#coding=utf-8

import sys
import os

# ------------------------------------
# 确定主要工作目录
lib_dir = r'D:\Workspace\PythonLibs\ObjectLib'

sys.path.append(lib_dir)

cur_dir = os.path.dirname(os.path.abspath(__file__))

parent_dir = os.path.abspath(cur_dir + r'\..')

data_dir = parent_dir + r'\Data\Renew2023'          #采用当前文件的相对路径来确定数据目录

result_dir = r'D:\Workspace\Data\YZProject\Results'           # 保存结果的目录

# 导入公共包
from dataTools.dataObjects import *
from timeTools.timeObjects import *

# 定义公用工具
time_tool = timeTool()
excel_tool = ExcelHelper()