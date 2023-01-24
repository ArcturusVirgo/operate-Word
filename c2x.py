# -- coding: utf-8 --
# @Time: 2022/9/5 11:21
# @Author: Zavijah  zavijah@qq.com
# @File: c2x.py
# @Software: PyCharm
# @Purpose:
import os
file_list = os.listdir('./2')
for i in range(len(file_list)):
    with open('a.txt', 'w') as f:
        f.write(str(i))
    os.system(r'.\venv\Scripts\python.exe .\doc2docx.py')
