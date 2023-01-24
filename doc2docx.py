# -- coding: utf-8 --
# @Time: 2022/3/14 13:31
# @Author: Zavijah  zavijah@qq.com
# @File: doc2docx.py
# @Software: PyCharm
# @Purpose:

import os
from win32com import client as wc

file_list = os.listdir('./2')
with open('a.txt', 'r') as f:
    n = int(f.read())


name = file_list[n]
s = name + 'x'
word = wc.Dispatch("Word.Application")
doc = word.Documents.Open(r'F:\Project\Python\PYCharm\操作办公文档\2\{}'.format(name))
doc.SaveAs(r'F:\Project\Python\PYCharm\操作办公文档\2-docx\{}'.format(s), 12)
doc.Close()
word.Quit()
