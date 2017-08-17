#!/usr/bin/env python
#coding=utf8

# Tkinter教程之Listbox篇
# Listbox为列表框控件，它可以包含一个或多个文本项(text item)，可以设置为单选或多选

'''''1.创建一个Listbox，向其中添加三个item'''
from Tkinter import *

root = Tk()
# lb = Listbox(root)
# for item in ['python', 'tkinter', 'widget']:
#     lb.insert(END, item)
# lb.pack()

# '''''2.创建一个可以多选的Listbox,使用属性selectmaod'''
# lb = Listbox(root, selectmode=MULTIPLE)
# for item in ['python2', 'tkinter2', 'widget2']:
#     lb.insert(END, item)
# lb.pack()

# '''''3这个属性selectmode还可以设置为BROWSE,可以通过鼠标来移动Listbox中的选中位置（不是移动item），
# 这个属性也是Listbox在默认设置的值，这个程序与1.程序运行的结果的一样的。'''
# lb = Listbox(root, selectmode=BROWSE)
# for item in ['python', 'tkinter', 'widget']:
#     lb.insert(END, item)
# lb.pack()
# # 与BROWSE相似 的为SINGLE，但不支持鼠标移动选中位置。
# lb = Listbox(root, selectmode=SINGLE)
# for item in ['python', 'tkinter', 'widget']:
#     lb.insert(END, item)
# lb.pack()

'''''4.使用selectmode  = EXPANDED使用Listbox来支持Shift和Control。'''
lb = Listbox(root, selectmode=EXTENDED)
for item in ['python', 'tkinter', 'widget']:
    lb.insert(END, item)
lb.pack()
# 运行程序，点中“python"，shift + 点击"widget"，会选中所有的item
# 运行程序，点中"python"，control + 点击"widget"，会选中python和widget，第二项tkinter处于非选中状态
root.mainloop()
