#!/usr/bin/env python
#coding=utf8

# python2下
from Tkinter import *

import tkSimpleDialog as dl
import tkMessageBox as mb

root = Tk()
w = Label(root, text="Label Title")
w.pack()

# mb.showinfo("welcome", "Welcome Message")
# print "here"

def cmd():
    print "cmd"

ok = mb.askokcancel('Python Tkinter', 'askokcancel')
print ok

# guess = dl.askinteger("Number", "Enter a number")
#
# output = 'This is output message'
# mb.showinfo("Output:", output)