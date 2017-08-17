#!/usr/bin/env python
#coding=utf8

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import threading
import os, time
import tkMessageBox
import time
# pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
# pdfmetrics.registerFont(TTFont('msyh', 'msyh.ttf'))

from Tkinter import *
from ScrolledText import ScrolledText  # 文本滚动条x

global DIR_WORK
DIR_WORK = 'D:\\YL-diandiao'
global DIR_EXCEL_TEMPLATE
DIR_EXCEL_TEMPLATE = DIR_WORK + "\\TEMPLATE"
# ================== 全局变量 ==================
global FILTER_WORKING
FILTER_WORKING = False

global DIANDIAO_WORKING
DIANDIAO_WORKING = False

global RET_Datum
global RET_Absender
global RET_Betreff_Content
global RET_Betreff_Url
global RET_Referenz

global DRIVER

from selenium import webdriver

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])

def sync_handle_filter():
    global FILTER_WORKING
    if FILTER_WORKING:
        print "==================== 正在扫描 ===================="
        text.insert(END, '==================== 正在扫描 ====================\n')
        text.see(END)
        import tkMessageBox as mb
        mb.showinfo("确定", "正在扫描！")
        return

    th = threading.Thread(target=handle_filter, args=())
    th.setDaemon(True)
    th.start()

def open_excel_dir():
    os.system("explorer.exe %s" % DIR_WORK)

def sync_handle_diandiao():
    global RET_Datum
    global RET_Absender
    global RET_Betreff_Content
    global RET_Betreff_Url
    global RET_Referenz

    import tkMessageBox as mb

    print str(len(RET_Referenz))
    ok = mb.askokcancel('确认', '确定点掉选中的邮件？共：' + str(len(lb_tohandle.curselection())) + '个！')
    if ok:
        print "==================== 开始点掉 ===================="
        text.insert(END, '==================== 开始点掉 ====================\n')
        text.see(END)
        Datum_toclick = []
        Absender_toclick = []
        Betreff_Content_toclick = []
        Betreff_Url_toclick = []
        Referenz_toclick = []
        for i in lb_tohandle.curselection():
            print "diandiao: " + str(i)
            Datum_toclick.append(RET_Datum[i])
            Absender_toclick.append(RET_Absender[i])
            Betreff_Content_toclick.append(RET_Betreff_Content)
            Betreff_Url_toclick.append(RET_Betreff_Url[i])
            Referenz_toclick.append(RET_Referenz[i])

        th = threading.Thread(target=handle_diandiao, args=(Datum_toclick, Absender_toclick, Betreff_Content_toclick, Betreff_Url_toclick, Referenz_toclick))
        th.setDaemon(True)
        th.start()

        # global ALL_DONE
        # ALL_DONE = False
        # while(True):
        #     if ALL_DONE:
        #         import tkMessageBox as mb
        #         mb.showinfo("完成任务", "处理邮件：" + str(len(times_toclick)) + "个\n点掉：" + str(DONE_COUNT) + "个\n" + "无效：" + str(NOTDONE_COUNT) + "个")
        #         break

    else:
        print "==================== 已经取消 ===================="
        text.insert(END, '==================== 已经取消 ====================\n')
        text.see(END)

def handle_diandiao(Datum_toclick, Absender_toclick, Betreff_Content_toclick, Betreff_Url_toclick, Referenz_toclick):
    Datum_done = []
    Absender_done = []
    Betreff_Content_done = []
    Betreff_Url_done = []
    Referenz_done = []

    global DRIVER
    DRIVER = webdriver.Chrome(chrome_options=options)
    done = 0;
    not_done = 0;
    for i in range(len(Datum_toclick)):
        print "处理：" + Datum_toclick[i] + "  " + Referenz_toclick[i]
        text.insert(END, "处理：" + Datum_toclick[i] + "  " + Referenz_toclick[i] + '\n')
        text.see(END)
        DRIVER.get(Betreff_Url_toclick[i])
        succ = False
        try:
            elem_check = DRIVER.find_element_by_xpath('//*[@id="noResponseRequired"]')
            elem_check.click()
            elem_replay_button = DRIVER.find_element_by_xpath('//*[@id="ReplyButton"]')
            elem_replay_button.click()
            succ = True
        except Exception, e:
            text.insert(END, "异常：" + str(e) + '\n')
            text.see(END)
            succ = False
        finally:
            if succ:
                Datum_done.append(Datum_toclick[i])
                Absender_done.append(Absender_toclick[i])
                Betreff_Content_done.append(Betreff_Content_toclick[i])
                Betreff_Url_done.append(Betreff_Url_toclick[i])
                Referenz_done.append(Referenz_toclick[i])
                done = done + 1
                tips = "完成："
            else:
                tips = "无效："
                not_done = not_done + 1

            print tips + Datum_toclick[i] + "    " + Referenz_toclick[i] + "    " + Absender_toclick[i]
            text.insert(END, tips + Datum_toclick[i] + "    " + Referenz_toclick[i] + "    " + Absender_toclick[i] + '\n')
            text.insert(END, '==================================================\n')
            text.see(END)

    # 退出浏览器
    DRIVER.quit()

    # 写入excel
    # write_to_excel(Datum_toclick, Absender_toclick, Betreff_Content_toclick, Betreff_Url_toclick, Referenz_toclick)
    write_to_excel(Datum_done, Absender_done, Betreff_Content_done, Betreff_Url_done, Referenz_done)

def refresh_info(i):
    varTipsDiandiaoSel.set('选择个数：' + str(len(lb_tohandle.curselection())))

def handle_filter():
    url_raw = e_url_input.get()
    # url_raw = "D:\\Work\\shmeof_git\\YL-diandiao\\test_url\\Bestellungen verwalten.html"
    # url_raw = "D:\\Work\\shmeof_git\\YL-diandiao\\test_url\\Bestellungen verwalten_new.html"
    # url_raw = DIR_WORK + "\\Bestellungen verwalten_new.html"
    if len(url_raw) <= 0:
        print "请填写列表页地址"
        text.insert(END, '请填写列表页地址\n')
        text.see(END)
        return

    global FILTER_WORKING
    print "==================== 正在扫描 ===================="
    text.insert(END, '==================== 正在扫描 ====================\n')
    text.see(END)
    FILTER_WORKING = True;

    # 清空数据
    if lb_tohandle.size() > 0:
        lb_tohandle.delete(0,END)

    ret_Datum = []
    ret_Absender = []
    ret_Betreff_Content = []
    ret_Betreff_Url = []
    ret_Referenz = []

    global DRIVER
    DRIVER = webdriver.Chrome(chrome_options=options)
    DRIVER.get(url_raw)
    counts = DRIVER.find_elements_by_xpath("//tbody/tr[@class='list-row-white']/td[@class='data-display-field-border-lbr']")
    print "新方式"
    print len(counts)
    for i in range(len(counts)):
        Datum = None
        Absender = None
        Betreff_Content = None
        Betreff_Url = None
        Referenz = None
        regex_Datum = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[2]'
        regex_Absender = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[3]/a'
        regex_Betreff_Content = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[4]/a/span'
        regex_Betreff_Url = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[4]/a'
        regex_Referenz = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[6]/a'

        try:
            Datum = DRIVER.find_element_by_xpath(regex_Datum)
        except Exception, e:
            print str(e)

        try:
            Absender = DRIVER.find_element_by_xpath(regex_Absender)
        except Exception, e:
            print str(e)

        try:
            Betreff_Content = DRIVER.find_element_by_xpath(regex_Betreff_Content)
        except Exception, e:
            print str(e)

        try:
            Betreff_Url = DRIVER.find_element_by_xpath(regex_Betreff_Url)
        except Exception, e:
            print str(e)

        try:
            Referenz = DRIVER.find_element_by_xpath(regex_Referenz)
        except Exception, e:
            print str(e)

        if Datum:
            ret_Datum.append(Datum.text)
        else:
            ret_Datum.append("")

        if Absender:
            ret_Absender.append(Absender.text)
        else:
            ret_Absender.append("")

        if Betreff_Content:
            ret_Betreff_Content.append(Betreff_Content.text)
        else:
            ret_Betreff_Content.append("")

        if Betreff_Url:
            ret_Betreff_Url.append(Betreff_Url.get_attribute('href'))
        else:
            ret_Betreff_Url.append("")

        if Referenz:
            ret_Referenz.append(Referenz.text)
            text.insert(END, '扫描到订单：' + Referenz.text + '\n')
            text.see(END)
        else:
            ret_Referenz.append("                                ")
            text.insert(END, '扫描到订单：\n')
            text.see(END)

    for i in range(len(ret_Datum)):
        if str(ret_Datum[i]).find(':') == -1:
            print "需要获取时间：" + ret_Datum[i]
            newtime = reget_time(ret_Datum[i], ret_Betreff_Url[i])
            print "时间：" + newtime
            ret_Datum[i] = newtime

    for i in range(len(ret_Datum)):
        lb_tohandle.insert(END, ret_Datum[i] + "    " + ret_Referenz[i] + "    " + ret_Absender[i] + "        " + ret_Betreff_Content[i])

    DRIVER.quit()

    print "==================== 扫描完成 ===================="
    print str(len(ret_Datum)) + "   " + str(len(ret_Referenz))
    text.insert(END, '==================== 扫描完成 ====================\n')
    text.see(END)
    FILTER_WORKING = False
    varTipsDiandiaoAll.set('邮件总数：' + str(len(ret_Datum)))

    global RET_Datum
    global RET_Absender
    global RET_Betreff_Content
    global RET_Betreff_Url
    global RET_Referenz
    RET_Datum = ret_Datum
    RET_Absender = ret_Absender
    RET_Betreff_Content = ret_Betreff_Content
    RET_Betreff_Url = ret_Betreff_Url
    RET_Referenz = ret_Referenz

    # driver.get(ret_urls[0])
    print "============================================="

def reget_time(oldtime, url):
    newtime = oldtime
    # DRIVER.get(url)
    DRIVER.get("D:\\Work\\shmeof_git\\YL-diandiao\\test_url\\305-7524038-4629161.html")
    try:
        newtime = DRIVER.find_element_by_xpath('//*[@id="headBlock"]/text()[2]')
        # print "时间多少个：" + str(len(newtime))
        print "before"
        print newtime
        print "got"
    except Exception,e:
        print str(e)

    return newtime

def write_to_excel(Datum_toclick, Absender_toclick, Betreff_Content_toclick, Betreff_Url_toclick, Referenz_toclick):
    import xlrd
    import xlutils.copy

    try:
        newFile = DIR_EXCEL_TEMPLATE + '\\diandiao.xls'
        newwb = xlrd.open_workbook(newFile)
        outwb = xlutils.copy.copy(newwb)
        outSheet = outwb.get_sheet(0)

        for i in range(len(Datum_toclick)):
            setOutCell(outSheet, 0, i, Datum_toclick[i])
        for i in range(len(Absender_toclick)):
            setOutCell(outSheet, 1, i, Absender_toclick[i])
        for i in range(len(Betreff_Content_toclick)):
            setOutCell(outSheet, 2, i, Betreff_Content_toclick[i])
        for i in range(len(Betreff_Url_toclick)):
            setOutCell(outSheet, 3, i, Betreff_Url_toclick[i])
        for i in range(len(Referenz_toclick)):
            setOutCell(outSheet, 4, i, Referenz_toclick[i])

        ISOTIMEFORMAT = '%Y%m%d-%X'
        timestr = time.strftime(ISOTIMEFORMAT, time.localtime())
        timestr = timestr.replace(":", "-")
        outwb.save(DIR_WORK + '\\diandiao_' + timestr + '.xls')
        print "==================== 保存Excel成功 ===================="
        text.insert(END, '==================== 保存Excel成功 ====================\n')
        text.see(END)
    except Exception, e:
        print str(e)
        text.insert(END, str(e) + '\n')
        text.see(END)
    finally:
        print "==================== 任务结束 ===================="
        text.insert(END, '==================== 任务结束 ====================\n')
        text.insert(END, '共点掉邮件：' + str(len(Datum_toclick)) + '个！\n')
        text.see(END)

# 修改值
def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """

    def _getOutCell(outSheet, colIndex, rowIndex):
        """ HACK: Extract the internal xlwt cell representation. """
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row: return None

        cell = row._Row__cells.get(colIndex)
        return cell

    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
            # END HACK

# ============================== 主页面 ==============================
root = Tk()
root.title('YL-diandiao')
root.geometry('+0+0') # 窗口呈现位置

varTipsLabel = StringVar()  # 设置变量
label_url = Label(root, font=('微软雅黑', 10), fg='black', textvariable=varTipsLabel)
label_url.grid(row = 0, column = 0)
varTipsLabel.set('列表页地址：')

e_url_input = Entry(root, width = 150)
e_url_input.grid(row=0,column=1,sticky=E, columnspan = 2)

buttonFilter = Button(root, text='扫描邮件', font=('微软雅黑', 10), fg='blue', command=sync_handle_filter)
buttonFilter.grid(row = 0, column = 3)

# 扫描列表显示
# scrollbar = Scrollbar(root)
# lb_tohandle = Listbox(root, selectmode=EXTENDED, width = 170, yscrollcommand = scrollbar.set)
lb_tohandle = Listbox(root, selectmode=EXTENDED, width = 180)
# scrollbar['command']=lb_tohandle.yview
lb_tohandle.grid(row = 1, column = 0, columnspan = 4)
# scrollbar.grid(row = 1, column = 3)
lb_tohandle.bind("<<ListboxSelect>>", refresh_info)

# 信息提示
varTipsDiandiaoAll = StringVar()  # 设置变量
label_diandiao_all = Label(root, font=('微软雅黑', 10), fg='black', textvariable=varTipsDiandiaoAll)
label_diandiao_all.grid(row = 2, column = 0)
varTipsDiandiaoAll.set('邮件总数：0')
varTipsDiandiaoSel = StringVar()  # 设置变量
label_diandiao_sel = Label(root, font=('微软雅黑', 10), fg='black', textvariable=varTipsDiandiaoSel)
label_diandiao_sel.grid(row = 2, column = 1)
varTipsDiandiaoSel.set('选择个数：0')
buttonDiandiao = Button(root, text='点掉邮件', font=('微软雅黑', 10), fg='blue', command=sync_handle_diandiao)
buttonDiandiao.grid(row = 2, column = 2)
buttonDakaiDir= Button(root, text='打开Excel目录', font=('微软雅黑', 10), fg='blue', command=open_excel_dir)
buttonDakaiDir.grid(row = 2, column = 3)

# 任务操作显示
text = ScrolledText(root, font=('微软雅黑', 10), fg='black', width = 155)
text.grid(row = 3, column = 0, columnspan = 4)

root.mainloop()
