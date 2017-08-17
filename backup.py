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

# ================== 全局变量 ==================
global FILTER_WORKING
FILTER_WORKING = False

global DIANDIAO_WORKING
DIANDIAO_WORKING = False

global RET_TIMES
global RET_IDS
global RET_URLS

global DRIVER

# global ALL_DONE
# ALL_DONE = False
# global DONE_COUNT
# DONE_COUNT = 0
# global NOTDONE_COUNT
# NOTDONE_COUNT = 0

from selenium import webdriver

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
DRIVER = webdriver.Chrome(chrome_options=options)

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

def sync_handle_diandiao():
    global RET_TIMES
    global RET_IDS
    global RET_URLS

    import tkMessageBox as mb

    print str(len(RET_IDS))
    ok = mb.askokcancel('确认', '确定点掉选中的邮件？共：' + str(len(lb_tohandle.curselection())) + '个！')
    if ok:
        print "==================== 开始点掉 ===================="
        text.insert(END, '==================== 开始点掉 ====================\n')
        text.see(END)
        times_toclick = []
        ids_toclick = []
        urls_toclick = []
        for i in lb_tohandle.curselection():
            print "diandiao: " + str(i)
            times_toclick.append(RET_TIMES[i])
            ids_toclick.append(RET_IDS[i])
            urls_toclick.append(RET_URLS[i])

        th = threading.Thread(target=handle_diandiao, args=(times_toclick, ids_toclick, urls_toclick))
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

def handle_diandiao(times_toclick, ids_toclick, urls_toclick):
    print times_toclick
    print ids_toclick
    print urls_toclick

    times_done = []
    ids_done = []
    urls_done = []

    global DRIVER
    done = 0;
    not_done = 0;
    for i in range(len(times_toclick)):
        print "处理：" + times_toclick[i] + "  " + ids_toclick[i]
        text.insert(END, "处理：" + times_toclick[i] + "  " + ids_toclick[i] + '\n')
        text.see(END)
        DRIVER.get(urls_toclick[i])
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
                times_done.append(times_toclick[i])
                ids_done.append(ids_toclick[i])
                urls_done.append(urls_toclick[i])
                done = done + 1
                tips = "完成："
            else:
                tips = "无效："
                not_done = not_done + 1
            print tips + times_toclick[i] + "  " + ids_toclick[i]
            text.insert(END, tips + times_toclick[i] + "  " + ids_toclick[i] + '\n')
            text.see(END)

    # 写入excel
    # write_to_excel(times_done, ids_done, urls_done)
    write_to_excel(times_toclick, ids_toclick, urls_toclick)

    # global DONE_COUNT
    # global NOTDONE_COUNT
    # global ALL_DONE
    # DONE_COUNT = done
    # NOTDONE_COUNT = not_done
    # ALL_DONE = True

def refresh_info(i):
    varTipsDiandiaoSel.set('选择个数：' + str(len(lb_tohandle.curselection())))

def handle_filter():
    global FILTER_WORKING
    print "==================== 正在扫描 ===================="
    text.insert(END, '==================== 正在扫描 ====================\n')
    text.see(END)
    FILTER_WORKING = True;

    # 清空数据
    if lb_tohandle.size() > 0:
        lb_tohandle.delete(0,END)

    # url_raw = e_url_input.get()
    url_raw = "D:\\Work\\shmeof_git\\YL-diandiao\\test_url\\Bestellungen verwalten.html"
    if len(url_raw) <= 0:
        print "请填写列表页地址"
        text.insert(END, '请填写列表页地址\n')
        text.see(END)
        return

    global DRIVER
    # DRIVER = webdriver.Chrome(chrome_options=options)
    DRIVER.get(url_raw)

    print "============================================="
    ret_times = []

    elem_time = DRIVER.find_elements_by_xpath("//tbody/tr[@class='list-row-white']/td")
    print len(elem_time)
    i = 0
    for elem_time_i in elem_time:
        if (i % 6) == 1:
            # print elem_time_i.text
            ret_times.append(elem_time_i.text)
        i = i + 1

    print "============================================="
    ret_urls = []
    elem_diandiao = DRIVER.find_elements_by_xpath("//tbody/tr[@class='list-row-white']/td[@class='data-display-field-border-lb']/a")
    print len(elem_diandiao)
    i = 0
    for elem_diandiao_i in elem_diandiao:
        if (i % 2 == 1):
            url_to_handle = elem_diandiao_i.get_attribute('href')  # 获取属性值
            # print "url_href: " + url_to_handle
            ret_urls.append(url_to_handle)
        i = i + 1
    print "============================================="
    ret_ids = []
    elem_dingdan_id = DRIVER.find_elements_by_xpath(
        "//tbody/tr[@class='list-row-white']/td[@class='data-display-field-border-lbr']")
    print len(elem_dingdan_id)
    for elem_dingdan_id_i in elem_dingdan_id:
        dingdan_id = elem_dingdan_id_i.text
        # print "dingdan_id: " + dingdan_id
        ret_ids.append(dingdan_id)
    print "============================================="


    # 新方式

    print "新方式"
    counts = DRIVER.find_elements_by_xpath('//*[@id="inboxPanel "]/table/tbody/tr')
    print len(counts)
    for i in range(len(counts)):
        regex_Datum = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[2]'
        regex_Absender = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[3]/a'
        regex_Betreff_Content = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[4]/a/span'
        regex_Betreff_Url = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[4]/a'
        regex_Referenz = '//*[@id="inboxPanel "]/table/tbody/tr[' + str(i + 1) + ']/td[6]/a'
        Datum = DRIVER.find_element_by_xpath(regex_Datum)
        Absender = DRIVER.find_element_by_xpath(regex_Absender)
        Betreff_Content = DRIVER.find_element_by_xpath(regex_Betreff_Content)
        Betreff_Url = DRIVER.find_element_by_xpath(regex_Betreff_Url)
        Referenz = DRIVER.find_element_by_xpath(regex_Referenz)
        print "================="
        print Datum.text
        print Absender.text
        print Betreff_Content.text
        print Betreff_Url.text
        print Referenz.text

    for i in range(len(ret_urls)):
        # print ret_times[i]
        # print ret_ids[i]
        # print ret_urls[i]

        # text.insert(END, '====================处理订单====================\n')
        # text.insert(END, '时间: ' + ret_times[i] + '\n')
        # text.insert(END, '单号: ' + ret_ids[i] + '\n')
        # text.insert(END, '链接: ' + ret_urls[i] + '\n')

        lb_tohandle.insert(END, ret_times[i] + "    " + ret_ids[i])

    print "==================== 扫描完成 ===================="
    print str(len(ret_times)) + "   " + str(len(ret_ids))
    text.insert(END, '==================== 扫描完成 ====================\n')
    text.see(END)
    FILTER_WORKING = False
    varTipsDiandiaoAll.set('邮件总数：' + str(len(ret_times)))

    global RET_TIMES
    global RET_IDS
    global RET_URLS
    RET_TIMES = ret_times
    RET_IDS = ret_ids
    RET_URLS = ret_urls

    # driver.get(ret_urls[0])
    print "============================================="

def write_to_excel(ret_times, ret_ids, ret_urls):
    import xlrd
    import xlutils.copy

    try:
        newFile = 'D:\\Work\\shmeof_git\\YL-diandiao\\diandiao.xls'
        newwb = xlrd.open_workbook(newFile)
        outwb = xlutils.copy.copy(newwb)

        outSheet = outwb.get_sheet(0)

        setOutCell(outSheet, 2, 4, '888333')


        ISOTIMEFORMAT = '%Y-%m-%d-%X'
        timestr = time.strftime(ISOTIMEFORMAT, time.localtime())
        timestr = timestr.replace(":", "-")
        outwb.save('D:\\Work\\shmeof_git\\YL-diandiao\\diandiao_' + timestr + '.xls')
    except Exception, e:
        print str(e)
    finally:
        print "保存excel完成"

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

e_url_input = Entry(root, width = 80)
e_url_input.grid(row=0,column=1,sticky=E)

buttonFilter = Button(root, text='扫描邮件', font=('微软雅黑', 10), fg='blue', command=sync_handle_filter)
buttonFilter.grid(row = 0, column = 2)

# 扫描列表显示
scrollbar = Scrollbar(root)
lb_tohandle = Listbox(root, selectmode=EXTENDED, width = 100, yscrollcommand = scrollbar.set)
scrollbar['command']=lb_tohandle.yview
lb_tohandle.grid(row = 1, column = 0, columnspan = 3)
scrollbar.grid(row = 1, column = 3)
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

# 任务操作显示
text = ScrolledText(root, font=('微软雅黑', 10), fg='black', width = 85)
text.grid(row = 3, column = 0, columnspan = 3)

root.mainloop()
