#!/usr/bin/env python
#coding=utf8

from selenium import webdriver
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"])
driver = webdriver.Chrome(chrome_options=options)

# ==================== 获取列表 =======================
url = "D:\\Work\\shmeof_git\\YL-diandiao\\test_url\\Bestellungen verwalten.html"
driver.get(url)
# elem_dh = driver.find_elements_by_xpath("//tr[@class='list-row-white']/td/a")
elem_biaohang = driver.find_elements_by_xpath("//tbody/tr[@class='list-row-white']")

print len(elem_biaohang)
count = 0;
for elem_biaohang_i in elem_biaohang:
    if count > 2:
        break

    count = count + 1
    print "=================================================="
    print elem_biaohang_i.text  # 获取正文
    # elem_to_diandiao_time = elem_biaohang_i.find_elements_by_xpath("//td")
    # for elem_to_diandiao_time_i in elem_to_diandiao_time:
    #     print "time: " + elem_to_diandiao_time_i.text

    # elem_to_diandiao_info = elem_biaohang_i.find_elements_by_xpath("//td/a")
    elem_to_diandiao_info = elem_biaohang_i.find_elements_by_xpath("//td[@class='data-display-field-border-lb']/a")
    for elem_to_diandiao_i in elem_to_diandiao_info:
        print elem_to_diandiao_i.text
        url_to_handle = elem_to_diandiao_i.get_attribute('href')  # 获取属性值
        if url_to_handle:
            print "url_href: " + url_to_handle

# ==================== 点掉网页 =======================
# url = "D:\\Work\\shmeof_git\\YL-diandiao\\test_url\\305-7524038-4629161.html"
# driver.get(url)
# elem_check = driver.find_element_by_id("noResponseRequired")
# elem_check.click()
# elem_replay_button = driver.find_element_by_id("ReplyButton")
# elem_replay_button.click()

# 相关信息加入文件


# driver.quit()

