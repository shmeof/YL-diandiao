from selenium import webdriver

options = webdriver.ChromeOptions()
DRIVER = webdriver.Chrome(chrome_options=options)
options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
DRIVER.get("D:\\Work\\shmeof_git\\YL-diandiao\\test_url\\305-7524038-4629161.html")
# newtime = DRIVER.find_element_by_xpath('//*[@id="headBlock"]')
newtime = DRIVER.find_element_by_xpath('//*[@id="headBlock"]')

print newtime
print newtime.text

ret = newtime.text.split("\n")
if ret[2].find("Gesendet") >= 0:
    print "ret" + ret[2][10:len(ret[2])]

DRIVER.quit()