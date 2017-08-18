# 技术点：
	Python
	Selenium
	Chrome
	Chromedriver

# 技术功能点：
	在浏览器中模拟点击
	在浏览器中模拟输入
	在浏览器中抓包分析

# 浏览器模拟点击的几种方式：
	selenium + phantomjs
	selenium + firefox
	selenium + chrome
	
## 业务功能点：
	1、自动获取列表页的待处理项
	2、自动点击勾选“无需回复”，自动点击回复
	3、已回复的信息，自动保存到Excel文件

## 使用方法：
	1、下载对应Chrome浏览器版本的chromedriver，参考：
		http://blog.csdn.net/qijingpei/article/details/68925392
		http://chromedriver.storage.googleapis.com/index.html
	2、放到Chrome的安装目录，如：C:\Program Files (x86)\Google\Chrome\Application
	3、将bin目录下的文件夹"YL-diandiao"拷贝到D盘根目录，双击运行即可
	
## 打包成exe：

## 推荐：
	元素定位方法A：
		1、用Chrome打开网页
		2、Chrome右上角更多功能入口->更多工具->开发者工具->
		3、Elements子项，找到对应的元素，鼠标右键->Copy->即可获取到各种形式的元素定位方式
	
	元素定位方法B：
		Selenium IDE

## 参考：
	python 安装selenium环境 实现模拟点击模拟输入
	http://www.cnblogs.com/ldy-miss/p/6689767.html
	
	chromedriver.exe
	http://chromedriver.storage.googleapis.com/2.7/chromedriver_win32.zip

	使用python通过selenium模拟打开chrome窗口报错 出现 "您使用的是不受支持的命令行标记:--ignore-certificate-errors
	http://www.cnblogs.com/jzss/p/5567253.html

	各版本ChromeDriver下载
	http://blog.csdn.net/qijingpei/article/details/68925392
	http://chromedriver.storage.googleapis.com/index.html

	selenium之操作ChromeDriver
	https://www.testwo.com/blog/6931

	Selenium2+python自动化7-xpath定位
	http://www.cnblogs.com/yoyoketang/p/6123938.html


	Python selenium —— 父子、兄弟、相邻节点定位方式详解
	http://blog.csdn.net/huilan_same/article/details/52541680

	Tkinter教程之Listbox篇
	http://blog.csdn.net/aa1049372051/article/details/51878578

	解决Tkinter中grid/pack布局中的listbox，scrollbar组合横置
	http://blog.csdn.net/mrlevo520/article/details/51854084

	listbox 点击事件
	http://blog.csdn.net/sofeien/article/details/49464473
	
