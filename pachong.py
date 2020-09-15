'''
Author: asterisk
Date: 2020-08-30 17:32:36
LastEditTime: 2020-09-15 09:47:26
LastEditors: Please set LastEditors
Description: In User Settings Edit
FilePath: /zhaobiao_pachong/pachong.py
'''
# coding:utf-8
import sys
import io
import time
import json
import requests
import xlwt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from lxml import etree

driver = webdriver.Chrome()

# Read cookies (Get cookies first)
with open('cookies.txt', 'r', encoding='utf8') as f:
    listCookies = json.loads(f.read())

# Open website and login
driver.get('https://www.okcis.cn/search/')  # 打开网页
for cookie in listCookies:
    driver.add_cookie(cookie)
driver.refresh()
# time.sleep(2)
# 利用xpath输入文本，模拟点击
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[1]/div[1]/div[1]/input[2]').send_keys('环卫 保洁')
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[4]/div[1]/input').send_keys('车 厕')
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[5]/div[1]/label[7]/input').click()
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[6]/div[1]/label[8]/input').click()
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[3]/a').click()
time.sleep(2)

# 爬取内容
text=[]

# 定位iframe，暂时没有实现自动翻页
# try:
iframe_element = driver.find_element_by_id('sosoIframe')
driver.switch_to.frame(iframe_element)
elements = driver.find_elements_by_xpath("//td[@name='result-list-title-td']/a")

# Create a new Excel
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('Result')

lenth = len(elements)
print(lenth)
worksheet.write(0, 0, label="招标内容（标名）")
worksheet.write(0, 1, label="链接")
for i in range (0,lenth):
    title = elements[i].get_attribute('title')
    url = elements[i].get_attribute('href')
    worksheet.write(i + 1, 0, label=title)
    worksheet.write(i + 1, 1, label=url)

print(len(driver.find_elements_by_xpath("//a[@text='下一页']")))
    # # 下面这个没有测试:(
    # driver.find_element_by_class_name('fanye_li02_20140617').click()
    # time.sleep(2)


# except Exception as e:
#     with open('result.json','wb') as f:
#         #:(写入文件没有实现
#         #json.dumps(text,f=f, ensure_ascii = False, indent = 4, sort_keys = True)
# with open('html.txt', 'w') as f:
#     f.write(source)
# print(source)
# time.sleep(3)
# workbook.save('Result.xls')
# time.sleep(20)
driver.quit()
