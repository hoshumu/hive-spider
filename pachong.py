'''
Author: asterisk
Date: 2020-08-30 17:32:36
LastEditTime: 2020-09-15 10:06:34
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

import urllib3
import base64

# get Token
host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=Va5yQRHlA4Fq5eR3LT0vuXV4&client_secret=0rDSjzQ20XUj5itV6WRtznPQSzr5pVw2&'

api = 'https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic'
access_token = ''

driver = webdriver.Chrome()

# Read cookies (Get cookies first)
with open('cookies.txt', 'r', encoding='utf8') as f:
    listCookies = json.loads(f.read())

# Open website and login
driver.get('https://www.okcis.cn/search/')# 打开网页
for cookie in listCookies:
    driver.add_cookie(cookie)
driver.refresh()

# 利用xpath输入文本，模拟点击
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[1]/div[1]/div[1]/input[2]').send_keys('环卫 保洁')
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[4]/div[1]/input').send_keys('车 厕')
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[5]/div[1]/label[7]/input').click()
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[6]/div[1]/label[8]/input').click()

# 选择省份
province_list = driver.find_elements_by_xpath("//div[@class='xb_20160118']/p")
province_list[0].find_element_by_xpath("./label/input").click()

# 开始搜索
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[3]/a').click()

time.sleep(2)

# 爬取内容
text=[]

# 定位iframe，暂时没有实现自动翻页
# try:
iframe_element = driver.find_element_by_id('sosoIframe')
driver.switch_to.frame(iframe_element)

# Create a new Excel
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('Result')

# lenth = len(elements)
# print(lenth)
worksheet.write(0, 0, label="招标内容（标名）")
worksheet.write(0, 1, label="链接")

page_list = int(driver.find_element_by_xpath("//ul[@class='fanye_ul_20140617']/li[10]/a").get_attribute("text"))
print(page_list)

for j in range(0,page_list - 1):

    time.sleep(8)
    print(j)

    elements = driver.find_elements_by_xpath("//td[@name='result-list-title-td']/a")
    lenth = len(elements)
    print(lenth)

    for i in range (0,lenth):
        
        title = elements[i].get_attribute('title')
        url = elements[i].get_attribute('href')
        worksheet.write(i + 1 + lenth * j, 0, label=title)
        worksheet.write(i + 1 + lenth * j, 1, label=url)

    driver.find_elements_by_xpath("//a[text()='下一页']")[0].click()


workbook.save('Result.xls')
time.sleep(10)
# driver.quit()
