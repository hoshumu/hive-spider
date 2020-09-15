# coding:utf-8
import sys
import io
import time
import json
import requests

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.action_chains import ActionChains

driver = webdriver.Chrome()

driver.get('https://www.okcis.cn/search/')  # 打开网页
ele=driver.find_element_by_xpath('/html/body/div[2]/div[2]/div/ul/li[4]/div/a')
ActionChains(driver).move_to_element(ele).perform()

# Cookie 需要每天更新
time.sleep(20)      # 手工输入账号密码验证码
driver.refresh()
cookies = driver.get_cookies()
jsoncookies = json.dumps(cookies)
with open('cookies.txt', 'w') as f:
    f.write(jsoncookies)
print('Get Cookies successfully')

driver.quit()
