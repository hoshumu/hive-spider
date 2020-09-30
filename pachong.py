# coding:utf-8
import sys
import time
import json

# execl处理模块
import xlwt
import xlrd
from xlutils.copy import copy

# selenium模块
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 读取设置的账号密码
import account

driver = webdriver.Chrome()
driver.maximize_window()

# 打开多个窗口
js = 'window.open();'
driver.execute_script(js)
driver.execute_script(js)
windows = driver.window_handles
driver.switch_to.window(windows[0])

# 读取Cookies并登录
with open('cookies.txt', 'r', encoding='utf8') as f:
    listCookies = json.loads(f.read())
driver.get('https://www.okcis.cn/search/')  # 打开网页
for cookie in listCookies:
    driver.add_cookie(cookie)
driver.refresh()

# 尝试读取爬取历史
try:
    with open('./history.txt', 'r') as f:
        province_start = int(f.readline())
        city_start = int(f.readline())
        page_start = int(f.readline())
        item_start = int(f.readline())
        sheet_index = int(f.readline())
        print(province_start, city_start, page_start, item_start)
        f.close()
    workbook = copy(xlrd.open_workbook('./Result.xls'))
    worksheet = workbook.get_sheet(-1)  # 获取工作簿中所有表格中的的第一个表格
    print('成功读取历史记录')
    has_his = 1

except:
    print(sys.exc_info()[0])
    print('读取历史记录失败，初始化中')
    province_start = 0
    city_start = 0
    page_start = 0
    item_start = 0
    # Create a new Excel
    workbook = xlwt.Workbook(encoding='utf-8')
    has_his = 0

# 搜索标题栏
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[1]/div[1]/div[1]/input[2]').send_keys('环卫')
# 去掉关键字
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[4]/div[1]/input').send_keys('车 厕')
# 点击中标结果btn
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[5]/div[1]/label[7]/input').click()
# 点击一年内btn
driver.find_element_by_xpath(
    '/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[6]/div[1]/label[8]/input').click()
# 点击标题搜索btn
driver.find_element_by_xpath(
    "/html/body/div[4]/div[2]/form/div/div[2]/div[2]/div[2]/div[1]//label[2]/input").click()

print("10秒后开始搜索")
time.sleep(10)

# 遍历省份列表
province_list = driver.find_elements_by_xpath("//div[@class='xb_20160118']/p")
print(len(province_list))

for province_index in range(province_start, len(province_list)):
    try:
        # 获得省份名
        province = driver.find_elements_by_xpath(
            "//div[@class='xb_20160118']/p")[province_index]
        province_name = province.find_element_by_xpath(
            "./label/span").get_attribute("textContent")
        province.click()
        print(province_name)
        time.sleep(0.5)

        # 获取城市列表
        city_list = driver.find_elements_by_xpath(
            "//div[@class='sj_d_20160118']/div[last()]//label[@class='sdq_qbf_20160118']")

        # 读取到爬取历史特殊处理
        if not has_his:
            worksheet = workbook.add_sheet(province_name)
            worksheet.write(0, 0, label="招标内容（标名）")
            worksheet.write(0, 1, label="链接")
            worksheet.write(0, 2, label="地区")
            worksheet.write(0, 3, label="中标公司")
            worksheet.write(0, 4, label="中标公司链接")
            worksheet.write(0, 5, label="发布日期")
            worksheet.write(0, 6, label="联系人")
            worksheet.write(0, 7, label="联系电话")
            sheet_index = 1

    except:
        print("error!skip this province")
        continue

    for city_index in range(city_start, len(city_list)):
        try:
            # 获取城市名
            print(len(driver.find_elements_by_xpath(
                "//div[@class='sj_d_20160118']/div[last()]//label[@class='sdq_qbf_20160118']")), city_index)
            city = driver.find_elements_by_xpath(
                "//div[@class='sj_d_20160118']/div[last()]//label[@class='sdq_qbf_20160118']")[city_index]
            city.click()
            time.sleep(1)
            city_name = city.find_element_by_xpath(
                "./span").get_attribute("textContent")
            print(city_name)

            # 开始搜索
            driver.find_element_by_xpath(
                "//div[@class='ljss_20160118']").click()
            time.sleep(8)

            # 定位iframe
            iframe_element = driver.find_element_by_id('sosoIframe')
            driver.switch_to.frame(iframe_element)

            # 获取页数
            page_list = int(driver.find_element_by_xpath(
                "//ul[@class='fanye_ul_20140617']/li[last()-2]/a").get_attribute("text"))
            print(page_list)

            for j in range(page_start, page_list):

                time.sleep(8)
                print(page_start, j, page_list)

                # 超过20页特殊处理
                if j >= 19:
                    break
                
                # 获取单页项目列表
                elements = driver.find_elements_by_xpath(
                    "//td[@name='result-list-title-td']/a")

                for i in range(item_start, len(elements)):
                    try:
                        # 获取单个项目信息
                        title = elements[i].get_attribute('title')
                        url = elements[i].get_attribute('href')
                        worksheet.write(sheet_index, 0, label=title)
                        worksheet.write(sheet_index, 1, label=url)
                        worksheet.write(sheet_index, 2, label=city_name)

                        # 打开新窗口并跳转项目详细信息
                        windows = driver.window_handles
                        driver.switch_to.window(windows[1])
                        driver.get(url)
                        time.sleep(3)
                    except:
                        print("error!skip this item")
                        continue
                    try:
                        # 获取时间信息
                        date = driver.find_element_by_xpath(
                            "//td[contains(text(),'更新时间')]/following-sibling::td[1]")
                        worksheet.write(
                            sheet_index, 5, label=date.get_attribute('textContent'))
                    except:
                        print('获取日期失败')
                        worksheet.write(sheet_index, 5, label='未标出')
                    try:
                        # 获取公司信息
                        company = driver.find_element_by_xpath(
                            "//td[contains(text(),'中标候选')]/following-sibling::td[1]/div/a[1]")
                    except:
                        print('获取公司')
                        worksheet.write(sheet_index, 3, label='未标出')
                        worksheet.write(sheet_index, 4, label='未标出')
                        worksheet.write(sheet_index, 6, label='未标出')
                        worksheet.write(sheet_index, 7, label='未标出')
                    else:
                        worksheet.write(
                            sheet_index, 3, label=company.get_attribute('textContent'))
                        worksheet.write(
                            sheet_index, 4, label=company.get_attribute('href'))
                        try:
                            # 打开公司页面并获取公司联系信息
                            driver.get(company.get_attribute('href'))
                            time.sleep(3)
                            name = driver.find_element_by_xpath(
                                "//td/samp[@id='page3_con_1_name']").get_attribute('textContent')
                            company_phone = driver.find_element_by_xpath(
                                "//td/samp[@id='page3_con_2']").get_attribute('textContent')
                            # print(name, company_phone)
                        except:
                            print('获取公司联系人失败')
                            worksheet.write(sheet_index, 6, label='未标出')
                            worksheet.write(sheet_index, 7, label='未标出')
                        else:
                            worksheet.write(
                                sheet_index, 6, label=name)
                            worksheet.write(
                                sheet_index, 7, label=company_phone)

                    # 保存数据表并初始化读取历史变量
                    workbook.save('Result.xls')
                    has_his = 0
                    province_start = 0
                    city_start = 0
                    page_start = 0
                    item_start = 0
                    windows = driver.window_handles
                    driver.switch_to.window(windows[0])
                    iframe_element = driver.find_element_by_id('sosoIframe')
                    driver.switch_to.frame(iframe_element)
                    sheet_index = sheet_index + 1

                    # 记录爬取历史以便断电继续工作
                    fo = open("./history.txt", "w")
                    fo.writelines(str(province_index) + '\n' + str(city_index) +
                                  '\n' + str(j) + '\n' + str(i) + '\n' + str(sheet_index))
                    fo.close()

                # 点击翻页btn
                driver.find_elements_by_xpath("//a[text()='下一页']")[0].click()

            # 退出iframe
            driver.switch_to.parent_frame()

            # 点击上个城市btn
            time.sleep(0.5)
            last_city = driver.find_elements_by_xpath(
                "//div[@class='sj_d_20160118']/div[last()]//label[@class='sdq_qbf_20160118']")[city_index].click()
            time.sleep(1)

        except:
            print("error!skip this city")
            driver.switch_to.parent_frame()
            time.sleep(0.5)
            last_city = driver.find_elements_by_xpath(
                "//div[@class='sj_d_20160118']/div[last()]//label[@class='sdq_qbf_20160118']")[city_index].click()
            time.sleep(1)
            workbook.save('Result.xls')
            continue

    # last_province = driver.find_elements_by_xpath("//div[@class='xb_20160118']/p")[province_index].click()

    # time.sleep(0.5)
time.sleep(10)
driver.quit()
