import json
import requests
import urllib3
import base64
from PIL import Image,ImageEnhance
from io import BytesIO

# 读取设置的账号密码
import account

def getNum(img):
    # get Token
    https = urllib3.PoolManager()
    host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=' + \
        account.api_key + '&client_secret=' + account.secret_key
    headers = {
        'Content-Type': 'application/json;charset=UTF-8'
    }
    res = https.request('GET', host, headers=headers)
    print(res.status)
    res_data = json.loads(res.data.decode('utf-8'))
    token = res_data['access_token']

    api = 'https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token=' + token
    imgR = base64.b64encode(img)
    params = {"image":imgR}
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(api, data=params, headers=headers)
    print(response.status_code)
    if response:
      num = ''
      num_str = ''
      for item in response.json()['words_result']:
        num = num + item['words']
      for str in num:
        if str != '=' and str != '?':
          num_str = num_str + str
      return num_str

def getCodeImg(driver, img, iframe_x = 0, iframe_y = 0):
        imgPath = './code.png'

        chromeSize = driver.get_window_size()
        location = img.location
        location['x'] = location['x'] + iframe_x
        location['y'] = location['y'] + iframe_y
        imgSize = img.size
        left = location['x']/chromeSize['width'] 
        top = location['y']/chromeSize['height']                           
        right = (location['x'] + imgSize['width'])/chromeSize['width']     
        bottom = (location['y'] + imgSize['height'])/chromeSize['height']  
        driver.get_screenshot_as_file(imgPath)
        screenshotImgSize = Image.open(imgPath).size
        img = Image.open(imgPath).convert('RGB').crop((
            left * screenshotImgSize[1] + 70,                
            top * screenshotImgSize[0] - 190,                 
            right * screenshotImgSize[1] + 270,               
            bottom * screenshotImgSize[0] - 250                
        ))
        # img = img.convert('L')                              # 转换模式：L | RGB
        img = ImageEnhance.Contrast(img)                    # 增强对比度
        img = img.enhance(0.8)                              # 增加饱和度
        output_buffer = BytesIO()
        img.save(output_buffer, format='JPEG')
        img.save(imgPath, format='JPEG')
        byte_data = output_buffer.getvalue()
        num = getNum(byte_data)
        return num