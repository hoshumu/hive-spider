# hive-spider

## 项目简介

本项目是用于对[招标采购导航网](https://www.okcis.cn/)的一个基于Selenium的爬虫，目前能够爬取的信息有

- 招标内容
- 项目链接
- 所属地区
- 中标公司
- 中标公司链接
- 发布日期
- 联系人
- 联系电话

## 运行指南

1. pip通过`requirements.txt`文件安装依赖，同时安装`WebDriver`
2. 在`account.py`中填入账号信息（百度云图像识别模块暂时废弃可不填）
3. 运行`getcookies.py`获取cookies（cookies需要每天更新）
4. 运行`pachong.py`爬取数据，数据会自动存储到`Result.xls`

### 关于断点续搜

本爬虫会记录之前爬取的数据并读取已经爬取到的数据，如果需要重新爬取请删除`history.txt`文件

## Contribution

本项目由[@chaoers](https://github.com/chaoers)和[@uniartisan](https://github.com/uniartisan)共同完成

## LICENSE

本项目遵循`GPLV3.0`开源协议
