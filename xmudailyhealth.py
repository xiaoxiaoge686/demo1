# -*- coding: UTF-8 -*-
import sys
import importlib
importlib.reload(sys)

import datetime
import json
import logging
import os
import re
import smtplib

# 负责构造图片
from email.mime.image import MIMEImage
# 负责将多个对象集合起来
from smtplib import SMTP_SSL
from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from time import sleep

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

# # 校园网用户名密码
# username = '用户名'
# password = '密码'

# # qq邮箱
# email = 'QQ邮箱'
# # qq邮箱里面设置
# auth = '授权码'


content = '<p>健康打卡链接如下：</p>' \
          '<p><a href="https://xmuxg.xmu.edu.cn/app/214">健康打卡</a></p>'
#content 内容设置
html_img = f'<p>{content}<br><img src="cid:image1"></br></p>' # html格式添加图片
email_server = 'smtp.qq.com'
email_title = datetime.datetime.now().strftime('%y-%m-%d-%H:%M:%S') + '厦大健康打卡'  # 邮件主题
def send_email():
    msg = MIMEMultipart() # 构建主体
    msg['Subject'] = Header(email_title,'utf-8')  # 邮件主题
    msg['From'] = send_usr  # 发件人
    msg['To'] = Header('健康打卡','utf-8') # 收件人--这里是昵称
    # msg.attach(MIMEText(content,'html','utf-8'))  # 构建邮件正文,不能多次构造
    attchment = MIMEApplication(open(r'./pic/{0}'.format(imagename),'rb').read()) # 文件
    attchment.add_header('Content-Disposition','attachment',filename=imagename)
    msg.attach(attchment)  # 添加附件到邮件
    f = open(r'./pic/{0}'.format(imagename), 'rb')  #打开图片
    msgimage = MIMEImage(f.read())
    f.close()
    msgimage.add_header('Content-ID', '<image1>')  # 设置图片
    msg.attach(msgimage)
    msg.attach(MIMEText(html_img,'html','utf-8'))  # 添加到邮件正文
    try:
        smtp = SMTP_SSL(email_server)  #指定邮箱服务器
        smtp.ehlo(email_server)   # 部分邮箱需要
        smtp.login(send_usr,send_pwd)  # 登录邮箱
        smtp.sendmail(send_usr,reverse,msg.as_string())  # 分别是发件人、收件人、格式
        smtp.quit()  # 结束服务
        print('邮件发送完成--')
    except:
        print('发送失败')


def del_files(path):
    for root, dirs, file in os.walk(path):
        for name in file:
            if name.endswith(".png"):
                os.remove(os.path.join(root, name))


if __name__ == '__main__':

    json_data = []
    with open('./config.json', 'r', encoding='utf8') as fp:
        json_data = json.load(fp)
        fp.close()

    username = json_data['username']
    password = json_data['password']
    send_usr = json_data['sender']
    send_pwd = json_data['auth']
    reverse = json_data['receiver']


    mylog = logging.getLogger('mylogger')
    mylog.setLevel(logging.INFO)

    # 处理器
    handler = logging.FileHandler('./logs/log_test.txt')
    handler.setLevel(logging.DEBUG)
    # 格式器
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)

    mylog.addHandler(handler)

    # 不显示浏览器
    desired_capabilities = DesiredCapabilities.EDGE
    desired_capabilities["pageLoadStrategy"] = "none"

    service = Service(executable_path="./msedgedriver.exe")
    driver = webdriver.Edge(service = service)
    # option = webdriver.ChromeOptions()
    # option.add_argument("headless")
    # 通过  edge://version/ 查看浏览器版本
    # https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/ 下载对应版本插件
    # 解压填入路径
    # driver = webdriver.Chrome("./chromedriver.exe", chrome_options=option)

    # 打卡网站
    # 最大化界面
    # driver.maximize_window()
    driver.get('https://xmuxg.xmu.edu.cn/xmu/login')
    driver.implicitly_wait(40)

    # 统一身份认证按钮
    driver.find_elements(By.CLASS_NAME, "primary-btn")[2].click()
    sleep(2)
    # 用户密码登录
    driver.find_element(By.ID, "username").clear()
    driver.find_element(By.ID, "password").clear()
    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    # 点击登录
    driver.find_element(By.CLASS_NAME, "auth_login_btn").click()
    sleep(5)
    # 这里修改了原作者的语句，把网址变了，在获取了cookie之后刷新一遍即可
    # driver.find_elements_by_class_name("app_child")[2].click()
    # 跳转到新网页
    driver.switch_to.window(driver.window_handles[-1])
    driver.get('https://xmuxg.xmu.edu.cn/app/214')
    driver.implicitly_wait(40)

    # 停止一下，不然太快
    sleep(5)
    # 点击我的表单
    driver.find_elements(By.CLASS_NAME,  "tab")[1].click()
    sleep(3)
    # 判断是否已经打卡
    text = driver.find_element(By.XPATH, "//div[@data-name='select_1582538939790']").text
    if('Yes' in text):
        #driver.find_element(By.XPATH, "//button[contains(.,'日志')]").click()
        driver.find_element(By.ID, "datetime_1611146487222").click()

        sleep(3)
        snipname = './pic/' + datetime.datetime.now().strftime('%y%m%d%H%M%S') + '.png'
        imagename = datetime.datetime.now().strftime('%y%m%d%H%M%S') + '.png'
        driver.get_screenshot_as_file(snipname)
        #sentemail('今日已经打卡', snipname, sender, auth, recevier)
        send_email()
        # sentemail("今日已经打卡")
        mylog.info("今日已经打卡")
    else:
        # vue和传统下拉框不一样
        driver.find_element(By.CSS_SELECTOR, "[data-name='select_1582538939790']").click()
        sleep(3)
        driver.find_element(By.CSS_SELECTOR, "[title='是 Yes'][class='btn-block']").click()
        # 点击保存
        driver.find_element(By.CLASS_NAME, "form-save").click()
        sleep(3)
        # 点击弹出框确定
        driver.switch_to.alert.accept()
        driver.refresh()
        sleep(5)
        # 点击我的表单
        driver.find_elements(By.CLASS_NAME, "tab")[1].click()
        sleep(3)
        #driver.find_element(By.XPATH, "//button[contains(.,'日志')]").click()
        driver.find_element(By.ID, "datetime_1611146487222").click()


        sleep(3)
        snipname = './pic/' + datetime.datetime.now().strftime('%y%m%d%H%M%S') + '.png'
        imagename = datetime.datetime.now().strftime('%y%m%d%H%M%S') + '.png'
        driver.get_screenshot_as_file(snipname)
        send_email()
        mylog.info("打卡成功")

    sleep(1)
    # 采用close()执行脚本的时候不会自动退出
    del_files("./pic/")
    driver.quit()
