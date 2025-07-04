# -*- coding: gbk -*-

from seleniumwire import webdriver
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import random
import time
import os
import sys


class BaseScraper:
    def __init__(self, headless=True, proxy=None):
        self.driver = self.init_driver(headless=headless, proxy=proxy)

    def init_driver(self):
        # 构建资源路径（兼容打包和开发环境）
        def resource_path(relative_path):
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        # 固定路径
        chrome_path = resource_path("Chrome_Tool/GoogleChromePortable/GoogleChromePortable.exe")
        chromedriver_path = resource_path("Chrome_Tool/chromedriver-win64/chromedriver.exe")

        if not os.path.exists(chrome_path):
            print("Chrome 浏览器路径不存在：", chrome_path)
            return None
        if not os.path.exists(chromedriver_path):
            print("ChromeDriver 路径不存在：", chromedriver_path)
            return None

        # 配置 Chrome 启动项
        options = uc.ChromeOptions()
        options.binary_location = chrome_path
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--start-maximized')

        # 启动 ChromeDriver
        try:
            driver = uc.Chrome(
                options=options,
                driver_executable_path=chromedriver_path,
                use_subprocess=True
            )
        except Exception as e:
            print("Chrome 启动失败：", e)
            return None

        # 请求拦截器（预留）
        def interceptor(request):
            pass

        # 响应拦截器，伪装 webdriver 特征
        def response_interceptor(request, response):
            if 'react_psnl_verification_' in request.path:
                body = response.body.decode('utf-8', errors='ignore')
                modified = body.replace('navigator.webdriver', 'navigator.qwerasdfzxcv')
                response.body = modified.encode('utf-8')

        try:
            driver.request_interceptor = interceptor
            driver.response_interceptor = response_interceptor
        except Exception as e:
            print("设置拦截器失败：", e)

        time.sleep(1)  # 稳定启动

        return driver