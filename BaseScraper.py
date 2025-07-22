# -*- coding: gbk -*-

from seleniumwire import webdriver
import undetected_chromedriver as uc
import random
import time
import os
import sys
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class BaseScraper:
    def __init__(self, headless=True, proxy=None):
        self.logger = self._init_logger()  # 先初始化日志
        self.driver = self.init_driver(headless=headless, proxy=proxy)

    def _init_logger(self):
        logger = logging.getLogger("Spider")
        logger.setLevel(logging.DEBUG)  # 可调为 INFO 或 ERROR

        # 创建日志目录
        if not os.path.exists("logs"):
            os.makedirs("logs")

        # 创建文件处理器，日志文件按时间或大小轮转也可用RotatingFileHandler
        file_handler = logging.FileHandler("logs/Mypider.log", encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)

        # 创建控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        # 格式化输出
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        # 添加处理器（避免重复添加）
        if not logger.hasHandlers():
            logger.addHandler(file_handler)
            logger.addHandler(console_handler)

        return logger

    def save_page_html(self, filename="page_snapshot.html", wait_for_element=None, timeout=10):
        """ 保存当前页面的 HTML 内容到文件，确保页面完全加载 """
        try:
            # 等待页面上某个特定元素加载完成（如果需要）
            if wait_for_element:
                WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, wait_for_element))
                )
        
            # 获取当前页面的 HTML 内容
            html_content = self.driver.page_source

            # 使用时间戳避免文件名重复
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            file_path = f"logs/{timestamp}_{filename}"

            # 确保 logs 目录存在
            if not os.path.exists("logs"):
                os.makedirs("logs")

            # 保存 HTML 到文件
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(html_content)

            self.logger.info(f"页面 HTML 已保存到: {file_path}")
        except Exception as e:
            self.logger.warning(f"保存页面 HTML 时出错: {e}")


    def init_driver(self, headless=None, proxy=None):
        # 构建资源路径（兼容打包和开发环境）
        def resource_path(relative_path):
            if hasattr(sys, '_MEIPASS'):
                return os.path.join(sys._MEIPASS, relative_path)
            return os.path.join(os.path.abspath("."), relative_path)

        # 固定路径
        chrome_path = resource_path("Chrome_Tool/GoogleChromePortable/GoogleChromePortable.exe")
        chromedriver_path = resource_path("Chrome_Tool/chromedriver-win64/chromedriver.exe")

        if not os.path.exists(chrome_path):
            self.logger.error("Chrome 浏览器路径不存在：%s", chrome_path)
            return None
        if not os.path.exists(chromedriver_path):
            self.logger.error("ChromeDriver 路径不存在：%s", chromedriver_path)
            return None

        # 配置 Chrome 启动项
        options = uc.ChromeOptions()
        options.binary_location = chrome_path
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--start-maximized')

        # 如果需要无头模式
        if headless:
            options.add_argument('--headless=new')

        # 启动 ChromeDriver
        try:
            driver = uc.Chrome(
                options=options,
                driver_executable_path=chromedriver_path,
                use_subprocess=True
            )
            self.logger.info("ChromeDriver 启动成功")
        except Exception as e:
            self.logger.warning("Chrome 启动失败：%s", e)
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
            self.logger.info("请求/响应拦截器设置成功")
        except Exception as e:
            self.logger.warning("设置拦截器失败：%s", e)

        time.sleep(1)  # 稳定启动
        return driver
        
        
