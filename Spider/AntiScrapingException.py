# -*- coding: utf-8 -*-

from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import random
from selenium.common.exceptions import NoSuchElementException


class CaptchaHandler:
    def __init__(self, driver, logger=None):
        self.driver = driver
        self.logger = logger

    def Taosliderl(self, max_iframes=6):
        """
        快速扫描页面中的 iframe 是否包含滑块验证结构。
        如果找到滑块验证，返回 True，表示需要人工处理。
        如果未找到滑块验证，返回 False。
        """
        try:
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            if self.logger:
                self.logger.info(f"共发现 {len(iframes)} 个 iframe，开始检测滑块...")

            for i, iframe in enumerate(iframes[:max_iframes]):
                try:
                    self.driver.switch_to.frame(iframe)
                    # 检查是否包含滑块元素
                    slider = self.driver.find_elements(By.ID, "puzzle-captcha-btn")
                    if slider:
                        if self.logger:
                            self.logger.info(f"第 {i+1} 个 iframe 中检测到滑块。")
                        self.driver.switch_to.default_content()
                        return True  # 检测到滑块验证，返回 True
                    self.driver.switch_to.default_content()
                except Exception as e:
                    if self.logger:
                        self.logger.error(f"切换第 {i+1} 个 iframe 异常: {e}")
                    self.driver.switch_to.default_content()
                    continue
        
            if self.logger:
                self.logger.info("所有 iframe 均未检测到滑块，正常继续。")
            return False  # 未检测到滑块

        except Exception as e:
            if self.logger:
                self.logger.error(f"滑块检测异常: {e}")
            self.driver.switch_to.default_content()
            return False

    def TaoCheckRisk(self):
        """检测是否被强制退出登录或触发了风控"""
        try:
            # 检查是否存在登录按钮，表明可能被淘宝风控，需要重新登录
            login_btn = self.driver.find_element(By.CSS_SELECTOR, 'span.KOBWr4M6dz--loginBtn--_3aec1d3')
            if login_btn.is_displayed():
                self.logger.warning("检测到 '一键登录' 按钮，可能被淘宝风控，需重新登录！")
                return True
        except:
            # 如果没有找到“登录”按钮，说明正常登录状态
            pass

        # 检查风控页面是否出现，检测风控警告框
        try:
            wind_control_alert = self.driver.find_element(By.CSS_SELECTOR, '.KOBWr4M6dz--variantTip--_621ff16')
            if wind_control_alert.is_displayed():
                self.logger.warning("检测到风控警告框，可能因违规插件导致被风控，需重新登录！")
                return True
        except:
            # 正常没找到风控提示框，说明没有被风控
            return False
    
    def PDDsliderl(self, min_seconds=2, max_seconds=3, max_iframes=6): 
        """
        检查当前页面是否跳转到拼多多的风控验证页。
        如果在主页面或 iframe 中检测到验证按钮，则返回 True，表示检测到风控。
        否则返回 False。
        """
        try:
            # 检查主页面是否有验证按钮
            if self.driver.find_elements(By.CSS_SELECTOR, ".intel-btn"):
                self.logger.warning("检测到拼多多风控页面（主页面）")
                return True

            # 检查 iframe 中是否有验证按钮
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            for i, iframe in enumerate(iframes[:max_iframes]):
                try:
                    self.driver.switch_to.frame(iframe)
                    if self.driver.find_elements(By.CSS_SELECTOR, ".intel-btn"):
                        self.logger.debug(f"检测到拼多多风控页面（iframe 第 {i+1} 个）")
                        self.driver.switch_to.default_content()
                        return True
                except Exception as e:
                    self.logger.debug(f"切换到第 {i+1} 个 iframe 失败: {e}")
                finally:
                    self.driver.switch_to.default_content()

            # 未检测到风控
            self.logger.info("未检测到拼多多风控页面")
            return False

        except Exception as e:
            self.logger.error(f"风控检测异常: {e}")
            self.driver.switch_to.default_content()
            return False

    def JDslider1(self):
        """
        检测当前页面是否出现京东风控滑块组件，
        有风控返回 True，否则返回 False。
        并使用 logger 记录检测结果。
        """
        try:
            risk_elements = self.driver.find_elements(By.CLASS_NAME, 'JDJRV-slide-main')
            if risk_elements:
                self.logger.warning("检测到京东滑块风控组件，需外部处理暂停或终止。")
                return True
            else:
                self.logger.info("未检测到京东滑块风控，继续正常爬取。")
                return False
        except NoSuchElementException:
            self.logger.info("未检测到京东滑块风控，继续正常爬取。")
            return False
        except Exception as e:
            self.logger.error(f"检测京东滑块风控时发生错误: {e}")
            return False

    def JDslider2(self):
        """
        检测当前页面是否出现京东风控验证组件（如‘快速验证’按钮）。
        如果有风控验证，返回 True；否则返回 False。
        并使用 logger 记录检测结果。
        """
        try:
            risk_elements = self.driver.find_elements(By.CLASS_NAME, 'verifyBtn')
            if risk_elements:
                self.logger.warning("检测到京东风控验证组件（快速验证按钮），需外部处理暂停或终止。")
                return True
            else:
                self.logger.info("未检测到京东风控验证组件，继续正常爬取。")
                return False
        except NoSuchElementException:
            self.logger.info("未检测到京东风控验证组件，继续正常爬取。")
            return False
        except Exception as e:
            self.logger.error(f"检测京东风控验证组件时发生错误: {e}")
            return False


    def AliCaptcha(self):
        """
        检测当前页面是否出现阿里风控验证码组件。
        如果有验证码，返回 True；否则返回 False。
        并使用 logger 记录检测结果。
        """
        try:
            captcha_element = self.driver.find_element(By.XPATH, "//div[@class='captcha']")
            if captcha_element:
                self.logger.warning("检测到阿里风控验证码组件，需外部处理暂停或终止。")
                return True
        except NoSuchElementException:
            # 明确捕获元素未找到异常，说明页面无验证码
            self.logger.info("未检测到阿里风控验证码组件，继续正常爬取。")
        except Exception as e:
            # 捕获其它异常，防止程序崩溃
            self.logger.error(f"检测验证码时发生异常: {e}")
        return False

    def A1688Captcha(self, min_seconds=2, max_seconds=3, max_iframes=6):
        """
        检测当前页面是否出现1688风控滑块验证码组件。
        如果检测到滑块，则返回 True，否则返回 False。
        并使用 logger 记录检测结果。
        优化为快速检索策略。
        """
        try:
            # 优化：直接检查主页面是否有风控滑块（旧版）
            if self.driver.find_elements(By.CSS_SELECTOR, ".captcha"):
                self.logger.warning("检测到1688风控滑块验证码（主页面 - 旧版）")
                return True

            # 新版风控：检查是否有滑块按钮（#scratch-captcha-btn）
            if self.driver.find_elements(By.CSS_SELECTOR, "#scratch-captcha-btn"):
                self.logger.warning("检测到1688风控滑块验证码（主页面 - 新版）")
                return True

            # 新版风控：检查是否有风控提示文字（如“请按照说明拖动滑块”）
            if self.driver.find_elements(By.XPATH, "//div[contains(text(), '请按照说明拖动滑块')]"):
                self.logger.warning("检测到1688风控提示文字（主页面 - 新版）")
                return True

            # 检查加载状态（如果存在加载动画）
            if self.driver.find_elements(By.CLASS_NAME, "scratch-captcha-loading"):
                self.logger.warning("检测到1688风控加载状态（主页面）")
                return True

            # 优化：减少不必要的 iframe 遍历次数
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            for i, iframe in enumerate(iframes[:max_iframes]):
                try:
                    self.driver.switch_to.frame(iframe)

                    # 快速检查 iframe 中是否包含风控滑块（旧版风控）
                    if self.driver.find_elements(By.XPATH, "//div[contains(@class, 'captcha')]"):
                        self.logger.debug(f"检测到1688风控滑块验证码（iframe 第 {i+1} 个 - 旧版）")
                        self.driver.switch_to.default_content()
                        return True

                    # 检查 iframe 中是否包含新版风控滑块按钮（#scratch-captcha-btn）
                    if self.driver.find_elements(By.CSS_SELECTOR, "#scratch-captcha-btn"):
                        self.logger.debug(f"检测到1688风控滑块验证码（iframe 第 {i+1} 个 - 新版）")
                        self.driver.switch_to.default_content()
                        return True

                    # 检查 iframe 中是否包含风控提示文字
                    if self.driver.find_elements(By.XPATH, "//div[contains(text(), '请按照说明拖动滑块')]"):
                        self.logger.debug(f"检测到1688风控提示文字（iframe 第 {i+1} 个 - 新版）")
                        self.driver.switch_to.default_content()
                        return True

                    # 检查 iframe 中是否有加载状态（加载动画）
                    if self.driver.find_elements(By.CLASS_NAME, "scratch-captcha-loading"):
                        self.logger.debug(f"检测到1688风控加载状态（iframe 第 {i+1} 个）")
                        self.driver.switch_to.default_content()
                        return True

                except Exception as e:
                    self.logger.debug(f"切换到第 {i+1} 个 iframe 失败: {e}")
                finally:
                    self.driver.switch_to.default_content()

            # 未检测到风控滑块
            self.logger.info("未检测到1688风控滑块验证码")
            return False

        except Exception as e:
            self.logger.error(f"风控检测异常: {e}")
            self.driver.switch_to.default_content()
            return False

    