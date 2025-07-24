# -*- coding: utf-8 -*-


from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import random


class CaptchaHandler:
    def __init__(self, driver):
        self.driver = driver

    def Taosliderl(self, max_iframes=6, timeout_per_iframe=1):
        """
        快速扫描页面中的 iframe 是否包含滑块验证结构。
        找到后等待人工完成滑块拖动。
        """
        try:
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            print(f"共发现 {len(iframes)} 个 iframe，开始检测滑块...")

            for i, iframe in enumerate(iframes[:max_iframes]):
                try:
                    self.driver.switch_to.frame(iframe)

                    # 检查是否包含滑块元素
                    slider = self.driver.find_elements(By.ID, "puzzle-captcha-btn")
                    if slider:
                        print(f" 第 {i+1} 个 iframe 中检测到滑块，请完成验证后按回车继续...")
                        input()
                        self.driver.switch_to.default_content()
                        print(" 验证完成，已返回主页面。")
                        return True  # 检测到并处理

                    self.driver.switch_to.default_content()

                except Exception as e:
                    print(f"切换第 {i+1} 个 iframe 异常: {e}")
                    self.driver.switch_to.default_content()
                    continue

            print(" 所有 iframe 均未检测到滑块，正常继续。")
            return False  # 未检测到

        except Exception as e:
            print(f"滑块检测异常: {e}")
            self.driver.switch_to.default_content()
            return False

    def TaoCheckRisk(self):
        """检测是否被强制退出登录或触发了风控"""
        try:
            # 检查是否存在登录按钮，表明可能被淘宝风控，需要重新登录
            login_btn = self.driver.find_element(By.CSS_SELECTOR, 'span.KOBWr4M6dz--loginBtn--_3aec1d3')
            if login_btn.is_displayed():
                print("检测到 '一键登录' 按钮，可能被淘宝风控，需重新登录！")
                return False
        except:
            # 如果没有找到“登录”按钮，说明正常登录状态
            pass

        # 检查风控页面是否出现，检测风控警告框
        try:
            wind_control_alert = self.driver.find_element(By.CSS_SELECTOR, '.KOBWr4M6dz--variantTip--_621ff16')
            if wind_control_alert.is_displayed():
                print("检测到风控警告框，可能因违规插件导致被风控，需重新登录！")
                return False
        except:
            # 正常没找到风控提示框，说明没有被风控
            return True
    
    def PDDsliderl(self, min_seconds=2, max_seconds=3, max_iframes=6):
        """
        检查当前页面是否跳转到拼多多的风控验证页，若找到验证按钮，则暂停程序等待用户手动完成验证。
        如果验证按钮在 iframe 中，则会扫描并切换到每个 iframe 进行检查。
        """
        try:
            # 首先检查主页面中是否包含验证按钮
            captcha_btn = self.driver.find_elements(By.CSS_SELECTOR, ".intel-btn")
            if captcha_btn:
                print(" 检测到拼多多风控页面，已暂停程序，建议降低抓取速度，请手动完成验证。")
                print(" 完成验证后，按下回车键继续...")
                input(" 等待中...")  # 等待用户完成验证
                return True  # 验证通过，继续执行

            # 如果主页面没有找到验证按钮，检查 iframe 中是否有
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            if len(iframes) > 0:
                print(f" 共发现 {len(iframes)} 个 iframe，开始检测滑块...")

                for i, iframe in enumerate(iframes[:max_iframes]):
                    try:
                        self.driver.switch_to.frame(iframe)

                        # 在每个 iframe 中查找验证按钮
                        captcha_btn_in_iframe = self.driver.find_elements(By.CSS_SELECTOR, ".intel-btn")
                        if captcha_btn_in_iframe:
                            print(f" 第 {i+1} 个 iframe 中检测到验证按钮，请完成验证后按回车继续...")
                            input(" 等待中...")  # 等待用户完成验证
                            self.driver.switch_to.default_content()
                            print(" 验证完成，已返回主页面。")
                            return True  # 验证通过，继续执行
                    
                        self.driver.switch_to.default_content()  # 切换回主页面

                    except Exception as e:
                        print(f" 切换第 {i+1} 个 iframe 异常: {e}")
                        self.driver.switch_to.default_content()
                        continue

            print(" 没有找到风控验证页面，继续抓取。")
            return False  # 没有检测到风控页面，继续执行抓取

        except Exception as e:
            print(f"风控验证检查异常: {e}")
            self.driver.switch_to.default_content()
            return False

    def JDslider1(self, min_seconds=2, max_seconds=3):
        """
        检查当前页面是否出现京东风控滑块组件，
        如果出现则暂停程序等待用户手动滑动验证，通过后按回车继续。
        否则正常等待。
        """
        # 查找滑块验证组件，判断是否为风控页面
        risk_elements = self.driver.find_elements(By.CLASS_NAME, 'JDJRV-slide-main')

        if risk_elements:
            print("检测到滑块风控页面，暂停程序等待人工处理。")
            print("完成滑块验证后请按回车继续...")
            input("等待中...")

            # 验证后再次检测
            risk_elements = self.driver.find_elements(By.CLASS_NAME, 'JDJRV-slide-main')
            if risk_elements:
                print("仍处于风控页面，验证可能失败。")
                return False
            else:
                print("验证通过，继续执行。")
                return True
        else:
            # 正常页面，等待
            time.sleep(random.uniform(min_seconds, max_seconds))
            return True

    def AliCaptcha(self):
        try:
            # 使用 self.driver 查找页面中特定的验证码元素
            captcha_element = self.driver.find_element_by_xpath("//div[@class='captcha']")
            if captcha_element:
                print("检测到验证码！暂停程序。")
                return True
        except Exception as e:
            # 如果没有找到验证码相关元素，返回 False
            return False
