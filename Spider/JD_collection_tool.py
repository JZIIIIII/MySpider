# -*- coding: utf-8 -*-
import time
import random
import json
import re
import requests
import os
import shutil
from io import BytesIO
from PIL import Image
from PIL import UnidentifiedImageError
from seleniumwire import webdriver  # selenium-wire 用于拦截请求
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from pyquery import PyQuery as pq
from openpyxl import Workbook
from Spider.BaseScraper import BaseScraper

from Spider.AntiScrapingException import CaptchaHandler
from path_utils import static_image_path ,static_hash_path , static_excel_path


class JDScraper(BaseScraper):
    def __init__(self, keyword, start_page, end_page, insert_image = True, max_items=100):
        super().__init__(headless=True, proxy=None)
        self.keyword = keyword
        self.page_start = start_page
        self.page_end = end_page
        self.max_items = max_items
        self.wait = WebDriverWait(self.driver, 10)
        self.excel = Workbook()
        self.sheet = self.excel.active
        self.count = 2
        self.insert_image = insert_image
        self._setup_excel()
        self.hash_set = set()
        #self.lock = threading.Lock()  # 用于同步的锁
        #self.anti_spider_triggered = False
        self.captcha_handler = CaptchaHandler(self.driver, self.logger)
        self.captcha_handler.JDslider1()  # 调用风控
        self.captcha_handler.JDslider2()

    def _setup_excel(self):
        headers = ['Num', 'Title', 'Price', 'Deal', 'Location', 'Shop', 'IsPostFree',
                   'Title_URL', 'Shop_URL', 'Img_URL', 'Num_Com', 'Image']
        for i, header in enumerate(headers, 1):
            self.sheet.cell(row=1, column=i, value=header)


    '''
    def save_cookies(self, path="JD_cookies.json"):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.driver.get_cookies(), f)

    def load_cookies(self, path="JD_cookies.json"):
        with open(path, "r", encoding="utf-8") as f:
            cookies = json.load(f)
            for cookie in cookies:
                self.driver.add_cookie(cookie)
        self.driver.refresh()

    def login(self, user, password):#没做
        print("正在尝试登录JD...")
        try:
            iframe = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, '//iframe[contains(@src, "login.taobao.com")]')))
            self.driver.switch_to.frame(iframe)
            self.driver.find_element(By.ID, 'fm-login-id').send_keys(user)
            self.human_sleep(1, 2)
            self.driver.find_element(By.ID, 'fm-login-password').send_keys(password)
            self.human_sleep(1, 2)
            self.driver.find_element(By.CSS_SELECTOR, 'button.fm-button.fm-submit.password-login').click()
            self.human_sleep(1, 2)
            input("如出现滑块，请手动完成验证后按 Enter 继续...")
            self.save_cookies()
        except Exception as e:
            self.logger.warning(print("登录失败:", e))
    '''

    def search(self):
        try:
            self.logger.info("尝试用定位搜索框和按钮...")
            search_box = self.wait.until(EC.element_to_be_clickable((By.ID, 'key')))
            search_btn = self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'button')))
        except Exception as e:
            self.logger.error("搜索框定位失败")
        try:
            search_box.clear()
            search_box.send_keys(self.keyword)
            self.human_sleep(1, 2)
            search_btn.click()
            self.human_sleep(1, 2)
            self.RiskPause(self.captcha_handler.JDslider1())
            self.RiskPause(self.captcha_handler.JDslider2())
        except Exception as e:
            self.logger.warning(f"搜索操作失败:{e}")

    def go_to_page(self, page_number):
        self.RiskPause(self.captcha_handler.JDslider1())
        self.RiskPause(self.captcha_handler.JDslider2())


        try:
            # 定位输入框
            input_elem = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, '//div[contains(@class, "_pagination_toPageNum_")]//input')
            ))
            input_elem.clear()
            input_elem.send_keys(str(page_number))
            input_elem.send_keys(Keys.ENTER)  # 使用 Enter 键触发跳转

            # 等待页面跳转到目标页
            self.wait.until(EC.text_to_be_present_in_element(
                (By.XPATH, '//div[contains(@class, "_pagination_item_") and contains(@class, "_active_")]'),
                str(page_number)
            ))

            self.human_sleep(2, 3)
            self.RiskPause(self.captcha_handler.JDslider1())
            self.RiskPause(self.captcha_handler.JDslider2())
        except Exception as e:
            self.logger.warning(f"跳转到第 {page_number} 页失败: {e}")
            self.save_page_html("error_page.html")

    def get_more(self, url):
        post_text = ''
        main_window = self.driver.current_window_handle

        try:
            # 提取商品ID，防止 href 匹配失败
            item_id_match = re.search(r'(\d+)', url)
            if not item_id_match:
                self.logger.warning("无法提取商品ID，跳过")
                return ''
            item_id = item_id_match.group(1)

            # 定位商品卡片（通过商品ID来定位卡片，而非点击 <a> 标签）
            card_xpath = f'//div[@data-sku="{item_id}"]//div[@class="_card_2xp6d_40"]'
            card_element = self.wait.until(EC.element_to_be_clickable((By.XPATH, card_xpath)))

            # 模拟点击卡片，确保点击区域正确
            self.logger.debug(f"即将点击商品卡片：{item_id}")
            actions = ActionChains(self.driver)
            actions.click(card_element).perform()  # 点击商品卡片本身，而非 <a> 标签
            self.human_sleep(1, 2)

            # 等待新标签页打开
            for _ in range(10):
                handles = self.driver.window_handles
                if len(handles) > 1:
                    break
                time.sleep(0.5)
            else:
                self.logger.warning("新标签页未打开")
                self.save_page_html("error_page.html")
                return ''

            # 切换到新标签页
            handles = self.driver.window_handles
            new_tabs = [h for h in handles if h != main_window]
            if not new_tabs:
                self.logger.warning("找不到新标签页句柄")
                self.save_page_html("error_page.html")
                return ''

            new_tab = new_tabs[-1]
            self.driver.switch_to.window(new_tab)
            self.human_sleep(1, 2)
            self.RiskPause(self.captcha_handler.JDslider1())  # 处理验证码
            self.RiskPause(self.captcha_handler.JDslider2())
            # 等待页面加载并获取包邮信息
            try:
                self.wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//span[contains(@class, "free-shipping")]')
                ))
                # 获取包邮信息
                post_text = self.driver.find_element(
                    By.XPATH,
                    '//span[contains(@class, "free-shipping")]'
                ).text.strip()
            except Exception:
                post_text = ''  # 如果没有找到包邮信息，返回空字符串
            
            try:
                comment_elem = self.wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//div[contains(@class, "comment-title")]')
                ))
                comment_text = comment_elem.text.strip()
                match = re.search(r'买家评价\(([^)]+)\)', comment_text)
                if match:
                    num_com = match.group(1)  # 例如 "500+"
                else:
                    num_com = 0
            except Exception:
                num_com = 0

        except Exception as e:
            self.logger.warning(f"详情失败: {e}")

        finally:
            # 关闭新标签页并切回主窗口
            self.human_sleep(1, 2)
            try:
                if self.driver.current_window_handle != main_window:
                    self.driver.close()
                    self.driver.switch_to.window(main_window)
                    self.RiskPause(self.captcha_handler.JDslider1())  # 处理验证码
                    self.RiskPause(self.captcha_handler.JDslider2())
            except Exception as e:
                self.logger.warning(f"关闭标签页或切换回主窗口失败: {e}")
                self.save_page_html("error_page.html")

        return post_text,num_com  # 返回包邮信息（例如：“包邮” 或 空字符串）

    def download_image(self, url, index):
        try:
            # 补全 url
            if url.startswith("//"):
                url = "https:" + url
            elif url.startswith("/"):
                url = "https://img.alicdn.com" + url

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Referer": "https://www.jd.com/",
            }

            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code != 200:
                self.logger.error(f"图片请求失败: {url}")
                return None

            try:
                # 读取图片并转成 RGB
                img = Image.open(BytesIO(response.content)).convert("RGB")
            except UnidentifiedImageError:
                self.logger.error(f"图片解码失败（格式可能不受支持，如webp）: {url}")
                return None
            except Exception as e:
                self.logger.error(f"图片处理异常: {e} | URL: {url}")
                return None

            # 创建文件夹（硬编码路径）
            folder = "images/JD"  # 使用硬编码路径
            if not os.path.exists(folder):
                os.makedirs(folder)

            # 文件名硬编码处理
            file_name = f"{index:04d}.jpg"  # 保持文件名格式不变
            file_path = os.path.join(folder, file_name)  # 拼接路径

            # 保存成 jpg 文件
            img.save(file_path, format="JPEG")

            # 返回文件路径
            return file_path

        except Exception as e:
            self.logger.error(f"图片下载异常: {e} | URL: {url}")
            return None

    def clean_image_folder(self):
        # 获取图片文件夹路径并清空（硬编码路径）
        folder = "images/JD"  # 使用硬编码路径
        if os.path.exists(folder):
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)  # 删除文件或快捷方式
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)  # 删除子文件夹
                except Exception as e:
                    self.logger.warning(f"[!] 删除缓存文件失败: {file_path} | 错误: {e}")
        else:
            self.logger.error(f"[!] 文件夹不存在或不是目录：{folder}")




    def parse_page(self, page_number, platform='JD',hash_json=None):
        html = self.driver.page_source
        doc = pq(html)
        items = doc('div._wrapper_2xp6d_3.plugin_goodsCardWrapper._row_6_2xp6d_13').items()
        slide_counter = 0  # 每页开始前初始化下滑计数器
        # 载入历史哈希
        if hash_json:
            hash_json = static_hash_path(hash_json)
            hash_set = self.load_hash_set(hash_json)
            self.hash_set.update(hash_set)  # 合并新的哈希池
        else:
            hash_set = self.hash_set  # 使用共享的哈希池

        for item in items:
            # 如果暂停，则等待恢复
            self.wait_if_paused()
            # 检查是否被强制终止
            if self.should_stop():
                self.logger.info("检测到提前终止命令，保存已抓取内容并退出 parse_page。")
                if hash_json:
                    self.save_hash_set(self.hash_set, hash_json)  # 使用类成员变量 self.hash_set
                return False
            try:
                # 检查是否达到最大抓取数量
                if self.count - 2 >= self.max_items:
                    self.logger.info("已达到最大抓取数量：{self.max_items}，停止抓取。")
                    if hash_json:
                        self.save_hash_set(self.hash_set, hash_json)       
                    return False
                
                
                '''
                if self.anti_spider_triggered == True:
                    print(f"触发强制风控账号冻结 应前往京东APP解封")
                    self.save_hash_set(hash_set, hash_json)            
                    return False
                '''

    #            if item.find('.title--RoseSo8H').text() or item.find('.headTitleText--hxVemljn').text():
    #                continue
                self.RiskPause(self.captcha_handler.JDslider1())
                self.RiskPause(self.captcha_handler.JDslider2())

                # 获取图片标签
                img_tag = item.find('img')

                # 处理懒加载情况，优先选择 data-lazy-img 或者 src
                img_url = img_tag.attr('data-lazy-img') or img_tag.attr('src')

                # 判断是否为相对路径，如果是，补全为完整的 URL
                if img_url and img_url.startswith('//'):
                    img_url = 'https:' + img_url  # 补全相对路径为完整 URL

                #print(img_url)

                # 依次尝试 src / data-src / src2 / data-ks-lazyload / data-lazyload
                img_url = (
                    img_tag.attr('src')
                    or img_tag.attr('data-src')
                    or img_tag.attr('src2')
                    or img_tag.attr('data-ks-lazyload')
                    or img_tag.attr('data-lazyload')
                    or ""
                ).strip()

                # 修复以 // 开头或缺少协议的情况
                if img_url.startswith("//"):
                    img_url = "https:" + img_url
                elif img_url.startswith("/"):
                    img_url = "https://img.alicdn.com" + img_url


                # 标题
                title_elem = item.find('span._text_1x4i2_30')
                title = title_elem.attr('title') if title_elem else ''

                # 价格
                price_elem = item.find('span._price_1tn4o_13')
                price_text = price_elem.text().strip() if price_elem else ''
                import re
                price_num_str = re.sub(r'[^\d.]', '', price_text)
                try:
                    price = float(price_num_str)
                except:
                    price = 0.0

                # 销量提取
                deal_elem = item.find('span._goods_volume_1xkku_1 span[title]')
                deal_text = deal_elem.attr('title') if deal_elem else ''
                deal = deal_text.replace('已售', '').strip() if deal_text else ''

                # 无地址
                location = ''

                # 店铺名称和链接提取
                shop_elem = item.find('a._name_d19t5_35')
                shop_name = shop_elem.find('span').text().strip() if shop_elem else ''
                shop = shop_name

                shop_href = shop_elem.attr('href') if shop_elem else ''
                if shop_href:
                    if shop_href.startswith('//'):
                        shop_url = 'https:' + shop_href
                    elif shop_href.startswith('/'):
                        shop_url = 'https://mall.jd.com' + shop_href
                    else:
                        shop_url = shop_href
                else:
                    shop_url = ''


                ############
                # 尝试从 a 标签中提取
                item_url_elem = item.find('a[href^="//item.jd.com/"]')
                item_url = item_url_elem.attr('href') if item_url_elem else ''

                # 如果没找到，使用 data-sku 拼接
                if not item_url:
                    sku = item.attr('data-sku')
                    item_url = f'https://item.jd.com/{sku}.html' if sku else ''
                elif item_url.startswith('//'):
                    item_url = 'https:' + item_url
                # ==== 新增哈希去重 ====
                item_dict = {'title': title, 'price': price}
                features_url = shop_url or ''  # 用图片链接代替特征链接

                item_hash = self.compute_hash(item_dict, platform, features_url)
                if item_hash in hash_set:
                    # print(f"[跳过] 已存在的商品: {title}")
                    continue
                self.hash_set.add(item_hash)  # 更新哈希池
                # ======================
                # 判断图片链接是否有效并下载到本地
                if not img_url or not isinstance(img_url, str) or not img_url.startswith(('/', '//', 'http')):
                    self.logger.warning(f"[警告] 第 {self.count - 1} 个商品图片 URL 无效，跳过插图")
                    self.logger.warning("商品 HTML 片段：")
                    self.logger.info(item.outer_html())
                    image_path = None
                else:
                    image_path = self.download_image(img_url, self.count - 1) if self.insert_image else None
                self.logger.info(item_url)

                # 包邮信息，评论数量
                post,num_com = self.get_more(item_url)
                # 调用父类的保存方法
                self.save_product_to_excel(
                    row=self.count,
                    title=title,
                    price=price,
                    deal=deal,
                    location=location,
                    shop=shop,
                    post=post,
                    item_url=item_url,
                    shop_url=shop_url,
                    img_url=img_url,
                    num_com=num_com,
                    image_path=image_path
                )
                self.count += 1
                slide_counter += 1  # 累加计数
                if slide_counter % 6 == 0:
                    #print(f"已抓取 {slide_counter} 个商品，调用一次下滑函数...")
                    self.scroll_step_down()
                self.human_sleep(1, 2)

            except Exception as e:
                self.logger.warning(f"[!] 解析商品异常: {e}")
                continue         
        # 保存哈希集，方便下次增量爬取
        if hash_json:
            self.save_hash_set(self.hash_set, hash_json)        
        return True

    def turn_page(self, page_number):
        try:
            # 定位输入框（新版）
            input_box = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, '//div[contains(@class, "_pagination_toPageNum_")]//input')
            ))
            input_box.clear()
            input_box.send_keys(str(page_number))
            input_box.send_keys(Keys.ENTER)  # 按下 Enter 触发跳页

            # 等待当前页码变更为目标页
            self.wait.until(EC.text_to_be_present_in_element(
                (By.XPATH, '//div[contains(@class, "_pagination_item_") and contains(@class, "_active_")]'),
                str(page_number)
            ))

            self.human_sleep(2, 3)
            self.RiskPause(self.captcha_handler.JDslider1())
            self.RiskPause(self.captcha_handler.JDslider2())

        except Exception as e:
            self.logger.error(f"翻页失败，目标页码 {page_number}：{e}")
            self.save_page_html("error_page.html")

    def run(self):
        self.driver.get("https://www.jd.com/")
        # 等待前端登录完成
        self.wait_for_login()
        self.search()

        if self.page_start > 1:
            self.multi_scroll_up_down()
            self.go_to_page(self.page_start)
            # 跳页后等待页面加载

        for page in range(self.page_start, self.page_end + 1):
            # 每页开始前检测停止和暂停
            if self.should_stop():
                self.logger.info("用户选择终止爬虫，提前结束运行。")
                break
            if self.count - 2 >= self.max_items:  # 检查是否已达到最大抓取数量
                self.logger.info(f"已抓取 {self.max_items} 个商品，停止抓取")
                break  # 达到最大数量时停止抓取

            # 如果刚刚跳转页了，这里可能重复滑动，可以根据实际需要调整滑动参数或次数
            self.multi_scroll_up_down(8, 8, -800, 800)

            proceed = self.parse_page(page)
            if not proceed:
                if self.should_stop():
                    self.logger.info(f"用户选择终止爬虫，提前结束运行。")
                    break
                else:
                    self.logger.warning(f"第 {page} 页爬取失败，跳过该页。")
                    continue

            # 当前页不是最后一页，执行翻页
            if page != self.page_end:
                self.turn_page(page + 1)
                # 翻页后等待页面加载
                self.multi_scroll_up_down(8, 8, -800, 800)

        filename = f"{self.keyword}_{time.strftime('%Y%m%d-%H%M')}.xlsx"
        # 使用硬编码路径
        folder = "excel"  # 直接使用硬编码的路径
        if not os.path.exists(folder):
            os.makedirs(folder)
        # 构建完整的文件路径
        filepath = os.path.join(folder, filename)
        self.excel.save(filepath)
        self.logger.info(f"文件已保存: {filename}")
        self.clean_image_folder()
        self.driver.close()
        self.driver.quit()

if __name__ == "__main__":
    keyword = "奉贤黄桃"#input("输入搜索关键词：")
    start_page =1 # int(input("起始页码："))
    end_page = 3 #int(input("终止页码："))
    max_items = 100 #int(input("最多抓取商品数量："))
    insert_image = input("是否插入商品图片到 Excel？(y/n)：").strip().lower() == "y"

    scraper = JDScraper(keyword, start_page, end_page,insert_image, max_items=max_items)
    scraper.run()
