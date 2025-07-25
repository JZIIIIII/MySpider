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
from BaseScraper import BaseScraper

from AntiScrapingException import CaptchaHandler
from path_utils import static_image_path ,static_hash_path , static_excel_path




class JDScraper(BaseScraper):
    def __init__(self, keyword, start_page, end_page, max_items=100):
        #super().__init__(headless=None, proxy=None)
        self.keyword = keyword
        self.page_start = start_page
        self.page_end = end_page
        self.max_items = max_items
        self.logger = self._init_logger()
        self.driver = self.init_driver()
        self.wait = WebDriverWait(self.driver, 10)
        self.excel = Workbook()
        self.sheet = self.excel.active
        self.count = 2
        self.insert_image = True
        self._setup_excel()
        #self.lock = threading.Lock()  # 用于同步的锁
        self.anti_spider_triggered = False
        self.captcha_handler = CaptchaHandler(self.driver)
        self.captcha_handler.JDslider1()  # 调用风控

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
            self.captcha_handler.JDslider1()
        except Exception as e:
            self.logger.warning(f"搜索操作失败:{e}")

    def go_to_page(self, page_number):

        self.captcha_handler.JDslider1(1,2)
        try:
            # 定位到输入框并输入页数
            search_input = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, '//span[@class="p-skip"]//input[@class="input-txt"]')
            ))
            search_input.clear()  # 清空输入框
            search_input.send_keys(page_number)  

            # 定位到确认按钮并点击
            confirm_button = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, '//span[@class="p-skip"]//a[@class="btn btn-default"]')
            ))
            confirm_button.click()
            # 模拟人工休眠
            self.captcha_handler.JDslider1(1,2) 
        except Exception as e:
            # 记录错误日志
            self.logger.warning(f"翻页失败: {e}")
            self.save_page_html("error_page.html")

    def get_more (self, url):
        post_text = ''
        main_window = self.driver.current_window_handle

        try:
            # 提取商品ID，防止href匹配失败
            item_id = re.search(r'/(\d+)\.html', url)
            if not item_id:
                self.logger.warning("无法提取商品ID，跳过")
                return 0
            item_id = item_id.group(1)

            # 用模糊匹配定位链接
            link_xpath = f'//a[contains(@href, "{item_id}")]'
            link_element = self.wait.until(EC.element_to_be_clickable((By.XPATH, link_xpath)))

            # Ctrl+点击打开新标签页
            ActionChains(self.driver).key_down(Keys.CONTROL).click(link_element).key_up(Keys.CONTROL).perform()

            self.captcha_handler.JDslider1(1,2) 


            # 等待新标签页打开
            for _ in range(10):
                handles = self.driver.window_handles
                if len(handles) > 1:
                    break
                time.sleep(0.5)
            else:
                self.logger.warning("新标签页未打开")
                self.save_page_html("error_page.html")

                return 0

            # 重新获取所有窗口句柄
            handles = self.driver.window_handles
            new_tabs = [h for h in handles if h != main_window]
            if not new_tabs:
                self.logger.warning("找不到新标签页句柄")
                return 0

            new_tab = new_tabs[-1]
            self.driver.switch_to.window(new_tab)

            self.captcha_handler.JDslider1() 


            # 等待页面刷新
            try:
                # 等待包邮标签出现
                self.wait.until(EC.presence_of_element_located(
                    (By.XPATH, '//span[contains(@class, "free-shipping")]')
                ))

                # 获取包邮文字
                post_text = self.driver.find_element(
                    By.XPATH,
                    '//span[contains(@class, "free-shipping")]'
                ).text.strip()

            except Exception:
                post_text = ''

            return post_text  # 例如返回：'包邮'

        except Exception as e:
            self.logger.warning("获取包邮失败")

        finally:
            # 关闭新标签页，切换回主窗口
            self.human_sleep(1,2)
            try:
                if self.driver.current_window_handle != main_window:
                    self.driver.close()
                    self.driver.switch_to.window(main_window)
                    self.captcha_handler.JDslider1(1,2) 
            except Exception as e:
                self.logger.warning("关闭标签页或切换主窗口失败")
                self.save_page_html("error_page.html")
                

        return post_text

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

            # 创建文件夹
            folder = "images/JD"
            if not os.path.exists(folder):
                os.makedirs(folder)

            # 统一使用 static_image_path 构造保存路径
            file_name = f"{index:04d}.jpg"
            file_path = static_image_path("JD", file_name)

            # 保存成 jpg 文件
            img.save(file_path, format="JPEG")

            # 返回文件路径
            return file_path

        except Exception as e:
            self.logger.error(f"图片下载异常: {e} | URL: {url}")
            return None

    def clean_image_folder(self):
        folder = static_image_path("JD")  # 获取图片根目录，不传filename只拿文件夹路径
        if os.path.exists(folder):
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # 删除文件或快捷方式
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # 删除子文件夹




    def parse_page(self, page_number, platform='JD',hash_json='hash_store.json'):
        html = self.driver.page_source
        doc = pq(html)
        items = doc('ul.gl-warp.clearfix > li').items()
        slide_counter = 0  # 每页开始前初始化下滑计数器
        # 载入历史哈希
        hash_json = static_hash_path("hash_store.json")
        hash_set = self.load_hash_set(hash_json)


        for item in items:
            try:
                if self.count - 2 >= self.max_items:
                    print(f"已达到最大抓取数量：{self.max_items}，停止抓取。")
                    return False

                if self.anti_spider_triggered == True:
                    print(f"触发强制风控账号冻结 应前往京东APP解封")
                    return False

    #            if item.find('.title--RoseSo8H').text() or item.find('.headTitleText--hxVemljn').text():
    #                continue

                # 获取图片标签
                img_tag = item.find('img')

                # 处理懒加载情况，优先选择 data-lazy-img 或者 src
                img_url = img_tag.attr('data-lazy-img') if img_tag.attr('data-lazy-img') else img_tag.attr('src')

                # 判断是否为相对路径，如果是，补全为完整的 URL
                if img_url and img_url.startswith('//'):
                    img_url = 'https:' + img_url  # 补全相对路径为完整 URL

                # 打印或保存图片链接
                print(img_url)

                # 依次尝试 src / data-src / src2 / data-ks-lazyload
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

                # 判断图片链接是否有效并下载到本地
                if not img_url or not isinstance(img_url, str) or not img_url.startswith(('/', '//', 'http')):
                    self.logger.warning(f"[警告] 第 {self.count - 1} 个商品图片 URL 无效，跳过插图")
                    self.logger.warning("商品 HTML 片段：")
                    print(item.outer_html())
                    image_path = None
                else:
                    image_path = self.download_image(img_url, self.count - 1) if self.insert_image else None

                #标题
                title_elem = item.find('.p-name em')
                title = title_elem.text().strip() if title_elem else ''
                #价格
                price_elem = item.find('.p-price i')
                price_text = price_elem.text().strip() if price_elem else ''          
                try:
                    price = float(price_text)
                except ValueError:
                    price = 0.0

                #京东不显示销量
                deal = 0
                #无地址
                location = ''
                #店铺名称和链接
                shop_elem = item.find('.p-shop a')
                shop_name = shop_elem.text().strip() if shop_elem else ''
                shop = shop_name
                shop_href = shop_elem.attr('href') if shop_elem else ''
                shop_url = 'https:' + shop_href if shop_href and shop_href.startswith('//') else shop_href
                #评论数量
                num_com = item.find('.p-commit a')
                num_com = num_com.text().strip() if num_com else ''
                #商品链接
                item_url = item.find('a[href^="//item.jd.com/"]')
                item_url = item_url.attr('href') if item_url else ''
                item_url = 'https:' + item_url if item_url and item_url.startswith('//') else item_url
                # ==== 新增哈希去重 ====
                item_dict = {'title': title, 'price': price}
                features_url = shop_url or ''  # 用图片链接代替特征链接

                item_hash = self.compute_hash(item_dict, platform, features_url)
                if item_hash in hash_set:
                    # print(f"[跳过] 已存在的商品: {title}")
                    continue
                hash_set.add(item_hash)
                # ======================
                post = self.get_more(item_url)
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
        self.save_hash_set(hash_set, hash_json)            
        return True

    def turn_page(self, page_number):
        try:
            # 找到页码输入框
            input_box = self.wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, '#J_bottomPage .p-skip .input-txt')
            ))
            input_box.clear()
            input_box.send_keys(str(page_number))

            # 点击“确定”按钮
            confirm_btn = self.driver.find_element(
                By.CSS_SELECTOR, '#J_bottomPage .p-skip .btn'
            )
            confirm_btn.click()
            self.captcha_handler.JDslider1(1,2) 

            # 等待当前页码变更为目标页
            self.wait.until(EC.text_to_be_present_in_element(
                (By.CSS_SELECTOR, '#J_bottomPage .p-num .curr'), str(page_number)
            ))

            # 停顿防止频繁操作
            self.human_sleep(2, 3)

        except Exception as e:
            self.logger.error(f"翻页失败，目标页码 {page_number}：{e}")
            self.save_page_html("error_page.html")

    def run(self):
        self.driver.get("https://www.jd.com/")
        input("请在浏览器中手动登录京东，如出现滑块，请手动完成验证后按 Enter 继续......")
        self.search()

        if self.page_start != 1:
            self.multi_scroll_up_down()
            self.go_to_page(self.page_start)

        for page in range(self.page_start, self.page_end + 1):
            if self.count - 2 >= self.max_items:  # 检查是否已达到最大抓取数量
                print(f"已抓取 {self.max_items} 个商品，停止抓取")
                break  # 达到最大数量时停止抓取

            self.multi_scroll_up_down(8,8,-800,800)
            proceed = self.parse_page(page)
            if not proceed:
                self.logger.warning(print(f"第 {page} 页爬取失败，跳过"))
                continue  # 如果该页抓取失败，跳过

            if page != self.page_end:
                self.turn_page(page + 1)

        filename = f"{self.keyword}_{time.strftime('%Y%m%d-%H%M')}.xlsx"
        filepath = static_excel_path(filename)
        self.excel.save(filepath)
        print(f"文件已保存: {filename}")
        self.clean_image_folder()
        self.driver.quit()

if __name__ == "__main__":
    keyword = "奉贤黄桃"#input("输入搜索关键词：")
    start_page =1 # int(input("起始页码："))
    end_page = 3 #int(input("终止页码："))
    max_items = 100 #int(input("最多抓取商品数量："))
    insert_image = input("是否插入商品图片到 Excel？(y/n)：").strip().lower() == "y"

    scraper = JDScraper(keyword, start_page, end_page, max_items=max_items)
    scraper.insert_image = insert_image
    scraper.run()

