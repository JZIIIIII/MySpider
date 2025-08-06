# -*- coding: utf-8 -*-
import time
import random
import re
import requests
import os
import shutil

from io import BytesIO
from PIL import Image
from PIL import UnidentifiedImageError
from tkinter import filedialog

from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
from seleniumwire import webdriver  # selenium-wire 用于拦截请求
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from pyquery import PyQuery as pq
from openpyxl import Workbook
from Spider.BaseScraper  import BaseScraper
from Spider.AntiScrapingException import CaptchaHandler
from path_utils import static_image_path ,static_hash_path , static_excel_path


class PDDScraper(BaseScraper):
    def __init__(self, keyword, start_page, end_page, max_items=100, insert_image=True):
        super().__init__(headless=True, proxy=None)
        self.keyword = keyword
        self.page_start = start_page
        self.page_end = end_page
        self.max_items = max_items
        self.insert_image = insert_image
        self.count = 2
        self.wait = WebDriverWait(self.driver, 10)
        self.excel = Workbook()
        self.sheet = self.excel.active
        self._setup_excel()
        self.empty_data_count = 0
        self.anti_spider_triggered = False
        self.captcha_handler = CaptchaHandler(self.driver, self.logger)
        self.captcha_handler.PDDsliderl()  # 调用风控

    def _setup_excel(self):
        # 根据修改后的 save_to_excel 调整表头顺序
        headers = ['Num', 'Title', 'Price', 'Deal', 'shop_url', 'CommentNum', 'ShopName', 'Postage', 'Tags', 'Image']
        for i, header in enumerate(headers, 1):
            self.sheet.cell(row=1, column=i, value=header)

    def simulate_click(self, element):
        try:
            actions = ActionChains(self.driver)
            actions.move_to_element(element).click().perform()
            self.logger.info("模拟鼠标点击成功")
            return True
        except Exception as e:
            self.logger.error(f"模拟鼠标点击失败:{e}")
            self.save_page_html("error_page.html")

            return False

    def click_fake_search_box(self, timeout=10):
        wait = WebDriverWait(self.driver, timeout)
        try:
            time.sleep(2)  # 等待页面加载

            # 等待目标元素可点击
            element = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "_18v23kPu")))

            # 调用封装好的模拟点击函数
            return self.simulate_click(element)
            self.RiskPause(self.captcha_handler.PDDsliderl())
        except Exception as e:
            self.logger.error(f"搜索框模拟点击失败:{e}")
            self.save_page_html("error_page.html")

            return False

    def scroll_step_down(self, base_step=800):
        """模拟人类向下较大幅度滑动"""
        step = random.randint(base_step - 300, base_step + 300)
        self.driver.execute_script(f"window.scrollBy(0, {step});")
        time.sleep(random.uniform(0.8, 1.5))  # 适当增加等待时间，保证加载

    def search(self):
        try:
            # 先尝试定位真实搜索框
            self.logger.info("尝试定位真实搜索框...")
            real_search_box = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[@type='search' and contains(@class, '_2bfwu6WT')]")
            ))
        except Exception:

            self.logger.info("请尝试点击首页搜索框进入搜索页面...")
            return  # 找不到搜索框，结束函数，等待用户点击假搜索框进入搜索页

        try:
            # 定位搜索按钮
            search_btn = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class, 'RuSDrtii') and text()='搜索']")
            ))
            # 输入关键词
            self.logger.info(f"输入关键词: {self.keyword}")
            real_search_box.clear()
            real_search_box.send_keys(self.keyword)

            # 等待 1~2 秒，模拟人操作
            self.human_sleep(1, 2)

            # 用你的 simulate_click 函数点击搜索按钮
            self.simulate_click(search_btn)

            # 等待 1~2 秒，等待搜索结果加载
            self.human_sleep(1, 2)

            self.logger.info("搜索成功")

        except Exception as e:
            self.logger.error(f"搜索失败:{e}")
            self.save_page_html("error_page.html")

    def parse_all_showcases(self, max_items=10 ,platform='PDD',hash_json=None): 
        results = []
        seen_titles = set()
        processed_count = 0
        # 载入历史哈希
        if hash_json:
            hash_json = static_hash_path(hash_json)
            hash_set = self.load_hash_set(hash_json)
        else:
            hash_set = set()
        previous_results_count = 0  # 记录每次下滑前的商品数

        for _ in range(30):  # 最多下滑 30 次
            items = self.driver.find_elements(By.CSS_SELECTOR, 'div.rjNMXsUm._1unt3Js-')
            for index in range(len(items)):
                if len(results) >= max_items:

                    break

                self.RiskPause(self.captcha_handler.PDDsliderl())
                self.RiskPause(self.anti_spider_triggered)
                # 如果暂停，则等待恢复
                self.wait_if_paused()
                # 检查是否被强制终止
                if self.should_stop():
                    self.logger.info("检测到提前终止命令，保存已抓取内容并退出 parse_page。")
                    if hash_json:
                        self.save_hash_set(hash_set, hash_json)  
                    return results


                try:

                    items = self.driver.find_elements(By.CSS_SELECTOR, 'div.rjNMXsUm._1unt3Js-')
                    if index >= len(items):
                        continue

                    elem = items[index]
                    item_html = elem.get_attribute('outerHTML')
                    item_doc = pq(item_html)

                    self.RiskPause(self.captcha_handler.PDDsliderl())
                    title = item_doc('div._3ANzdjkc').text().strip()
                    if not title or title in seen_titles:
                        continue
                    seen_titles.add(title)

                    price = ''
                    price_dec_elem = item_doc('span._2aP8LGPL')
                    if price_dec_elem:
                        price_int = price_dec_elem.prev().text()
                        price_dec = price_dec_elem.text()
                        price = f"{price_int}{price_dec}"
                    elif not price:
                        price_elem = item_doc('span._3_U04GgA')
                        if price_elem:
                            coupon_tag = price_elem('span._2nKWnaqa')
                            if coupon_tag:
                                price = price_elem('span._3f_Cp5GQ + span').text()
                            else:
                                price = price_elem('span._3f_Cp5GQ + span').text()


                    tag_list = [tag.text() for tag in item_doc('div._299OVZvt > div').items()]
                    tags_str = ' '.join(tag_list)

                    img_elem = item_doc('div._1o7l_Qm- img')
                    img_url = img_elem.attr('src') or img_elem.attr('data-src') or img_elem.attr('data-lazy')

                    deal_num_elem = item_doc('div[style*="width: 255px;"] span')
                    deal_num_text = deal_num_elem.text().strip()
                    deal_num = ''.join([ch for ch in deal_num_text if ch.isdigit()])
                    if not deal_num:
                        deal_num = 0

                    # ==== 新增哈希去重 ====
                    item_dict = {'title': title, 'price': price}
                    features_url = img_url or ''  # 用图片链接代替特征链接

                    item_hash = self.compute_hash(item_dict, platform, features_url)
                    if item_hash in hash_set:
                        self.logger.info(f"[跳过] 已存在的商品: {title}")
                        continue
                    hash_set.add(item_hash)
                    # ======================

                    # 只有通过哈希检查才调用 get_more，节省性能

                    comment_count, shop_name, postage_info, shop_url = self.get_more(elem)
                    if comment_count is None:
                        comment_count = 0

                    img_path = None
                    if self.insert_image and img_url:
                        img_path = self.download_image(img_url, title)

                    result = {
                        'title': title,
                        'price': price,
                        'deal_num': deal_num,
                        'shop_url': shop_url,
                        'comment_count': comment_count,
                        'shop_name': shop_name,
                        'postage_info': postage_info,
                        'tags': tags_str,
                        'img_url': img_url,
                        'img_path': img_path
                    }
                    results.append(result)
                    processed_count += 1
                    self.count = processed_count + 2
                    self.logger.info (self.count)
                    if processed_count % 2 == 0:
                        self.scroll_step_down(base_step=800)
                        time.sleep(1)

                except Exception as e:
                    self.logger.warning(f"[!] 处理商品失败: {e}")
                    continue

            if len(results) >= max_items:
                break


            if len(results) == previous_results_count:
                self.logger.info(f"[!] 商品数量没有增加，停止继续下滑。")
                break
            previous_results_count = len(results)

            self.scroll_step_down(base_step=1200)
            time.sleep(1)

            # 在每次下滑后检查是否触发风控
            if getattr(self, 'anti_spider_triggered', False):
                self.logger.warning("[!] 触发风控，暂停爬虫等待用户处理")
                self.pause()
                self.wait_if_paused()
                if self.should_stop():
                    self.logger.warning("[!] 用户选择终止，提前退出任务")
                    if hash_json:
                        self.save_hash_set(hash_set, hash_json)
                    return results
                else:
                    self.logger.info("[*] 用户选择继续，重置风控状态")
                    self.anti_spider_triggered = False
                
        self.logger.info(f"\n共提取到 {len(results)} 个商品（上限：{max_items}）")
        # 保存哈希集，方便下次增量爬取
        if hash_json:
            self.save_hash_set(hash_set, hash_json)  

        return results

    def save_to_excel(self, results, filename='results.xlsx'):
        wb = Workbook()
        ws = wb.active
        ws.title = '商品信息'

        # 表头
        headers = ['标题', '价格', '成交量', '店铺连接', '评论数量', '店铺名称', '包邮信息', '标签', '图片']
        ws.append(headers)

        # 设置图片插入列列宽（第9列）
        img_col_letter = 'I'
        ws.column_dimensions[img_col_letter].width = 15

        for i, item in enumerate(results, start=2):
            ws.cell(row=i, column=1, value=item['title'])
            ws.cell(row=i, column=2, value=item['price'])
            ws.cell(row=i, column=3, value=item['deal_num'])
            ws.cell(row=i, column=4, value=str(item['shop_url']))
            ws.cell(row=i, column=5, value=item['comment_count'])
            ws.cell(row=i, column=6, value=item['shop_name'])
            ws.cell(row=i, column=7, value=item['postage_info'])
            ws.cell(row=i, column=8, value=item['tags'])

            # 如果开启插图功能
            if self.insert_image:
                img_path = item.get('img_path')
                if img_path and os.path.exists(img_path):
                    try:
                        # 尝试用 Pillow 打开并重新编码为 JPEG（修复部分 JPEG 插入失败问题）
                        with Image.open(img_path) as img:
                            rgb_img = img.convert("RGB")
                            rgb_img.save(img_path, format='JPEG')  # 覆盖原图
                    
                        # 插入图片
                        excel_img = ExcelImage(img_path)
                        excel_img.width = 80
                        excel_img.height = 80
                        ws.row_dimensions[i].height = 60
                        excel_img.anchor = f'{img_col_letter}{i}'
                        ws.add_image(excel_img)
                    except Exception as e:
                        self.logger.error(f"[!] 图片插入失败: {img_path} | 错误: {e}")
                else:
                    self.logger.warning(f"[!] 图片路径无效或文件不存在: {img_path}")
            
        # 保存 Excel 文件
        wb.save(filename)

        self.logger.info(f"[√] 数据已保存到 Excel：{filename}")

    def webp_to_png(self, webp_bytes, save_path):
        try:
            img = Image.open(BytesIO(webp_bytes))
            if img.format != 'WEBP':
                self.logger.warning(f"警告：不是WEBP格式，实际是 {img.format}")
            img = img.convert("RGBA")  # 保留透明度
            img.save(save_path, format="PNG")
            self.logger.info(f"WEBP图片转换PNG并保存成功: {save_path}")
            return save_path
        except Exception as e:
            self.logger.warning(f"WEBP转PNG失败: {e}")
            return None

    def download_image(self, img_url, title, platform='PDD'):
        """
        下载图片并保存到 images/<platform>/ 目录下，文件名由标题生成。
        若为 WebP 格式将自动转为 PNG，其余转为 JPG。
        """
        # 清理标题中的特殊字符，避免路径错误
        filename = re.sub(r'[\\/:*?"<>|]', '_', title[:20].strip().replace(' ', '_')) + ".jpg"

        # 硬编码路径，直接指定保存路径
        folder = f"images/{platform}"  # 使用硬编码的路径
        if not os.path.exists(folder):
            os.makedirs(folder)  # 如果目录不存在，则创建

        path = os.path.join(folder, filename)  # 生成完整路径

        # 拼多多常见图片链接修复
        if img_url.startswith("//"):
            img_url = "https:" + img_url
        elif img_url.startswith("/"):
            img_url = "https://img.pddpic.com" + img_url

        img_url = img_url.split('?')[0]  # 去掉查询参数

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
            "Referer": "https://mobile.pinduoduo.com/",
            "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
        }

        try:
            resp = requests.get(img_url, headers=headers, timeout=10)
            if resp.status_code != 200:
                self.logger.error(f"图片下载失败，状态码: {resp.status_code} | URL: {img_url}")
                return None

            try:
                img = Image.open(BytesIO(resp.content))

                if img.format == 'WEBP':
                    self.logger.info(f"[!] 图片为WEBP，转换为PNG: {img_url}")
                    png_path = os.path.splitext(path)[0] + ".png"
                    result = self.webp_to_png(resp.content, png_path)
                    return result

                img = img.convert('RGB')
                img.save(path, format='JPEG')
                self.logger.info(f"图片已保存: {path}")
                return path

            except UnidentifiedImageError:
                self.logger.error(f"[!] 图片格式不支持或路径无效: {img_url}")
                return None
            except Exception as e:
                self.logger.error(f"[!] 图片处理异常: {e} | URL: {img_url}")
                return None

        except Exception as e:
            self.logger.error(f"[!] 图片下载异常: {e} | URL: {img_url}")
            return None

    def clean_url(self, url):
        """
        使用正则表达式清理 URL，保留 goods_id 参数，去掉其他多余的参数。
        """
        try:
            # 使用正则匹配并提取 base_url 和 goods_id 参数
            match = re.search(r"(https?://[^\?]+)(\?[^#]*)", url)
            if match:
                base_url = match.group(1)  # 基础 URL
                query_string = match.group(2)  # 查询部分

                # 解析查询参数
                query_params = parse_qs(query_string[1:])  # 去掉 '?'
                
                # 只保留 'goods_id' 参数
                cleaned_params = {key: value for key, value in query_params.items() if key == 'goods_id'}
                
                # 构造新的查询字符串
                cleaned_query = urlencode(cleaned_params, doseq=True)
                
                # 如果有有效的 cleaned_query，则重建新的 URL
                if cleaned_query:
                    cleaned_url = f"{base_url}?{cleaned_query}"
                else:
                    cleaned_url = base_url  # 如果没有有效的参数，返回没有参数的 URL
                
                return cleaned_url
            else:
                return url  # 如果没有匹配，返回原始 URL
        except Exception as e:
            self.logger.error(f"Error cleaning URL: {e}")
            return url

    def get_more(self, element):
        comment_count = 0
        shop_name = ''
        postage_info = ''
        shop_url = ''  # 新增用于存储店铺URL

        try:
            # 先滚动到元素中间，避免被遮挡
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.5)
            '''
            # 等待遮挡元素消失（如果知道遮挡的class，替换下）
            try:
                self.wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.浮动遮挡层的class')))
            except TimeoutException:
                pass  # 或打印日志
            '''
            # 直接用 JS 点击，绕过遮挡
            self.driver.execute_script("arguments[0].click();", element)
        
            # 获取当前页面 URL
            current_url = self.driver.current_url
            shop_url = self.clean_url(current_url)
            self.RiskPause(self.captcha_handler.PDDsliderl())
            self.wait_if_paused()
            if self.should_stop():
                self.logger.info(f"用户选择终止爬虫，提前结束运行。")
                self.stop()  # 立即停止爬虫
                return 0 , None, None, None   # 返回 0，表示终止操作
        
            # 等待详情页关键元素加载
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.F2MXl7Xc')))
            except Exception:
                self.logger.warning("[!] 无法加载评论数量元素，跳过")
                comment_count = 0  # 默认评论数为0

            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.BAq4Lzv7')))
            except Exception:
                self.logger.warning("[!] 无法加载店铺名元素，跳过")
                shop_name = "无店铺名称"  # 默认店铺名称

            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.RbQ7MTuU')))
            except Exception:
                self.logger.warning("[!] 无法加载包邮信息元素，跳过")
                postage_info = "无包邮信息"  # 默认包邮信息

            # 获取评论数（如果没有评论，这一步会失败）
            try:
                comment_text_element = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.F2MXl7Xc'))
                )
                comment_text = comment_text_element.text
                match = re.search(r'\((\d+)\)', comment_text)
                if match:
                    comment_count = int(match.group(1))
            except Exception:
                comment_count = 0  # 如果没有评论或无法加载评论，设置为 0

            # 获取店铺名称
            try:
                shop_name_element = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.BAq4Lzv7'))
                )
                shop_name = shop_name_element.text.strip()
            except Exception as e:
                self.logger.warning(f"[!] 获取店铺名称失败: {repr(e)}")
                shop_name = "无店铺名称"  # 设置默认值


            # 获取包邮状况
            try:
                postage_elements = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.RbQ7MTuU span.KDFIGUNK'))
                )
                postage_info = ' '.join([e.text.strip() for e in postage_elements if e.text.strip()])
            except Exception as e:
                self.logger.warning(f"[!] 获取包邮信息失败: {repr(e)}")
                postage_info = "无包邮信息"  # 设置默认值

        except Exception as e:
            self.logger.error(f"[!] 获取详情数据失败: {repr(e)}")
            self.save_page_html("error_page.html")


        finally:
            self.driver.back()
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.rjNMXsUm._1unt3Js-')))
            except TimeoutException:
                time.sleep(1)
            self.human_sleep(1, 2)

        if (
            comment_count == 0 and
            shop_name in ["", "无店铺名称"] and
            postage_info in ["", "无包邮信息"] and
            shop_url == ''
        ):
            self.empty_data_count = getattr(self, 'empty_data_count', 0) + 1
            self.logger.warning(f"[!] 获取为空数据 {self.empty_data_count}/3 次")
            if self.empty_data_count >= 3:
                self.logger.error("[!] 连续 3 次获取为空，程序即将暂停")
                self.anti_spider_triggered = True  #  设置风控标志位
                return 0 , None, None, None      #  提前返回特殊值
        else:

            self.empty_data_count = 0  # 有效数据则复位
        
        return comment_count, shop_name, postage_info, shop_url  # 返回店铺URL

    def clear_image_cache(self, subfolder='PDD'):
        """
        清空指定平台的图片缓存目录，默认是 images/PDD
        """
        # 使用硬编码的路径，不再调用 static_image_path
        folder = f"images/{subfolder}"

        if os.path.exists(folder) and os.path.isdir(folder):
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    self.logger.warning(f"[!] 删除缓存文件失败: {file_path} | 错误: {e}")
        else:
            self.logger.error(f"[!] 文件夹不存在或不是目录：{folder}")

    def run(self):
        try:
            self.driver.get("https://mobile.pinduoduo.com/")
            self.logger.info("登录拼多多成功")
        except:
            self.logger.error("登录拼多多失败")

        self.wait_for_login()

        self.click_fake_search_box()
        self.search()
        self.RiskPause(self.captcha_handler.PDDsliderl())
        data = self.parse_all_showcases(max_items=self.max_items)

        if data:
            # 使用硬编码路径
            filename = f"{self.keyword}_{time.strftime('%Y%m%d_%H%M')}.xlsx"
            folder = "excel"  # 硬编码路径
            if not os.path.exists(folder):
                os.makedirs(folder)  # 如果目录不存在，则创建

            # 生成完整的文件路径
            filepath = os.path.join(folder, filename)

            self.save_to_excel(data, filepath)  # 保存数据到 Excel
        self.clear_image_cache()  # 清理图片缓存
        self.driver.close()
        self.driver.quit()



if __name__ == "__main__":
    kw = input("输入关键词：")
    sp = 1 #int(input("起始页码："))
    ep = 1 #int(input("结束页码："))
    mi = int(input("最大商品数："))
    show_img = input("是否插入图片 (y/n)：").strip().lower() == 'y'

    spider = PDDScraper(kw, sp, ep, mi, show_img)
    spider.run()

