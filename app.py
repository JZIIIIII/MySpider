from flask import Flask, request, jsonify
import threading
import traceback
import time
import os
import subprocess
import json
import tkinter as tk
from tkinter import filedialog
from werkzeug.utils import secure_filename
from Spider.Tao_collection_tool import TaobaoScraper
from Spider.PDD_collection_tool import PDDScraper
from Spider.JD_collection_tool import JDScraper
from Spider.Ali1688_collection_tool import Ali_1688Scraper
from Spider.ExceLChange import ExcelProcessor
from LicenceRecode import LicenseValidator  
# 配置文件路径
config_file_path = 'config.json'

app = Flask(__name__)

# 全局存放线程和爬虫实例的映射
scraper_threads = {}
scraper_instances = {}

def scraper_runner(scraper, platform):
    try:
        scraper.run()
    except Exception as e:
        traceback.print_exc()
    finally:
        scraper._stop_flag.set()
        time.sleep(1)
        scraper_threads.pop(platform, None)
        scraper_instances.pop(platform, None)

def mask_file_path(file_path):
    """对文件路径进行脱敏处理，只显示路径的最后部分"""
    if file_path:
        return os.path.join("...", file_path.split(os.sep)[-1])
    return file_path

def mask_keyword(keyword):
    """对关键字进行脱敏处理，显示部分关键字"""
    if keyword:
        return keyword[:3] + "****" if len(keyword) > 3 else keyword
    return keyword

@app.route('/start_spider', methods=['POST'])
def start_spider():
    data = request.get_json()
    platform = data.get('platform', '').lower()
    keyword = data.get('keyword')
    max_items = int(data.get('max_items', 20))
    insert_image = bool(data.get('insert_image', False))
    start_page = int(data.get('start_page', 1))
    end_page = int(data.get('end_page', 1))

    if not platform or not keyword:
        return jsonify({"status": "error", "message": "缺少平台或关键词"}), 400
    if platform not in ['taobao', 'pdd', 'jd', '1688']:
        return jsonify({"status": "error", "message": "不支持的平台"}), 400
    if platform in scraper_threads:
        return jsonify({"status": "error", "message": "该平台的爬虫正在运行中"}), 400

    SCRAPER_CLASS_MAP = {
        'taobao': TaobaoScraper,
        'pdd': PDDScraper,
        'jd': JDScraper,
        '1688': Ali_1688Scraper
    }
    scraper_class = SCRAPER_CLASS_MAP.get(platform)
    if scraper_class is None:
        return jsonify({"status": "error", "message": "找不到爬虫类"}), 500

    # 实例化爬虫
    scraper = scraper_class(
        keyword=keyword,
        max_items=max_items,
        insert_image=insert_image,
        start_page=start_page,
        end_page=end_page
    )
    scraper._pause_flag = threading.Event()
    scraper._pause_flag.set()
    scraper._stop_flag = threading.Event()

    scraper_instances[platform] = scraper

    thread = threading.Thread(target=scraper_runner, args=(scraper, platform), daemon=True)
    scraper_threads[platform] = thread
    thread.start()

    return jsonify({"status": "success", "message": f"{platform} 爬虫启动成功，关键词：{mask_keyword(keyword)}"})


@app.route('/get_item_count', methods=['GET'])
def get_item_count():
    platform = request.args.get('platform', '').lower()
    scraper = scraper_instances.get(platform)
    
    if scraper:
        current_item_count = scraper.count
        return jsonify({
            "status": "success",
            "platform": platform,
            "item_count": current_item_count
        })
    
    return jsonify({"status": "error", "message": "该平台爬虫未启动"}), 400


@app.route('/deduplication', methods=['POST'])
def deduplication():
    file_path = request.form.get('file_path')
    if not file_path:
        return jsonify({"status": "error", "message": "文件路径未提供"}), 400

    processor = ExcelProcessor(file_path)
    success = processor.Deduplication(file_path)
    
    if success:
        return jsonify({"status": "success", "message": "去重成功", "file_path": mask_file_path(file_path)}), 200
    else:
        return jsonify({"status": "error", "message": "去重失败"}), 500


@app.route('/screening', methods=['POST'])
def screening():
    file = request.files.get('file')
    file_path = request.form.get('file_path')
    tag = request.form.get('tag', "")  
    mode = request.form.get('mode', "True") == "True"  

    if not file_path:
        return jsonify({"status": "error", "message": "文件路径未提供"}), 400

    processor = ExcelProcessor(file_path)
    success = processor.Screening(file_path, mode, tag)

    if success:
        return jsonify({"status": "success", "message": "筛选成功", "file_path": mask_file_path(file_path)}), 200
    else:
        return jsonify({"status": "error", "message": "筛选失败"}), 500

@app.route('/get_spider_status', methods=['GET'])
def get_spider_status():
    platform = request.args.get('platform', '').lower()
    scraper = scraper_instances.get(platform)
    if scraper:
        current_status = scraper._pause_flag.is_set()
        status = 1 if current_status else 2
        return jsonify({
            "status": "success",
            "platform": platform,
            "current_status": status
        })
    
    return jsonify({
        "status": "success",
        "platform": platform,
        "current_status": 3
    })


@app.route('/select_license_file', methods=['POST'])
def select_license_file():
    root = tk.Tk()
    root.withdraw()
    license_file_path = filedialog.askopenfilename(title="选择许可证文件", filetypes=[("License Files", "*.lic")])
    return jsonify({"license_path": mask_file_path(license_file_path)})


if __name__ == '__main__':
    license_path = get_license_path()

    if not license_path:
        license_path = select_license_file()

        if license_path:
            save_license_path(license_path)
        else:
            sys.exit(1)

    if check_license(license_path):
        app.run(debug=True, host='0.0.0.0', port=5000)
    else:
        pass
