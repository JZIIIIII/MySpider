from flask import Flask, request, jsonify
import threading
import traceback
import time
import uuid
from CheckAccount import check_login, get_machine_code
from Spider.Tao_collection_tool import TaobaoScraper
from Spider.PDD_collection_tool import PDDScraper
from Spider.JD_collection_tool import JDScraper
from Spider.Ali1688_collection_tool import Ali_1688Scraper
from Spider.ExceLChange import ExcelProcessor

app = Flask(__name__)

# ---------------- 登录状态 ----------------
is_logged_in = False  # 全局登录状态

def require_login(f):
    """装饰器，检查全局登录状态"""
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not is_logged_in:
            return jsonify({"success": False, "msg": "未登录"}), 401
        return f(*args, **kwargs)
    return decorated

# ---------------- 爬虫线程管理 ----------------
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

# ---------------- 登录接口 ----------------
@app.route("/login", methods=["POST"])
def login():
    global is_logged_in
    data = request.get_json()
    account = data.get("account")
    password = data.get("password")
    machine_code = get_machine_code()

    if check_login(account, password):
        is_logged_in = True  # 登录成功，修改全局状态为 True
        return jsonify({"success": True, "msg": "登录成功"})
    else:
        return jsonify({"success": False, "msg": "账号或密码错误"}), 401

# ---------------- 爬虫相关接口 ----------------
@app.route('/start_spider', methods=['POST'])
@require_login
def start_spider():
    data = request.get_json()
    platform = data.get('platform', '').lower()
    keyword = data.get('keyword')
    max_items = int(data.get('max_items', 20))
    insert_image = bool(data.get('insert_image', False))
    start_page = int(data.get('start_page', 1))
    end_page = int(data.get('end_page', 1))
    pierce_span = data.get('pierce_span', [0, 0])
    if not isinstance(pierce_span, list) or len(pierce_span) != 2:
        pierce_span = [0, 0]
    # 保证是数字
    try:
        pierce_span = [float(pierce_span[0]), float(pierce_span[1])]
    except:
        pierce_span = [0, 0]

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

    scraper = scraper_class(
        keyword=keyword,
        max_items=max_items,
        insert_image=insert_image,
        start_page=start_page,
        end_page=end_page,
        pierce_span = pierce_span
    )
    scraper._pause_flag = threading.Event()
    scraper._pause_flag.set()
    scraper._stop_flag = threading.Event()

    scraper_instances[platform] = scraper

    thread = threading.Thread(target=scraper_runner, args=(scraper, platform), daemon=True)
    scraper_threads[platform] = thread
    thread.start()

    return jsonify({"status": "success", "message": f"{platform} 爬虫启动成功"})

@app.route('/get_item_count', methods=['GET'])
@require_login
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

@app.route('/get_spider_status', methods=['GET'])
@require_login
def get_spider_status():
    platform = request.args.get('platform', '').lower()
    scraper = scraper_instances.get(platform)
    if scraper:
        current_status = scraper._pause_flag.is_set()
        status = 1 if current_status else 2
        return jsonify({"status": "success", "platform": platform, "current_status": status})
    return jsonify({"status": "success", "platform": platform, "current_status": 3})

@app.route('/pause_spider', methods=['POST'])
@require_login
def pause_spider():
    platform = request.json.get('platform', '').lower()
    scraper = scraper_instances.get(platform)
    if scraper and hasattr(scraper, '_pause_flag'):
        scraper._pause_flag.clear()
        return jsonify({"status": "success", "message": f"{platform} 已暂停"}), 200
    return jsonify({"status": "error", "message": "该平台爬虫未运行或无法暂停"}), 400

@app.route('/resume_spider', methods=['POST'])
@require_login
def resume_spider():
    platform = request.json.get('platform', '').lower()
    scraper = scraper_instances.get(platform)
    if scraper and hasattr(scraper, '_pause_flag'):
        scraper._pause_flag.set()
        return jsonify({"status": "success", "message": f"{platform} 已继续运行"}), 200
    return jsonify({"status": "error", "message": "该平台爬虫未运行或无法继续"}), 400

@app.route('/stop_spider', methods=['POST'])
@require_login
def stop_spider():
    platform = request.json.get('platform', '').lower()
    scraper = scraper_instances.get(platform)
    if scraper and hasattr(scraper, 'stop'):
        scraper.stop()
        return jsonify({"status": "success", "message": f"{platform} 已发送停止信号"}), 200
    return jsonify({"status": "error", "message": "该平台爬虫未运行或无法停止"}), 400

# ---------------- Excel处理接口 ----------------
@app.route('/deduplication', methods=['POST'])
@require_login
def deduplication():
    file_path = request.form.get('file_path')
    if not file_path:
        return jsonify({"status": "error", "message": "文件路径未提供"}), 400
    processor = ExcelProcessor(file_path)
    success = processor.Deduplication(file_path)
    if success:
        return jsonify({"status": "success", "message": "去重成功", "file_path": file_path}), 200
    else:
        return jsonify({"status": "error", "message": "去重失败"}), 500

@app.route('/screening', methods=['POST'])
@require_login
def screening():
    file_path = request.form.get('file_path')
    tag = request.form.get('tag', "")
    mode = request.form.get('mode', "True") == "True"
    if not file_path:
        return jsonify({"status": "error", "message": "文件路径未提供"}), 400
    processor = ExcelProcessor(file_path)
    success = processor.Screening(file_path, mode, tag)
    if success:
        return jsonify({"status": "success", "message": "筛选成功", "file_path": file_path}), 200
    else:
        return jsonify({"status": "error", "message": "筛选失败"}), 500

@app.route('/')
def index():
    return '爬虫后端服务运行中！'

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
