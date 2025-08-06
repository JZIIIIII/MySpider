import os
import sys

# 获取项目根目录（以当前文件为基础）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path: str) -> str:
    """
    生成兼容打包与开发环境的资源路径
    """
    try:
        # PyInstaller打包后的临时目录
        base_path = sys._MEIPASS  
    except AttributeError:
        # 未打包时使用当前工作目录
        base_path = os.path.abspath(".")
    
    # 使用 os.path.join 连接路径
    return os.path.join(base_path, relative_path)

def ensure_dir(path):
    """
    若路径目录不存在则创建
    """
    if not os.path.exists(path):
        os.makedirs(path)
    return path

def static_excel_path(filename=None):
    """
    构建 Excel 保存路径
    """
    # 使用相对路径而非硬编码路径
    folder = os.path.join(resource_path('excel'))  # 在打包时获取正确的路径
    ensure_dir(folder)  # 确保目录存在
    return os.path.join(folder, filename) if filename else folder

def static_image_path(subfolder="default", filename=None):
    """
    构建图片保存路径，例如 images/default/0001.jpg
    """
    folder = os.path.join(resource_path('images'), subfolder)
    ensure_dir(folder)  # 确保目录存在
    return os.path.join(folder, filename) if filename else folder

def static_log_path(filename="application.log"):
    """
    构建日志文件保存路径
    """
    folder = os.path.join(resource_path('logs'))
    ensure_dir(folder)  # 确保目录存在
    return os.path.join(folder, filename)

def static_hash_path(filename="data_store.json"):
    """
    构建数据文件路径
    """
    folder = os.path.join(resource_path('data'))
    ensure_dir(folder)  # 确保目录存在
    return os.path.join(folder, filename)
