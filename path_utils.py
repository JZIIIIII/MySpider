# -*- coding: utf-8 -*-
import os
import sys


# 获取项目根目录（以当前文件为基础）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path: str) -> str:
    """
    生成兼容打包与开发环境的资源路径
    """
    try:
        base_path = sys._MEIPASS  # PyInstaller打包后的临时目录
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def ensure_dir(path):
    """
    若路径目录不存在则创建
    """
    if not os.path.exists(path):
        os.makedirs(path)
    return path

def static_excel_path(filename=None):
    folder = os.path.join(BASE_DIR, 'excel')
    ensure_dir(folder)
    return os.path.join(folder, filename) if filename else folder

def static_image_path(subfolder="Name", filename=None):
    """
    构建图片保存路径，例如 images/1688/0001.jpg
    """
    folder = os.path.join(BASE_DIR, 'images', subfolder)
    ensure_dir(folder)
    return os.path.join(folder, filename) if filename else folder

def static_log_path(filename="Mypider.log"):
    folder = os.path.join(BASE_DIR, "logs")
    if not os.path.exists(folder):
        os.makedirs(folder)
    return os.path.join(folder, filename)

def static_hash_path(filename="hash_store.json"):
    folder = os.path.join(BASE_DIR, "data")
    ensure_dir(folder)
    return os.path.join(folder, filename)
