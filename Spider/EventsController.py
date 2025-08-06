# -*- coding: utf-8 -*-

import threading
import logging
from path_utils import static_log_path
import logging


class SpiderEventsController:
    def __init__(self):
        self.logger = self._init_logger()  # 初始化日志
        self._pause_flag = threading.Event()
        self._pause_flag.set()  # 默认运行
        self._stop_flag = threading.Event()




    def _init_logger(self):
        logger = logging.getLogger("Spider")
        logger.setLevel(logging.DEBUG)  # 可调为 INFO 或 ERROR

        log_file_path = static_log_path("Mypider.log")

        file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)

        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)

        # 避免重复添加相同类型的 handler
        if not any(isinstance(h, logging.FileHandler) for h in logger.handlers):
            logger.addHandler(file_handler)
        if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
            logger.addHandler(console_handler)

        return logger

    # ===== 暂停/恢复 =====
    def pause(self):
        self._pause_flag.clear()
        self.logger.info("爬虫已暂停")

    def resume(self):
        self._pause_flag.set()
        self.logger.info("爬虫已恢复")

    def should_pause(self):
        return not self._pause_flag.is_set()

    def wait_if_paused(self):
        while True:
            # 等待1秒看看是否恢复
            if self._pause_flag.wait(timeout=1):
                break  # 恢复了

            if self._stop_flag.is_set():
                self.logger.info("检测到终止信号，立即跳出暂停阻塞")
                break

            self.logger.debug("爬虫暂停中，等待恢复...")

    # ===== 停止爬虫 =====
    def stop(self):
        self._stop_flag.set()
        self._pause_flag.set()  # 确保不会卡在暂停状态
        self.logger.info("爬虫已停止")

    def should_stop(self):
        return self._stop_flag.is_set()

    # ===== 登录等待 =====
    def wait_for_login(self, timeout=300):
        self.logger.info("请手动登录并完成滑块，等待前端点击“继续登录”按钮...")
        self._pause_flag.clear()
        finished = self._pause_flag.wait(timeout=timeout)
        if finished:
            self.logger.info("已收到登录继续信号，继续执行")
        else:
            self.logger.info(f"登录等待超时({timeout}秒)，继续执行")




    
    # ===== 状态查询 =====
    @property
    def is_paused(self):
        return self.should_pause()

