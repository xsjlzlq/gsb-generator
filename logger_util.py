import logging
import os
from datetime import datetime
import sys
from PyQt5.QtCore import pyqtSignal, QObject


class LoggerSignals(QObject):
    """用于在GUI中显示日志的信号类"""
    log_message = pyqtSignal(str, str)  # 日志级别，格式化后的日志消息


class GUILogHandler(logging.Handler):
    """
    全局GUI日志处理器，挂载在根记录器上，
    能够捕获主线程及各子线程（GSB_GUI、GenerateGSB 等）的所有日志并安全发射 Qt 信号。
    """
    def __init__(self):
        super().__init__()
        self.signals = LoggerSignals()
        self.setLevel(logging.INFO)  # GUI 默认显示 INFO 及以上级别的日志
        self.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S'))

    def emit(self, record):
        try:
            msg = self.format(record)
            self.signals.log_message.emit(record.levelname, msg)
        except Exception:
            self.handleError(record)


_global_gui_handler = None


def get_gui_log_handler():
    """获取全局GUI日志处理器单例，并自动挂载到根记录器"""
    global _global_gui_handler
    if _global_gui_handler is None:
        _global_gui_handler = GUILogHandler()
        root_logger = logging.getLogger()
        if root_logger.level > logging.INFO:
            root_logger.setLevel(logging.INFO)
        if _global_gui_handler not in root_logger.handlers:
            root_logger.addHandler(_global_gui_handler)
    return _global_gui_handler


class GUILogger(logging.Logger):
    """扩展标准Logger类（保持兼容旧代码）"""
    def __init__(self, name, level=logging.NOTSET):
        super().__init__(name, level)
        self.signals = get_gui_log_handler().signals


def setup_logger(name="GSB"):
    """设置并返回日志记录器，同时确保全局GUI日志处理器已正确挂载"""
    gui_handler = get_gui_log_handler()

    # 注册自定义日志类
    logging.setLoggerClass(GUILogger)

    # 获取日志记录器
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    logger.signals = gui_handler.signals

    # 设置文件与控制台日志输出（如果尚未配置）
    logs_dir = "logs"
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)

    all_handlers = logger.handlers + logging.getLogger().handlers
    has_file = any(isinstance(h, logging.FileHandler) for h in all_handlers)
    if not has_file:
        log_filename = os.path.join(logs_dir, f"gsb_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        file_handler = logging.FileHandler(log_filename, encoding='utf-8')
        file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

    has_console = any(isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler) for h in all_handlers)
    if not has_console:
        console_handler = logging.StreamHandler(sys.stdout)
        console_formatter = logging.Formatter('%(levelname)s: %(message)s')
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)

    return logger 