import logging
import os
from datetime import datetime
import sys
from PyQt5.QtCore import pyqtSignal, QObject

class LoggerSignals(QObject):
    """用于在GUI中显示日志的信号类"""
    log_message = pyqtSignal(str, str)  # 日志级别，日志消息

class GUILogger(logging.Logger):
    """扩展标准Logger类，增加GUI集成功能"""
    
    def __init__(self, name, level=logging.NOTSET):
        super().__init__(name, level)
        self.signals = LoggerSignals()
        self._setup_handlers()
    
    def _setup_handlers(self):
        """设置日志处理器"""
        # 创建logs文件夹，如果不存在
        logs_dir = "logs"
        if not os.path.exists(logs_dir):
            os.makedirs(logs_dir)
        
        # 生成日志文件名，基于当前时间
        log_filename = os.path.join(logs_dir, f"gsb_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        
        # 创建文件处理器
        file_handler = logging.FileHandler(log_filename, encoding='utf-8')
        file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_formatter)
        
        # 创建控制台处理器
        console_handler = logging.StreamHandler(sys.stdout)
        console_formatter = logging.Formatter('%(levelname)s: %(message)s')
        console_handler.setFormatter(console_formatter)
        
        # 添加处理器
        self.addHandler(file_handler)
        self.addHandler(console_handler)
        
        # 创建自定义处理器，发送信号到GUI
        self.addHandler(self._create_signal_handler())
    
    def _create_signal_handler(self):
        """创建发送信号到GUI的处理器"""
        class SignalHandler(logging.Handler):
            def __init__(self, signals):
                super().__init__()
                self.signals = signals
            
            def emit(self, record):
                try:
                    level = record.levelname
                    msg = self.format(record)
                    self.signals.log_message.emit(level, msg)
                except Exception:
                    self.handleError(record)
        
        handler = SignalHandler(self.signals)
        handler.setLevel(logging.INFO)  # 只发送INFO及以上级别的日志到GUI
        handler.setFormatter(logging.Formatter('%(message)s'))
        return handler

def setup_logger(name="GSB"):
    """设置并返回日志记录器"""
    # 注册自定义日志类
    logging.setLoggerClass(GUILogger)
    
    # 获取日志记录器
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    return logger 