import os
import sys
import logging
import traceback
from datetime import datetime

# 修复numpy和pickle循环导入问题
try:
    # 预加载numpy相关模块以避免循环导入
    import numpy
    import numpy.core
    import numpy.random
    # 强制加载_pickle模块
    import numpy.random._pickle
except ImportError:
    pass

# 修复pandas相关的导入问题
try:
    # 预加载pickle模块
    import pickle
    import _pickle
except ImportError:
    pass

# 配置日志记录
def setup_logging():
    # 创建logs文件夹，如果不存在
    logs_dir = "logs"
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)
    
    # 生成日志文件名，基于当前时间
    log_filename = os.path.join(logs_dir, f"gsb_gui_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    # 配置日志记录器
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return logging.getLogger("GSB_GUI")

# 异常处理函数
def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        # 如果是键盘中断，使用系统默认处理
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    
    # 记录未捕获的异常
    logger.critical("未捕获的异常:", exc_info=(exc_type, exc_value, exc_traceback))
    
    # 在Windows上显示错误消息
    import ctypes
    error_msg = f"程序发生错误:\n{exc_type.__name__}: {exc_value}\n\n详细错误信息已记录到日志文件。"
    ctypes.windll.user32.MessageBoxW(0, error_msg, "公示表生成工具错误", 0x10)

if __name__ == "__main__":
    # 设置日志记录
    logger = setup_logging()
    logger.info("应用程序启动")
    
    # 设置未捕获异常处理器
    sys.excepthook = handle_exception
    
    try:
        # 导入GUI模块
        from GenerateGSB_GUI import QApplication, MainWindow
        
        # 启动应用程序
        logger.info("初始化GUI")
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        
        # 运行应用程序
        logger.info("GUI显示完成，进入事件循环")
        exit_code = app.exec_()
        logger.info(f"应用程序退出，退出码: {exit_code}")
        sys.exit(exit_code)
    
    except Exception as e:
        logger.critical(f"启动应用程序时发生错误: {str(e)}")
        logger.critical(traceback.format_exc())
        
        # 在Windows上显示错误消息
        import ctypes
        error_msg = f"启动应用程序时发生错误:\n{str(e)}\n\n详细错误信息已记录到日志文件。"
        ctypes.windll.user32.MessageBoxW(0, error_msg, "公示表生成工具错误", 0x10)
        sys.exit(1) 