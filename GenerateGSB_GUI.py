import sys
import os

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


from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                            QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QGroupBox,
                            QTextEdit, QScrollArea, QSplitter, QRadioButton, QButtonGroup, QTabWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor
from GenerateGSB import generate_gsb, generate_mdb_tables_only, generate_jinzhai_mdb_tables_only, generate_dkdc_tables_only, generate_hzzz_sms_only, generate_yhzd_cover_only, generate_jinzhai_gsb_only, generate_yushe_mdb_tables_only, generate_yushe_gsb_only, generate_fbfdc_table_only, generate_yhqd_table_only, generate_yjjsqd_table_only, generate_tsqt_statistics, generate_tsqt_statistics_xlsx, generate_jinzhai_survey_info_tables_only
from logger_util import setup_logger

# 全局日志记录器
logger = setup_logger("GSB_GUI")

class GenerateGSBThread(QThread):
    """处理生成公示表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir, jtjjzz_name):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir
        self.jtjjzz_name = jtjjzz_name

    def run(self):
        try:
            self.progress.emit("开始生成公示表...")
            logger.info(f"开始生成公示表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_gsb(self.mdb_file, self.shp_file, self.output_dir, self.jtjjzz_name)
            logger.info("公示表生成成功")
            self.finished.emit(True, "公示表生成成功！")
        except Exception as e:
            logger.error(f"生成公示表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateMDBThread(QThread):
    """处理生成摸底调查表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成摸底调查表...")
            logger.info(f"开始生成摸底调查表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_mdb_tables_only(self.mdb_file, self.shp_file, self.output_dir)
            logger.info("摸底调查表生成成功")
            self.finished.emit(True, "摸底调查表生成成功！")
        except Exception as e:
            logger.error(f"生成摸底调查表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateJinzhaiThread(QThread):
    """处理生成金寨县摸底调查表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成金寨县摸底调查表...")
            logger.info(f"开始生成金寨县摸底调查表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_jinzhai_mdb_tables_only(self.mdb_file, self.shp_file, self.output_dir)
            logger.info("金寨县摸底调查表生成成功")
            self.finished.emit(True, "金寨县摸底调查表生成成功！")
        except Exception as e:
            logger.error(f"生成金寨县摸底调查表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateDKDCThread(QThread):
    """处理生成承包地块调查表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir, jtjjzz_name):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir
        self.jtjjzz_name = jtjjzz_name

    def run(self):
        try:
            self.progress.emit("开始生成承包地块调查表...")
            logger.info(f"开始生成承包地块调查表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_dkdc_tables_only(self.mdb_file, self.shp_file, self.output_dir, self.jtjjzz_name)
            logger.info("承包地块调查表生成成功")
            self.finished.emit(True, "承包地块调查表生成成功！")
        except Exception as e:
            logger.error(f"生成承包地块调查表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateHZZZThread(QThread):
    """处理生成户主代表声明书的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成户主代表声明书...")
            logger.info(f"开始生成户主代表声明书，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_hzzz_sms_only(self.mdb_file, self.shp_file, self.output_dir)
            logger.info("户主代表声明书生成成功")
            self.finished.emit(True, "户主代表声明书生成成功！")
        except Exception as e:
            logger.error(f"生成户主代表声明书时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateYHZDCoverThread(QThread):
    """处理生成一户一档封面的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成一户一档封面...")
            logger.info(f"开始生成一户一档封面，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_yhzd_cover_only(self.mdb_file, self.shp_file, self.output_dir)
            logger.info("一户一档封面生成成功")
            self.finished.emit(True, "一户一档封面生成成功！")
        except Exception as e:
            logger.error(f"生成一户一档封面时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateJinzhaiGSBThread(QThread):
    """处理生成金寨县公示表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir, jtjjzz_name):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir
        self.jtjjzz_name = jtjjzz_name

    def run(self):
        try:
            self.progress.emit("开始生成金寨县公示表...")
            logger.info(f"开始生成金寨县公示表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_jinzhai_gsb_only(self.mdb_file, self.shp_file, self.output_dir, self.jtjjzz_name)
            logger.info("金寨县公示表生成成功")
            self.finished.emit(True, "金寨县公示表生成成功！")
        except Exception as e:
            logger.error(f"生成金寨县公示表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")


class GenerateJinzhaiSurveyInfoThread(QThread):
    """处理生成金寨县调查信息表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(self, mdb_file, shp_file, output_dir, region_prefix='341524'):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir
        self.region_prefix = region_prefix

    def run(self):
        try:
            self.progress.emit("开始生成金寨县调查信息表...")
            logger.info(f"开始生成金寨县调查信息表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_jinzhai_survey_info_tables_only(self.mdb_file, self.shp_file, self.output_dir, self.region_prefix)
            logger.info("金寨县调查信息表生成成功")
            self.finished.emit(True, "金寨县调查信息表生成成功！")
        except Exception as e:
            logger.error(f"生成金寨县调查信息表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")


class GenerateYusheMDBThread(QThread):
    """处理生成榆社县摸底调查表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成榆社县摸底调查表...")
            logger.info(f"开始生成榆社县摸底调查表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_yushe_mdb_tables_only(self.mdb_file, self.shp_file, self.output_dir)
            logger.info("榆社县摸底调查表生成成功")
            self.finished.emit(True, "榆社县摸底调查表生成成功！")
        except Exception as e:
            logger.error(f"生成榆社县摸底调查表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateYusheGSBThread(QThread):
    """处理生成榆社县公示表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)  # 成功标志和消息
    progress = pyqtSignal(str)  # 进度信息

    def __init__(self, mdb_file, shp_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.shp_file = shp_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成榆社县公示表...")
            logger.info(f"开始生成榆社县公示表，参数：mdb={self.mdb_file}, shp={self.shp_file}, output={self.output_dir}")
            generate_yushe_gsb_only(self.mdb_file, self.shp_file, self.output_dir)
            logger.info("榆社县公示表生成成功")
            self.finished.emit(True, "榆社县公示表生成成功！")
        except Exception as e:
            logger.error(f"生成榆社县公示表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateFBFDCThread(QThread):
    """处理生成发包方调查表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(self, mdb_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成发包方调查表...")
            logger.info(f"开始生成发包方调查表，参数：mdb={self.mdb_file}, output={self.output_dir}")
            generate_fbfdc_table_only(self.mdb_file, self.output_dir)
            logger.info("发包方调查表生成成功")
            self.finished.emit(True, "发包方调查表生成成功！")
        except Exception as e:
            logger.error(f"生成发包方调查表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateYHQDThread(QThread):
    """处理生成一户一档清单的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(self, mdb_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成一户一档清单...")
            logger.info(f"开始生成一户一档清单，参数：mdb={self.mdb_file}, output={self.output_dir}")
            generate_yhqd_table_only(self.mdb_file, self.output_dir)
            logger.info("一户一档清单生成成功")
            self.finished.emit(True, "一户一档清单生成成功！")
        except Exception as e:
            logger.error(f"生成一户一档清单时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateYJJSQDThread(QThread):
    """处理生成移交技术资料清单的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(self, mdb_file, output_dir):
        super().__init__()
        self.mdb_file = mdb_file
        self.output_dir = output_dir

    def run(self):
        try:
            self.progress.emit("开始生成移交技术资料清单...")
            logger.info(f"开始生成移交技术资料清单，参数：mdb={self.mdb_file}, output={self.output_dir}")
            generate_yjjsqd_table_only(self.mdb_file, self.output_dir)
            logger.info("移交技术资料清单生成成功")
            self.finished.emit(True, "移交技术资料清单生成成功！")
        except Exception as e:
            logger.error(f"生成移交技术资料清单时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class GenerateTSQTThread(QThread):
    """处理生成特殊群体统计表的线程，避免GUI冻结"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)

    def __init__(self, survey_dir, output_dir, cover_date):
        super().__init__()
        self.survey_dir = survey_dir
        self.output_dir = output_dir
        self.cover_date = cover_date

    def run(self):
        try:
            self.progress.emit("开始生成特殊群体统计表...")
            logger.info(f"开始生成特殊群体统计表，参数：survey_dir={self.survey_dir}, output={self.output_dir}")
            generate_tsqt_statistics(self.survey_dir, self.output_dir, self.cover_date)
            generate_tsqt_statistics_xlsx(self.survey_dir, self.output_dir)
            logger.info("特殊群体统计表生成成功")
            self.finished.emit(True, "特殊群体统计表生成成功！")
        except Exception as e:
            logger.error(f"生成特殊群体统计表时发生错误: {str(e)}", exc_info=True)
            self.finished.emit(False, f"发生错误: {str(e)}")

class LogDisplay(QTextEdit):
    """日志显示区域"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setLineWrapMode(QTextEdit.NoWrap)
        self.setFont(QFont("Consolas", 9))
        self.log_colors = {
            "DEBUG": QColor(100, 100, 100),
            "INFO": QColor(0, 0, 0),
            "WARNING": QColor(200, 150, 0),
            "ERROR": QColor(200, 0, 0),
            "CRITICAL": QColor(255, 0, 0)
        }
    
    def append_log(self, level, message):
        """添加带颜色的日志消息"""
        color = self.log_colors.get(level, QColor(0, 0, 0))
        self.setTextColor(color)
        self.append(f"{message}")
        # 滚动到底部
        scrollbar = self.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("农村土地承包信息公示表&摸底调查表生成工具")
        self.setMinimumSize(1000, 700)  # 增加最小尺寸
        self.resize(1000, 700)  # 设置初始尺寸
        
        # 连接日志信号
        if hasattr(logger, 'signals'):
            logger.signals.log_message.connect(self.display_log)
        
        self.initUI()

    def initUI(self):
        # 主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 标题
        title_label = QLabel("农村土地二轮承包到期后再延长三十年承包信息公示表&摸底调查表生成工具")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 创建分割器，分为上下两部分
        splitter = QSplitter(Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        
        # 上部分：输入区域
        input_widget = QWidget()
        input_layout = QVBoxLayout(input_widget)
        input_layout.setContentsMargins(0, 0, 0, 0)

        # 功能选择 - 使用标签页分类
        function_group = QGroupBox("功能选择")
        function_layout = QVBoxLayout(function_group)

        # 创建标签页
        self.tab_widget = QTabWidget()
        
        # 南谯区标签页
        nanqiao_tab = QWidget()
        nanqiao_layout = QVBoxLayout(nanqiao_tab)
        
        # 金寨县标签页
        jinzhai_tab = QWidget()
        jinzhai_layout = QVBoxLayout(jinzhai_tab)
        
        # 榆社县标签页
        yushe_tab = QWidget()
        yushe_layout = QVBoxLayout(yushe_tab)
        
        # 明光市标签页
        mingguang_tab = QWidget()
        mingguang_layout = QVBoxLayout(mingguang_tab)
        
        # 创建功能按钮组
        self.function_button_group = QButtonGroup()
        
        # 南谯区功能
        self.generate_gsb_radio = QRadioButton("生成公示表")
        self.generate_mdb_radio = QRadioButton("生成摸底调查表")
        self.generate_dkdc_radio = QRadioButton("生成承包地块调查表")
        self.generate_hzzz_radio = QRadioButton("生成户主代表声明书")
        self.generate_yhzd_radio = QRadioButton("生成一户一档封面")
        
        # 金寨县功能
        self.generate_jinzhai_radio = QRadioButton("生成摸底调查表")
        self.generate_jinzhai_gsb_radio = QRadioButton("生成公示表")
        self.generate_jinzhai_survey_info_radio = QRadioButton("生成土地承包经营权调查信息表")
        
        # 榆社县功能
        self.generate_yushe_radio = QRadioButton("生成摸底调查表")
        self.generate_yushe_gsb_radio = QRadioButton("生成公示表")
        
        # 明光市功能
        self.generate_fbfdc_radio = QRadioButton("生成发包方调查表")
        self.generate_yhqd_radio = QRadioButton("生成一户一档清单")
        self.generate_yjjsqd_radio = QRadioButton("生成移交技术资料清单")
        self.generate_tsqt_radio = QRadioButton("生成特殊群体统计表")

        # 默认选择南谯区的公示表
        self.generate_gsb_radio.setChecked(True)
        
        # 添加按钮到按钮组
        self.function_button_group.addButton(self.generate_gsb_radio, 1)
        self.function_button_group.addButton(self.generate_mdb_radio, 2)
        self.function_button_group.addButton(self.generate_jinzhai_radio, 3)
        self.function_button_group.addButton(self.generate_jinzhai_gsb_radio, 4)
        self.function_button_group.addButton(self.generate_jinzhai_survey_info_radio, 13)
        self.function_button_group.addButton(self.generate_dkdc_radio, 5)
        self.function_button_group.addButton(self.generate_hzzz_radio, 6)
        self.function_button_group.addButton(self.generate_yhzd_radio, 7)
        self.function_button_group.addButton(self.generate_yushe_radio, 8)
        self.function_button_group.addButton(self.generate_yushe_gsb_radio, 14)
        self.function_button_group.addButton(self.generate_fbfdc_radio, 9)
        self.function_button_group.addButton(self.generate_yhqd_radio, 10)
        self.function_button_group.addButton(self.generate_yjjsqd_radio, 11)
        self.function_button_group.addButton(self.generate_tsqt_radio, 12)
        
        # 添加按钮到对应标签页
        nanqiao_layout.addWidget(self.generate_gsb_radio)
        nanqiao_layout.addWidget(self.generate_mdb_radio)
        nanqiao_layout.addWidget(self.generate_dkdc_radio)
        nanqiao_layout.addWidget(self.generate_hzzz_radio)
        nanqiao_layout.addWidget(self.generate_yhzd_radio)
        
        jinzhai_layout.addWidget(self.generate_jinzhai_radio)
        jinzhai_layout.addWidget(self.generate_jinzhai_gsb_radio)
        jinzhai_layout.addWidget(self.generate_jinzhai_survey_info_radio)
        
        yushe_layout.addWidget(self.generate_yushe_radio)
        yushe_layout.addWidget(self.generate_yushe_gsb_radio)
        
        # 明光市布局
        mingguang_layout.addWidget(self.generate_fbfdc_radio)
        mingguang_layout.addWidget(self.generate_yhqd_radio)
        mingguang_layout.addWidget(self.generate_yjjsqd_radio)
        mingguang_layout.addWidget(self.generate_tsqt_radio)
        
        # 添加标签页到标签部件
        self.tab_widget.addTab(nanqiao_tab, "南谯区")
        self.tab_widget.addTab(jinzhai_tab, "金寨县")
        self.tab_widget.addTab(yushe_tab, "榆社县")
        self.tab_widget.addTab(mingguang_tab, "明光市")
        
        function_layout.addWidget(self.tab_widget)
        input_layout.addWidget(function_group)

        # 连接功能选择信号
        self.function_button_group.buttonToggled.connect(self.toggle_function)

        # 集体经济组织名称
        self.org_group = QGroupBox("发包方信息")
        org_layout = QVBoxLayout(self.org_group)

        jtjjzz_layout = QHBoxLayout()
        jtjjzz_label = QLabel("集体经济组织名称:")
        self.jtjjzz_input = QLineEdit()
        self.jtjjzz_input.setPlaceholderText("请输入发包方名称（如XX村民委员会）")
        jtjjzz_layout.addWidget(jtjjzz_label)
        jtjjzz_layout.addWidget(self.jtjjzz_input, 1)
        org_layout.addLayout(jtjjzz_layout)

        input_layout.addWidget(self.org_group)

        # 特殊群体统计表参数
        self.tsqt_group = QGroupBox("特殊群体统计表参数")
        tsqt_layout = QVBoxLayout(self.tsqt_group)

        survey_dir_layout = QHBoxLayout()
        survey_dir_label = QLabel("摸底表目录:")
        self.survey_dir_input = QLineEdit()
        self.survey_dir_input.setReadOnly(True)
        self.survey_dir_input.setPlaceholderText("请选择家庭成员摸底信息表所在目录")
        survey_dir_button = QPushButton("浏览...")
        survey_dir_button.clicked.connect(self.select_survey_dir)
        survey_dir_layout.addWidget(survey_dir_label)
        survey_dir_layout.addWidget(self.survey_dir_input, 1)
        survey_dir_layout.addWidget(survey_dir_button)
        tsqt_layout.addLayout(survey_dir_layout)

        tsqt_date_layout = QHBoxLayout()
        tsqt_date_label = QLabel("封面日期:")
        self.tsqt_date_input = QLineEdit()
        self.tsqt_date_input.setPlaceholderText("    年  月  日")
        from datetime import date
        today = date.today()
        self.tsqt_date_input.setText(f"{today.year}年{today.month}月{today.day}日")
        tsqt_date_layout.addWidget(tsqt_date_label)
        tsqt_date_layout.addWidget(self.tsqt_date_input, 1)
        tsqt_layout.addLayout(tsqt_date_layout)

        input_layout.addWidget(self.tsqt_group)
        self.tsqt_group.hide()

        # 文件选择区域
        file_group = QGroupBox("数据文件选择")
        file_layout = QVBoxLayout(file_group)

        # MDB文件选择
        mdb_layout = QHBoxLayout()
        mdb_label = QLabel("MDB数据库文件:")
        self.mdb_input = QLineEdit()
        self.mdb_input.setReadOnly(True)
        self.mdb_input.setPlaceholderText("请选择MDB数据库文件")
        mdb_button = QPushButton("浏览...")
        mdb_button.clicked.connect(self.select_mdb_file)
        mdb_layout.addWidget(mdb_label)
        mdb_layout.addWidget(self.mdb_input, 1)
        mdb_layout.addWidget(mdb_button)
        file_layout.addLayout(mdb_layout)

        # SHP文件选择
        shp_layout = QHBoxLayout()
        shp_label = QLabel("SHP矢量文件路径:")
        self.shp_input = QLineEdit()
        self.shp_input.setReadOnly(True)
        self.shp_input.setPlaceholderText("请选择SHP矢量文件所在的文件夹")
        shp_button = QPushButton("浏览...")
        shp_button.clicked.connect(self.select_shp_dir)
        shp_layout.addWidget(shp_label)
        shp_layout.addWidget(self.shp_input, 1)
        shp_layout.addWidget(shp_button)
        file_layout.addLayout(shp_layout)

        # 输出目录选择
        output_layout = QHBoxLayout()
        output_label = QLabel("输出目录:")
        self.output_input = QLineEdit()
        self.output_input.setReadOnly(True)
        self.output_input.setPlaceholderText("请选择输出目录")
        output_button = QPushButton("浏览...")
        output_button.clicked.connect(self.select_output_dir)
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_input, 1)
        output_layout.addWidget(output_button)
        file_layout.addLayout(output_layout)

        input_layout.addWidget(file_group)

        # 状态显示
        status_group = QGroupBox("处理状态")
        status_layout = QVBoxLayout(status_group)
        self.status_label = QLabel("就绪，请填写信息后点击生成按钮")
        self.status_label.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.status_label)
        input_layout.addWidget(status_group)

        # 生成按钮
        self.generate_button = QPushButton("生成公示表")
        self.generate_button.setMinimumHeight(40)
        self.generate_button.clicked.connect(self.generate_gsb)
        input_layout.addWidget(self.generate_button)
        
        # 下部分：日志区域
        log_group = QGroupBox("日志信息")
        log_layout = QVBoxLayout(log_group)
        
        self.log_display = LogDisplay()
        log_layout.addWidget(self.log_display)
        
        # 添加到分割器
        splitter.addWidget(input_widget)
        splitter.addWidget(log_group)
        
        # 设置初始大小比例
        splitter.setSizes([400, 200])
        
        # 添加分割器到主布局
        main_layout.addWidget(splitter)

        # 底部版权信息
        footer_label = QLabel("© 2025 农村土地承包信息管理系统")
        footer_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(footer_label)
        
        # 初始化功能状态（在所有控件创建完成后调用）
        self.toggle_function(self.generate_gsb_radio, True)

        self.current_function = "gsb"  # 默认是公示表

        # 记录日志
        logger.info("应用程序界面初始化完成")

    def select_mdb_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择MDB数据库文件", "", "Access数据库 (*.mdb *.accdb)"
        )
        if file_path:
            self.mdb_input.setText(file_path)
            logger.info(f"已选择MDB数据库文件: {file_path}")

    def select_shp_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择SHP矢量文件所在的文件夹")
        if dir_path:
            self.shp_input.setText(dir_path)
            logger.info(f"已选择SHP矢量文件路径: {dir_path}")

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_input.setText(dir_path)
            logger.info(f"已选择输出目录: {dir_path}")

    def select_survey_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择家庭成员摸底信息表所在目录")
        if dir_path:
            self.survey_dir_input.setText(dir_path)
            logger.info(f"已选择摸底表目录: {dir_path}")

    def generate_gsb(self):
        # 验证输入
        jtjjzz_name = ""
        if self.current_function in ["gsb", "dkdc", "jinzhai_gsb"]:
            jtjjzz_name = self.jtjjzz_input.text().strip()
            if not jtjjzz_name:
                QMessageBox.warning(self, "输入错误", "请输入集体经济组织名称")
                logger.warning("未输入集体经济组织名称")
                return
        elif self.current_function == "yhzd":
            # 一户一档封面不需要集体经济组织名称
            pass
        elif self.current_function == "fbfdc":
            # 发包方调查表不需要集体经济组织名称
            pass
        elif self.current_function == "yhqd":
            # 一户一档清单不需要集体经济组织名称
            pass
        elif self.current_function == "yjjsqd":
            # 移交技术资料清单不需要集体经济组织名称
            pass
        elif self.current_function == "tsqt":
            # 特殊群体统计表不需要集体经济组织名称
            pass

        mdb_file = self.mdb_input.text().strip()
        shp_file = self.shp_input.text().strip()
        output_dir = self.output_input.text().strip()

        if self.current_function != "tsqt":
            if not mdb_file:
                QMessageBox.warning(self, "输入错误", "请选择MDB数据库文件")
                logger.warning("未选择MDB数据库文件")
                return
            if self.current_function not in ("fbfdc", "yhqd", "yjjsqd") and not shp_file:
                QMessageBox.warning(self, "输入错误", "请选择SHP矢量文件所在的文件夹")
                logger.warning("未选择SHP矢量文件所在的文件夹")
                return
        if not output_dir:
            QMessageBox.warning(self, "输入错误", "请选择输出目录")
            logger.warning("未选择输出目录")
            return

        if self.current_function == "tsqt":
            survey_dir = self.survey_dir_input.text().strip()
            if not survey_dir:
                QMessageBox.warning(self, "输入错误", "请选择摸底表目录")
                logger.warning("未选择摸底表目录")
                return
            if not os.path.exists(survey_dir):
                QMessageBox.warning(self, "文件错误", "摸底表目录不存在")
                logger.error(f"摸底表目录不存在: {survey_dir}")
                return
            cover_date = self.tsqt_date_input.text().strip()
            if not cover_date:
                cover_date = "    年  月  日"

        # 验证文件是否存在
        if self.current_function != "tsqt" and not os.path.exists(mdb_file):
            QMessageBox.warning(self, "文件错误", "MDB数据库文件不存在")
            logger.error(f"MDB数据库文件不存在: {mdb_file}")
            return
        if self.current_function not in ("fbfdc", "yhqd", "yjjsqd", "tsqt") and not os.path.exists(shp_file):
            QMessageBox.warning(self, "文件错误", "SHP矢量文件所在的文件夹不存在")
            logger.error(f"SHP矢量文件所在的文件夹不存在: {shp_file}")
            return
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                logger.info(f"已创建输出目录: {output_dir}")
            except Exception as e:
                QMessageBox.critical(self, "目录错误", f"无法创建输出目录: {str(e)}")
                logger.error(f"无法创建输出目录: {output_dir}, 错误: {str(e)}")
                return

        # 根据选择的模式创建相应的线程
        if self.current_function == "gsb":
            self.status_label.setText("正在生成公示表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成公示表需要较长时间，请耐心等待。\n生成的公示表文件将保存在指定的输出目录中。")
            logger.info("开始生成公示表")
        elif self.current_function == "mdb":
            self.status_label.setText("正在生成摸底调查表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成摸底调查表需要较长时间，请耐心等待。\n生成的摸底调查表文件将保存在指定的输出目录中。")
            logger.info("开始生成摸底调查表")
        elif self.current_function == "jinzhai":
            self.status_label.setText("正在生成金寨县摸底调查表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成金寨县摸底调查表需要较长时间，请耐心等待。\n生成的金寨县摸底调查表文件将保存在指定的输出目录中。")
            logger.info("开始生成金寨县摸底调查表")
        elif self.current_function == "jinzhai_gsb":
            self.status_label.setText("正在生成金寨县公示表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成金寨县公示表需要较长时间，请耐心等待。\n生成的金寨县公示表文件将保存在指定的输出目录中。")
            logger.info("开始生成金寨县公示表")
        elif self.current_function == "jinzhai_survey_info":
            self.status_label.setText("正在生成金寨县调查信息表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成金寨县调查信息表需要较长时间，请耐心等待。\n生成的金寨县调查信息表文件将保存在指定的输出目录中。")
            logger.info("开始生成金寨县调查信息表")
        elif self.current_function == "dkdc":
            self.status_label.setText("正在生成承包地块调查表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成承包地块调查表需要较长时间，请耐心等待。\n生成的承包地块调查表文件将保存在指定的输出目录中。")
            logger.info("开始生成承包地块调查表")
        elif self.current_function == "hzzz":
            self.status_label.setText("正在生成户主代表声明书，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成户主代表声明书需要较长时间，请耐心等待。\n生成的户主代表声明书文件将保存在指定的输出目录中。")
            logger.info("开始生成户主代表声明书")
        elif self.current_function == "yhzd":
            self.status_label.setText("正在生成一户一档封面，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成一户一档封面需要较长时间，请耐心等待。\n生成的一户一档封面文件将保存在指定的输出目录中。")
            logger.info("开始生成一户一档封面")
        elif self.current_function == "yushe":
            self.status_label.setText("正在生成榆社县摸底调查表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成榆社县摸底调查表需要较长时间，请耐心等待。\n生成的榆社县摸底调查表文件将保存在指定的输出目录中。")
            logger.info("开始生成榆社县摸底调查表")
        elif self.current_function == "yushe_gsb":
            self.status_label.setText("正在生成榆社县公示表，请稍候...")
            QMessageBox.information(self, "功能提示",
                                  "生成榆社县公示表需要较长时间，请耐心等待。\n生成的榆社县公示表文件将保存在指定的输出目录中。")
            logger.info("开始生成榆社县公示表")
        elif self.current_function == "fbfdc":
            self.status_label.setText("正在生成发包方调查表，请稍候...")
            logger.info("开始生成发包方调查表")
        elif self.current_function == "yhqd":
            self.status_label.setText("正在生成一户一档清单，请稍候...")
            logger.info("开始生成一户一档清单")
        elif self.current_function == "yjjsqd":
            self.status_label.setText("正在生成移交技术资料清单，请稍候...")
            logger.info("开始生成移交技术资料清单")
        elif self.current_function == "tsqt":
            self.status_label.setText("正在生成特殊群体统计表，请稍候...")
            logger.info("开始生成特殊群体统计表")

        QApplication.processEvents()

        # 创建并启动工作线程
        if self.current_function == "gsb":
            self.worker_thread = GenerateGSBThread(mdb_file, shp_file, output_dir, jtjjzz_name)
        elif self.current_function == "mdb":
            self.worker_thread = GenerateMDBThread(mdb_file, shp_file, output_dir)
        elif self.current_function == "jinzhai":
            self.worker_thread = GenerateJinzhaiThread(mdb_file, shp_file, output_dir)
        elif self.current_function == "jinzhai_gsb":
            self.worker_thread = GenerateJinzhaiGSBThread(mdb_file, shp_file, output_dir, jtjjzz_name)
        elif self.current_function == "jinzhai_survey_info":
            self.worker_thread = GenerateJinzhaiSurveyInfoThread(mdb_file, shp_file, output_dir)
        elif self.current_function == "dkdc":
            self.worker_thread = GenerateDKDCThread(mdb_file, shp_file, output_dir, jtjjzz_name)
        elif self.current_function == "hzzz":
            self.worker_thread = GenerateHZZZThread(mdb_file, shp_file, output_dir)
        elif self.current_function == "yhzd":
            self.worker_thread = GenerateYHZDCoverThread(mdb_file, shp_file, output_dir)
        elif self.current_function == "yushe":
            self.worker_thread = GenerateYusheMDBThread(mdb_file, shp_file, output_dir)
        elif self.current_function == "yushe_gsb":
            self.worker_thread = GenerateYusheGSBThread(mdb_file, shp_file, output_dir)
        elif self.current_function == "fbfdc":
            self.worker_thread = GenerateFBFDCThread(mdb_file, output_dir)
        elif self.current_function == "yhqd":
            self.worker_thread = GenerateYHQDThread(mdb_file, output_dir)
        elif self.current_function == "yjjsqd":
            self.worker_thread = GenerateYJJSQDThread(mdb_file, output_dir)
        elif self.current_function == "tsqt":
            self.worker_thread = GenerateTSQTThread(survey_dir, output_dir, cover_date)

        self.worker_thread.progress.connect(self.update_progress)
        self.worker_thread.finished.connect(self.process_finished)
        self.worker_thread.start()

        function_name = {"gsb": "公示表", "mdb": "摸底调查表", "jinzhai": "金寨县摸底调查表", "jinzhai_gsb": "金寨县公示表", "jinzhai_survey_info": "金寨县调查信息表", "dkdc": "承包地块调查表", "hzzz": "户主代表声明书", "yhzd": "一户一档封面", "yushe": "榆社县摸底调查表", "yushe_gsb": "榆社县公示表", "fbfdc": "发包方调查表", "yhqd": "一户一档清单", "yjjsqd": "移交技术资料清单", "tsqt": "特殊群体统计表"}.get(self.current_function, "")
        logger.info(f"已启动生成{function_name}线程")

    def update_progress(self, message):
        self.status_label.setText(message)
        logger.info(f"进度更新: {message}")

    def process_finished(self, success, message):
        function_name = {"gsb": "公示表", "mdb": "摸底调查表", "jinzhai": "金寨县摸底调查表", "jinzhai_gsb": "金寨县公示表", "jinzhai_survey_info": "金寨县调查信息表", "dkdc": "承包地块调查表", "hzzz": "户主代表声明书", "yhzd": "一户一档封面", "yushe": "榆社县摸底调查表", "yushe_gsb": "榆社县公示表", "fbfdc": "发包方调查表", "yhqd": "一户一档清单", "yjjsqd": "移交技术资料清单", "tsqt": "特殊群体统计表"}.get(self.current_function, "")
        if success:
            QMessageBox.information(self, "处理完成", message)
            self.status_label.setText("处理完成")
            logger.info(f"{function_name}生成成功完成")
        else:
            QMessageBox.critical(self, "处理失败", message)
            self.status_label.setText("处理失败")
            logger.error(f"{function_name}生成失败: {message}")
    
    def toggle_function(self, button, checked):
        """切换功能选择"""
        if not checked:
            return  # 只在选中时处理

        if button == self.generate_gsb_radio:
            self.current_function = "gsb"
            self.jtjjzz_input.setEnabled(True)
            self.jtjjzz_input.show()
            self.generate_button.setText("生成公示表")
            self.status_label.setText("就绪，请填写信息后点击生成公示表按钮")
            self.org_group.show()
        elif button == self.generate_mdb_radio:
            self.current_function = "mdb"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成摸底调查表")
            self.status_label.setText("就绪，请填写信息后点击生成摸底调查表按钮")
            self.org_group.hide()
        elif button == self.generate_jinzhai_radio:
            self.current_function = "jinzhai"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成摸底调查表_金寨县")
            self.status_label.setText("就绪，请填写信息后点击生成摸底调查表_金寨县按钮")
            self.org_group.hide()
        elif button == self.generate_jinzhai_gsb_radio:
            self.current_function = "jinzhai_gsb"
            self.jtjjzz_input.setEnabled(True)
            self.jtjjzz_input.show()
            self.generate_button.setText("生成公示表_金寨县")
            self.status_label.setText("就绪，请填写信息后点击生成公示表_金寨县按钮")
            self.org_group.show()
        elif button == self.generate_jinzhai_survey_info_radio:
            self.current_function = "jinzhai_survey_info"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成调查信息表_金寨县")
            self.status_label.setText("就绪，请填写信息后点击生成调查信息表_金寨县按钮")
            self.org_group.hide()
        elif button == self.generate_dkdc_radio:
            self.current_function = "dkdc"
            self.jtjjzz_input.setEnabled(True)
            self.jtjjzz_input.show()
            self.generate_button.setText("生成承包地块调查表")
            self.status_label.setText("就绪，请填写信息后点击生成承包地块调查表按钮")
            self.org_group.show()
        elif button == self.generate_hzzz_radio:
            self.current_function = "hzzz"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成户主代表声明书")
            self.status_label.setText("就绪，请填写信息后点击生成户主代表声明书按钮")
            self.org_group.hide()
        elif button == self.generate_yhzd_radio:
            self.current_function = "yhzd"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成一户一档封面")
            self.status_label.setText("就绪，请填写信息后点击生成一户一档封面按钮")
            self.org_group.hide()
        elif button == self.generate_yushe_radio:
            self.current_function = "yushe"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成摸底调查表_榆社县")
            self.status_label.setText("就绪，请填写信息后点击生成摸底调查表_榆社县按钮")
            self.org_group.hide()
        elif button == self.generate_yushe_gsb_radio:
            self.current_function = "yushe_gsb"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成公示表_榆社县")
            self.status_label.setText("就绪，请填写信息后点击生成公示表_榆社县按钮")
            self.org_group.hide()
        elif button == self.generate_fbfdc_radio:
            self.current_function = "fbfdc"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成发包方调查表")
            self.status_label.setText("就绪，请填写信息后点击生成发包方调查表按钮")
            self.org_group.hide()
        elif button == self.generate_yhqd_radio:
            self.current_function = "yhqd"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成一户一档清单")
            self.status_label.setText("就绪，请填写信息后点击生成一户一档清单按钮")
            self.org_group.hide()
        elif button == self.generate_yjjsqd_radio:
            self.current_function = "yjjsqd"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成移交技术资料清单")
            self.status_label.setText("就绪，请填写信息后点击生成移交技术资料清单按钮")
            self.org_group.hide()
        elif button == self.generate_tsqt_radio:
            self.current_function = "tsqt"
            self.jtjjzz_input.setEnabled(False)
            self.jtjjzz_input.hide()
            self.generate_button.setText("生成特殊群体统计表")
            self.status_label.setText("就绪，请填写信息后点击生成特殊群体统计表按钮")
            self.org_group.hide()
            self.tsqt_group.show()
        # 隐藏tsqt_group如果切换到非tsqt功能
        if button != self.generate_tsqt_radio and hasattr(self, 'tsqt_group'):
            self.tsqt_group.hide()
        logger.info(f"功能切换到: {self.current_function}")

    def display_log(self, level, message):
        """接收日志信号并显示"""
        self.log_display.append_log(level, message)
        
if __name__ == "__main__":
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 