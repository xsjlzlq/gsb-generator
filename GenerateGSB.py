#### 导入mdb和矢量DK数据生成公示信息表格####

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

# 导入SHP编码处理模块
import sys
import os
from shp_encoding_handler import setup_chinese_encoding, read_shp_safe, detect_encoding

# 初始化中文编码环境
setup_chinese_encoding()

import pandas as pd
import geopandas as gpd
import os
import tempfile
import pyodbc
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter
import fnmatch
import logging
import sys
import traceback
import shutil
import warnings
import zipfile
from copy import copy

# 导入python-docx用于Word文档操作
try:
    from docx import Document
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("Warning: python-docx not installed.户主代表声明书功能将不可用。")
    Document = None


# 导入xlsxwriter用于替代实现
try:
    import xlsxwriter
except ImportError:
    print("Warning: xlsxwriter not installed. Using openpyxl as fallback for generate_mdb_table1.")
    xlsxwriter = None

# 设置默认日志记录器
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def generate_gsb(mdb_file_path, shp_file_path, output_directory,jtjjzz_name):
    logger.debug(f"mdb_file_path: {mdb_file_path}")
    logger.debug(f"shp_file_path: {shp_file_path}")
    logger.debug(f"output_directory: {output_directory}")
    logger.debug(f"jtjjzz_name: {jtjjzz_name}")
    # 确保输出目录存在
    os.makedirs(output_directory, exist_ok=True)

    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

    # 创建MDB连接
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + mdb_file_path + ';'
    )
    logger.debug(f"conn_str: {conn_str}")
    conn = pyodbc.connect(conn_str)

    try:
        # 读取MDB文件中的表
        tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
        # 抑制pandas的SQLAlchemy连接警告
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}
        logger.debug(f"Tables read successfully: {tables}")

        # 处理CBF_JTCY表和CBF表
        cbf_jtcy = dfs['CBF_JTCY']
       
        cbf = dfs['CBF']
         # 假设'date_column'是包含日期的列
        cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
        cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
        # 处理CBHT表，提取CBQXQ和CBQXZ字段
        cbht = dfs['CBHT']
        cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
        cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
        # 选择所需字段
        cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]
        
        # 合并CBF_JTCY和CBF表，保留CBF表中的CBFBM字段
        merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left', suffixes=('', '_original'))
        logger.debug(f"Merged CBF_JTCY and CBF successfully")
        
        # 合并CBHT表
        merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
        logger.debug(f"Merged with CBHT successfully")

        # 用CBFBM的前14位挂接FBF表的FBFBM字段
        merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
        fbf = dfs['FBF']
        fbf['FBFDCRQ'] = pd.to_datetime(fbf['FBFDCRQ'], errors='coerce')
        merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')
        logger.debug(f"Merged with FBF successfully")

        # 处理CBDKXX表和矢量地块数据
        cbdxx = dfs['CBDKXX']
        # 查找名称中包含"DK"的shp文件
        shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
        logger.debug(f"Found SHP files: {shp_files_with_dk}")

        # 初始化一个空的GeoDataFrame来存储所有符合条件的矢量地块数据
        all_vector_data = gpd.GeoDataFrame()

        for shp_file in shp_files_with_dk:
            try:
                # 使用改进的SHP编码处理模块
                logger.debug(f"开始读取SHP文件: {shp_file}")
                vector_data = read_shp_safe(shp_file)
            except Exception as e:
                logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                raise
            all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)
        logger.debug(f"Read and concatenated SHP files successfully")

        try:
            # 使用改进的SHP编码处理模块
            logger.debug(f"开始读取主SHP文件: {shp_file_path}")
            vector_data = read_shp_safe(shp_file_path)
        except Exception as e:
            logger.error(f"读取主SHP文件失败 {shp_file_path}: {str(e)}")
            raise
        merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')
        logger.debug(f"Merged CBDKXX and vector data successfully")

        # 将数据写入两个Excel表格
        with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
            merged_all.to_excel(writer, sheet_name='Sheet1', index=False)
        logger.debug(f"Saved merged_all to Excel file: {excel_file1}")

        with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
            # 将GeoDataFrame转换为DataFrame以便写入Excel
            merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
            merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)
        logger.debug(f"Saved merged_cbdxx_vector to Excel file: {excel_file2}")

        merge_excel_files(excel_file1, excel_file2, output_directory, jtjjzz_name)
        logger.debug(f"Merge Excel files completed")

    except Exception as e:
        logger.error(f"发生错误: {e}")
        logger.error(traceback.format_exc())
        raise

    finally:
        # 关闭数据库连接
        conn.close()
        logger.debug("Database connection closed")

def find_shp_files_with_dk(shp_file_path):
    """
    在指定路径下查找名称中包含"DK"的shp文件。
    
    参数:
    shp_file_path (str): shp文件所在的目录路径。
    
    返回:
    list: 包含"DK"的shp文件的完整路径列表。
    """
    shp_files_with_dk = []
    for root, dirs, files in os.walk(shp_file_path):
        for filename in fnmatch.filter(files, '*.shp'):
            if "DK" in filename:
                shp_files_with_dk.append(os.path.join(root, filename))
    return shp_files_with_dk


def yhzgx_code_to_name(code):
    """
    将 YHZGX 代码转换为对应的名称。
    """
    code_name_map = {
        '01': "本人",
        '02': "户主",
        '10': "配偶",
        '11': "夫",
        '12': "妻",
        '20': "子",
        '21': "独生子",
        '22': "长子",
        '23': "次子",
        '24': "三子",
        '25': "四子",
        '26': "五子",
        '27': "养子或继子",
        '28': "女婿",
        '29': "其他儿子",
        '30': "女",
        '31': "独生女",
        '32': "长女",
        '33': "次女",
        '34': "三女",
        '35': "四女",
        '36': "五女",
        '37': "养女或继女",
        '38': "儿媳",
        '39': "其他女儿",
        '40': "孙子、孙女或外孙子、外孙女",
        '41': "孙子",
        '42': "孙女",
        '43': "外孙子",
        '44': "外孙女",
        '45': "孙媳妇或外孙媳妇",
        '46': "孙女婿或外孙女婿",
        '47': "曾孙子或外曾孙子",
        '48': "曾孙女或外曾孙女",
        '49': "其他孙子、孙女或外孙子、外孙女",
        '50': "父母",
        '51': "父亲",
        '52': "母亲",
        '53': "公公",
        '54': "婆婆",
        '55': "岳父",
        '56': "岳母",
        '57': "继父或养父",
        '58': "继母或养母",
        '59': "其他父母关系",
        '60': "祖父母或外祖父母",
        '61': "祖父",
        '62': "祖母",
        '63': "外祖父",
        '64': "外祖母",
        '65': "配偶的祖父母或外祖父母",
        '66': "曾祖父",
        '67': "曾祖母",
        '68': "配偶的曾祖父或外曾祖父",
        '69': "他祖父母或外祖父母关系",
        '70': "兄弟姐妹",
        '71': "兄",
        '72': "嫂",
        '73': "弟",
        '74': "弟媳",
        '75': "姐姐",
        '76': "姐夫",
        '77': "妹妹",
        '78': "妹夫",
        '79': "其他兄弟姐妹",
        '80': "其他",
        '81': "伯父",
        '82': "伯母",
        '83': "叔父",
        '84': "婶母",
        '85': "舅父",
        '86': "舅母",
        '87': "姨父",
        '88': "姨母",
        '89': "姑父",
        '90': "姑母",
        '91': "堂兄弟、堂姐妹",
        '92': "表兄弟、表姐妹",
        '93': "侄子",
        '94': "侄女",
        '95': "外甥",
        '96': "外甥女",
        '97': "其他亲属",
        '98': "非亲属",
        '99': "其他关系",
        '100': "其他关系",
    }

    return code_name_map.get(code, "其他关系") if code else ""


def xb_code_to_name(code):
    """
    将 XB 代码转换为对应的性别名称。
    """
    code_name_map = {
        '1': "男",
        '2': "女",
        '3': "未知"
    }
    return code_name_map.get(str(code), "") if code else ""


# -------------------------------------------------------------
# 农村土地承包调查标准代码字典定义
# -------------------------------------------------------------
DKLB_MAP = {
    "10": "承包地块",
    "21": "自留地",
    "22": "机动地",
    "23": "开荒地",
    "99": "其他集体土地"
}

DLDJ_MAP = {
    "01": "一等地",
    "02": "二等地",
    "03": "三等地",
    "04": "四等地",
    "05": "五等地",
    "06": "六等地",
    "07": "七等地",
    "08": "八等地",
    "09": "九等地",
    "10": "十等地"
}

JZDLX_MAP = {
    "1": "实测法界址点",
    "2": "航测法界址点",
    "3": "图解法界址点"
}

JBLX_MAP = {
    "1": "钢钉",
    "2": "水泥桩",
    "3": "石灰桩",
    "4": "喷涂",
    "5": "瓷标志",
    "6": "无标志",
    "7": "木桩",
    "8": "埋石",
    "9": "其他界标"
}

JZXLB_MAP = {
    "01": "田埂",
    "02": "沟渠",
    "03": "道路",
    "04": "行树",
    "05": "围墙",
    "06": "墙壁",
    "07": "栅栏",
    "08": "两点连线",
    "09": "其他界线"
}

JZXWZ_MAP = {
    "01": "内",
    "02": "中",
    "03": "外",
    "1": "内",
    "2": "中",
    "3": "外"
}

ZJLX_MAP = {
    "1": "居民身份证",
    "2": "军官证",
    "3": "行政、企事业单位机构代码证或法人代码证",
    "4": "户口簿",
    "5": "护照",
    "9": "其他证件"
}

CBFLX_MAP = {
    "1": "农户",
    "2": "个人",
    "3": "单位"
}

XB_MAP = {
    "1": "男",
    "2": "女"
}

SF_MAP = {
    "1": "是",
    "2": "否"
}

TDYT_MAP = {
    "1": "种植业",
    "2": "林业",
    "3": "畜牧业",
    "4": "渔业",
    "5": "非农业用途"
}


SYQXZ_MAP = {
    "10": "国有土地所有权",
    "30": "集体土地所有权",
    "31": "村民小组",
    "32": "村级集体经济组织",
    "33": "乡级集体经济组织",
    "34": "其他农民集体经济组织",
}


def syqxz_code_to_name(code):
    """
    将 SYQXZ 代码转换为对应的所有权性质名称。
    """
    return SYQXZ_MAP.get(str(code), str(code)) if code else ""


def dklb_code_to_name(code):
    """
    将 DKLB 代码转换为对应的地块类别名称。
    """
    return DKLB_MAP.get(str(code), str(code)) if code else ""


def tdlylx_code_to_name(code):
    """
    将 TDLYLX 代码转换为对应的土地利用类型名称。
    """
    code_name_map = {
        '01': "耕地",
        '011': "水田",
        '012': "水浇地",
        '013': "旱地",
        '02': "园地",
        '021': "果园",
        '022': "茶园",
        '023': "其他园地",
        '03': "林地",
        '031': "有林地",
        '032': "灌木林地",
        '033': "其他林地",
        '04': "草地",
        '041': "天然草地",
        '042': "人工牧草地",
        '043': "其他草地",
        '05': "商业服务业用地",
        '051': "批发零售用地",
        '052': "餐饮用地",
        '053': "商务金融用地",
        '054': "其他商服用地",
        '06': "工矿用地",
        '061': "工业用地",
        '062': "采矿用地",
        '063': "仓储用地",
        '07': "住宅用地",
        '071': "城镇住宅用地",
        '072': "农村宅基地",
        '08': "公共管理与公共服务用地",
        '081': "机关团体用地",
        '082': "新闻出版用地",
        '083': "教育用地",
        '084': "医疗卫生用地",
        '085': "文化设施用地",
        '086': "社会福利用地",
        '087': "公园与绿地",
        '088': "风景名胜设施用地",
        '09': "特殊用地",
        '091': "军事设施用地",
        '092': "使领馆用地",
        '093': "监教场所用地",
        '094': "宗教用地",
        '095': "殡葬用地",
        '10': "交通运输用地",
        '101': "铁路用地",
        '102': "公路用地",
        '103': "城镇村道路用地",
        '104': "农村道路",
        '105': "机场用地",
        '106': "港口码头用地",
        '107': "管道运输用地",
        '11': "水域及水利设施用地",
        '111': "河流水面",
        '112': "湖泊水面",
        '113': "水库水面",
        '114': "坑塘水面",
        '115': "沿海滩涂",
        '116': "内陆滩涂",
        '117': "沟渠",
        '118': "水工建筑用地",
        '119': "冰川及永久积雪",
        '12': "其他土地",
        '121': "空闲地",
        '122': "设施农用地",
        '123': "田坎",
        '124': "盐碱地",
        '125': "沙地",
        '126': "裸土地",
        '127': "裸岩石砾地",
        '20': "扩展用地",
        '204': "采矿用地",
        '205': "特殊用地",
    }
    return code_name_map.get(str(code), str(code)) if code else ""


def dldj_code_to_name(code):
    """
    将 DLDJ 代码转换为对应的地力等级名称。
    """
    return DLDJ_MAP.get(str(code), str(code)) if code else ""


def sfjbnt_code_to_name(code):
    """
    将 SFJBNT 代码转换为对应的是否基本农田名称。
    """
    return SF_MAP.get(str(code), str(code)) if code else ""


def jzdlx_code_to_name(code):
    """
    将界址点类型代码转换为对应名称。
    """
    return JZDLX_MAP.get(str(code), str(code)) if code else ""


def jblx_code_to_name(code):
    """
    将界标类型代码转换为对应名称。
    """
    return JBLX_MAP.get(str(code), str(code)) if code else ""


def jzxlb_code_to_name(code):
    """
    将界址线类别代码转换为对应名称。
    """
    return JZXLB_MAP.get(str(code), str(code)) if code else ""


def jzxwz_code_to_name(code):
    """
    将界址线位置代码转换为对应名称。
    """
    return JZXWZ_MAP.get(str(code), str(code)) if code else ""


def zjlx_code_to_name(code):
    """
    将证件类型代码转换为对应名称。
    """
    return ZJLX_MAP.get(str(code), str(code)) if code else ""


def cbflx_code_to_name(code):
    """
    将承包方类型代码转换为对应名称。
    """
    return CBFLX_MAP.get(str(code), str(code)) if code else ""


def xb_code_to_name(code):
    """
    将性别代码转换为对应名称（支持 1/男、2/女）。
    """
    if not code:
        return ""
    code_str = str(code).strip()
    return XB_MAP.get(code_str, code_str)


def sf_code_to_name(code):
    """
    将是否代码转换为对应名称（1->是，2->否）。
    """
    return SF_MAP.get(str(code), str(code)) if code else ""


def tdyt_code_to_name(code):
    """
    将土地用途代码转换为对应名称。
    """
    if not code:
        return ""
    code_str = str(code).strip()
    return TDYT_MAP.get(code_str, code_str)


def _load_code_table_file(path):
    """读取一个权属代码表文件（.xlsx 或 .xls），返回 {代码(14位str): 名称} 的字典"""
    code_map = {}
    _, ext = os.path.splitext(path)
    if ext.lower() == '.xlsx':
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            ws = wb.active
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    continue
                if len(row) < 2:
                    continue
                code = str(row[0]).strip() if row[0] is not None else ''
                name = str(row[1]).strip() if row[1] is not None else ''
                if code and name:
                    code_map[code] = name
            wb.close()
        except Exception as e:
            logger.error(f"读取权属代码表 .xlsx 失败 {path}: {str(e)}")
            return {}
    elif ext.lower() == '.xls':
        try:
            import xlrd
            wb = xlrd.open_workbook(path)
            ws = wb.sheets()[0]
            for i in range(ws.nrows):
                if i == 0:
                    continue
                code_raw = ws.cell(i, 0).value
                name_raw = ws.cell(i, 1).value
                if isinstance(code_raw, float):
                    code = str(int(code_raw))
                elif isinstance(code_raw, int):
                    code = str(code_raw)
                else:
                    code = str(code_raw).strip() if code_raw is not None else ''
                name = str(name_raw).strip() if name_raw is not None else ''
                if code and name:
                    code_map[code] = name
            wb.release_resources()
        except Exception as e:
            logger.error(f"读取权属代码表 .xls 失败 {path}: {str(e)}")
            return {}
    return code_map


def _find_code_table_file(resources_dir, region_prefix):
    """在 resources_dir 中查找以 region_prefix 开头的权属单位代码表文件，返回完整路径"""
    if not os.path.isdir(resources_dir):
        return None
    for fname in os.listdir(resources_dir):
        if fname.startswith(region_prefix) and '权属单位代码表' in fname:
            fpath = os.path.join(resources_dir, fname)
            if os.path.isfile(fpath):
                return fpath
    return None


def load_ownership_code_table(region_prefix=None):
    """加载资源目录下的权属代码表，返回 {权属单位代码(14位str): 权属单位名称} 的字典
    region_prefix: 区域前缀（如 '341524'），搜索 resources/ 中以该前缀开头的权属单位代码表文件
    """
    resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
    code_map = {}
    if region_prefix:
        fpath = _find_code_table_file(resources_dir, region_prefix)
        if fpath and os.path.exists(fpath):
            code_map = _load_code_table_file(fpath)
            if code_map:
                logger.info(f"已加载权属代码表 ({region_prefix})：{os.path.basename(fpath)}，共 {len(code_map)} 条记录")
                return code_map
        logger.warning(f"权属代码表不存在 (region={region_prefix})，将回退到原始FBFMC逻辑")
        return {}
    generic = os.path.join(resources_dir, '权属代码表.xlsx')
    if os.path.exists(generic):
        code_map = _load_code_table_file(generic)
        if code_map:
            logger.info(f"已加载通用权属代码表，共 {len(code_map)} 条记录")
            return code_map
    logger.warning("权属代码表不存在，将回退到原始FBFMC逻辑")
    return {}


def parse_ownership_name(name):
    """解析权属单位名称为 (镇名, 村名, 组名) 三元组
    支持 "明光市明光街道蔬菜村股份经济合作社大李一组" 或 "金寨县梅山镇蔬菜村股份经济合作社第一组" 等格式
    """
    if not name:
        return ('', '', '')
    for prefix in ['安徽省滁州市明光市', '明光市', '安徽省六安市金寨县', '金寨县']:
        if name.startswith(prefix):
            name = name[len(prefix):]
            break
    town = village = group = ''
    town_end = -1
    for sep in ['街道', '镇']:
        idx = name.find(sep)
        if idx != -1:
            town_end = idx + len(sep)
            break
    if town_end != -1:
        town = name[:town_end]
        rest = name[town_end:]
    else:
        rest = name
    village_end = rest.find('村')
    if village_end != -1:
        village = rest[:village_end + 1]
        rest = rest[village_end + 1:]
    else:
        rest = ''
    group = rest.replace('股份经济合作社', '').strip()
    return (town, village, group)


def merge_excel_files(family_info_path, land_info_path, output_dir, jtjjzz_name):
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'YHZGX': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str})  # 强制将DKBM读取为字符串

    # 确保每个CBFBM对应的家庭成员信息和地块信息不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 将日期列转换为日期格式，确保它们可以被格式化
    family_info['CBQXQ'] = pd.to_datetime(family_info['CBQXQ'], errors='coerce')
    family_info['CBQXZ'] = pd.to_datetime(family_info['CBQXZ'], errors='coerce')
    
    # 按FBFMC分类
    def extract_group_name(fbfmc):
        if '合作社' in fbfmc:
            group = fbfmc.split('合作社')[-1].strip()
            return group if group else fbfmc
        elif '委员会' in fbfmc:
            group = fbfmc.split('委员会')[-1].strip()
            return group if group else fbfmc
        return fbfmc

    family_info['组名'] = family_info['FBFMC'].apply(extract_group_name)
    unique_groups = family_info['组名'].unique()

    for group in unique_groups:
        # 筛选出当前分组的家庭成员信息和地块信息
        group_family_info = family_info[family_info['组名'] == group]
        group_land_info = land_info[land_info['CBFBM'].isin(group_family_info['CBFBM'])]
        fbfmc = group_family_info['FBFMC'].iloc[0] 
        # 创建一个新的DataFrame来存储最终的结果
        result = []

        # 遍历每个CBFBM
        unique_codes = set(group_family_info['CBFBM']).union(set(group_land_info['CBFBM']))
        for code in unique_codes:
            # 获取当前CBFBM对应的家庭成员信息
            family_rows = group_family_info[group_family_info['CBFBM'] == code]
            # 获取当前CBFBM对应的地块信息
            land_rows = group_land_info[group_land_info['CBFBM'] == code]

            # 获取合同期限信息
            cbqxq = family_rows['CBQXQ'].iloc[0] if not family_rows.empty and pd.notna(family_rows['CBQXQ'].iloc[0]) else None
            cbqxz = family_rows['CBQXZ'].iloc[0] if not family_rows.empty and pd.notna(family_rows['CBQXZ'].iloc[0]) else None
            
            # 格式化合同期限为字符串
            cbqxq_str = cbqxq.strftime('%Y年%m月%d日') if cbqxq is not None else ''
            cbqxz_str = cbqxz.strftime('%Y年%m月%d日') if cbqxz is not None else ''
            contract_period = f"{cbqxq_str} - {cbqxz_str}" if cbqxq_str and cbqxz_str else ''

            # 计算家庭成员信息和地块信息的数量
            num_family_rows = len(family_rows)
            num_land_rows = len(land_rows)

            # 确定最终的行数
            max_rows = max(num_family_rows, num_land_rows)

            # 构建最终的行数据
            for i in range(max_rows):
                family_row = family_rows.iloc[i] if i < num_family_rows else {'CBFBM': code, 'CYXM': '', 'YHZGX': '', 'CYZJHM': '', 'LXDH_x': ''}
                land_row = land_rows.iloc[i] if i < num_land_rows else {'CBFBM': code, 'DKBM': '', 'DKMC': '', 'DKDZ': '', 'DKNZ': '', 'DKXZ': '', 'DKBZ': '', 'HTMJM': '','DKBZXX':''}

                # 合并行数据
                row = {
                    'CBFBM': str(code),  # 确保显示为文本
                    'CBFMC': '',
                    'CYXM': family_row['CYXM'],
                    'YHZGX': str(yhzgx_code_to_name(family_row['YHZGX'])),  # 确保显示为文本
                    'CYZJHM': family_row['CYZJHM'],
                    'LXDH': family_row['LXDH_x'],
                    'DKBM': str(land_row['DKBM']),  # 确保显示为文本
                    'DKMC': land_row['DKMC'],
                    'DKDZ': land_row['DKDZ'],
                    'DKNZ': land_row['DKNZ'],
                    'DKXZ': land_row['DKXZ'],
                    'DKBZ': land_row['DKBZ'],
                    'HTMJM': land_row['HTMJM'],
                    'SCZMJM': None,  # 初始化为None
                    'BZ': land_row['DKBZXX'],
                    'CONTRACT_PERIOD': contract_period  # 添加合同期限信息
                }
                result.append(row)

        # 将结果转换为DataFrame
        result_df = pd.DataFrame(result)

        # 计算每个CBFBM的实测总面积亩
        total_areas = group_land_info.groupby('CBFBM')['HTMJM'].sum().reset_index()
        total_areas.columns = ['CBFBM', 'SCZMJM_y']

        # 确保 'CBFBM' 列的数据类型一致
        result_df['CBFBM'] = result_df['CBFBM'].astype(str)
        total_areas['CBFBM'] = total_areas['CBFBM'].astype(str)

        # 合并实测总面积亩
        result_df = pd.merge(result_df, total_areas, on='CBFBM', how='left')

        # 将实测总面积亩_y 的值赋值给 实测总面积亩
        result_df['SCZMJM'] = result_df['SCZMJM_y']
        result_df = result_df.drop(columns=['SCZMJM_y'])

        # 设置承包方名称
        for code in unique_codes:
            # 过滤掉"与户主关系"为空的记录
            filtered_family_info = group_family_info[(group_family_info['CBFBM'] == code) & (group_family_info['YHZGX'].isin(['02', '01'])) & (group_family_info['YHZGX'].notna())]
            if not filtered_family_info.empty:
                household_name = filtered_family_info['CYXM'].iloc[0]
            else:
                household_name = ''  # 如果没有符合条件的记录，设置为空字符串
            result_df.loc[result_df['CBFBM'] == str(code), 'CBFMC'] = household_name

        # 添加XH字段
        result_df['XH'] = 0
        current_number = 1
        for code in unique_codes:
            result_df.loc[result_df['CBFBM'] == str(code), 'XH'] = current_number
            current_number += 1

        # 重新排序列
        result_df = result_df[['XH', 'CBFBM', 'CBFMC', 'CYXM', 'YHZGX', 'CYZJHM', 'LXDH', 'DKBM', 'DKMC', 'DKDZ', 'DKNZ', 'DKXZ', 'DKBZ', 'HTMJM', 'SCZMJM', 'BZ', 'CONTRACT_PERIOD']]

        # 保存到新的Excel文件
        output_file = f"{output_dir}/公示表（{group}）.xlsx"
        result_df.to_excel(output_file, index=False, sheet_name='Sheet1')

        # 使用openpyxl进行进一步处理
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
        keyword = "村民委员会"
        zm_name = fbfmc.split(keyword)[-1]
        # 找到最后一页的行数
        # 插入标题行
        ws.insert_rows(1, 2)
        if group:
            ws['A1'] = f"农村土地二轮承包到期后再延长三十年承包信息公示表({group})"
        else:
            ws['A1'] = f"农村土地二轮承包到期后再延长三十年承包信息公示表"
        ws['A2'] = "发包方："
        ws.merge_cells('A1:Q1')  # 修改为包含所有列
        ws['A2'] = f"发包方：{jtjjzz_name}"
        ws.merge_cells('A2:H2')
        ws['I2'] = "公示时间："
        ws.merge_cells('I2:Q2')  # 修改为包含所有列

        # 获取第一个非空的合同期限值
        contract_period = ''
        for index, row in result_df.iterrows():
            if row['CONTRACT_PERIOD']:
                contract_period = row['CONTRACT_PERIOD']
                break
        
        # 在公示时间后添加合同期限
        if contract_period:
            ws['I2'] = f"公示时间："
        else:
            ws['I2'] = "公示时间："

        # 设置表头
        headers = [
            "序号", "承包方编码", "承包方姓名", "成员姓名", "与户主关系", "身份证号码", "联系电话",
            "地块编码", "地块名称", "东至", "南至", "西至", "北至", "实测面积亩", "实测总面积亩", "备注"
        ]
        for col_num, header in enumerate(headers, 1):
            ws.cell(row=3, column=col_num, value=header).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # 隐藏 CONTRACT_PERIOD 列，不显示在表格中
        ws.column_dimensions[openpyxl.utils.get_column_letter(17)].hidden = True

        # 设置单元格格式
        for row in ws.iter_rows(min_row=3, max_col=16, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # 设置第二行内容居左
        for cell in ws[2]:
            cell.alignment = Alignment(horizontal='left', vertical='center')

        # 设置第一行格式
        for cell in ws[1]:
            cell.font = Font(size=18, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 32

        # 设置列宽
        #column_widths = [5, 15, 7.5, 8, 6.75, 13, 12.5, 15, 6.25, 7.75, 7.75, 7.75, 7.75, 5.5, 5.5, 5.5]
        column_widths = [4.5, 7.9, 6.8, 6.8, 6.75, 10.75, 12, 10.75, 6.25, 7.75, 7.75, 7.75, 7.75, 5.6, 5.6, 19]
        for col_num, width in enumerate(column_widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = width
        
        # 合并CBFBM和承包方名称单元格
        for code in unique_codes:
            rows = result_df[result_df['CBFBM'] == str(code)].index.tolist()
            if len(rows) > 1:
                start_row = rows[0] + 4  # Excel索引从1开始，Pandas从0开始
                end_row = rows[-1] + 4
                num_rows = end_row - start_row + 1  # 当前CBFBM的行数
                ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)  # 序号
                ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)  # 承包方编码
                ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)  # 承包方名称
                ws.merge_cells(start_row=start_row, start_column=15, end_row=end_row, end_column=15)  # 实测面积亩
           
        # 设置关键字段的单元格格式为文本
        for row in ws.iter_rows(min_row=4, max_col=16, max_row=ws.max_row):
            row[1].number_format = '@'  # 序号(XH)
            row[2].number_format = '@'  # 承包方编码(CBFBM)
            row[5].number_format = '@'  # 与户主关系(YHZGX)
            row[6].number_format = '@'  # 身份证号码(CYZJHM)
            row[7].number_format = '@'  # 联系电话(LXDH)
            row[8].number_format = '@'  # 地块编码(DKBM)
        # 设置纸张大小、方向和页边距
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_margins = PageMargins(left=1.41/2.54, right=1.41/2.54, top=2.04/2.54, bottom=2.04/2.54)

        # 设置第三行为打印标题行
        ws.print_title_rows = '3:3'
        # 插入页脚页码
        ws.oddFooter.center.text = "第 &[Page] 页 共 &[Pages] 页"
         # 在最后一行下方添加注释
        ws.append([""])
        wb.save(output_file)
        # 保存工作簿
        print(f"处理完成，已保存到 {output_file}")

def generate_mdb_table(family_info_path, land_info_path, output_directory):
    """
    生成摸底调查表（独立的函数，可不依赖公示表生成）
    接收excel文件路径参数，可以单独调用
    """
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'YHZGX': str, 'LXDH_x': str, 'CYZJHM': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str})  # 强制将DKBM读取为字符串
    # 去重处理，确保数据不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 获取CBFBM的唯一值列表
    unique_cbfbm = list(set(family_info['CBFBM'].unique()).union(set(land_info['CBFBM'].unique())))

    # 获取模板文件夹路径（相对于GenerateGSB.py文件）
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '单户摸底调查表.xlsx')

    for cbfbm in unique_cbfbm:
        # 过滤当前CBFBM的家庭成员信息和地块信息
        family_data = family_info[family_info['CBFBM'] == cbfbm]
        land_data = land_info[land_info['CBFBM'] == cbfbm]

        # 如果当前CBFBM没有有效数据，跳过
        if family_data.empty:
            continue

        # 获取承包方名称（CBFMC）
        cbfmc = family_data['CBFMC'].dropna().iloc[0] if 'CBFMC' in family_data.columns and not family_data['CBFMC'].dropna().empty else ''

        # 生成文件名：{CBFBM+CBFMC-摸底调查表.xlsx}
        filename = f"{cbfbm}{cbfmc}-摸底调查表.xlsx"
        output_file = os.path.join(output_directory, filename)

        # 复制模板文件
        shutil.copy2(template_path, output_file)

        # 打开工作簿
        wb = openpyxl.load_workbook(output_file)
        ws = wb['Sheet1']  # 使用Sheet1作为工作表

        # 填充数据
        # B4: CBFMC（承包方名称）
        ws.cell(row=4, column=2).value = cbfmc

        # E4: CBFZJHM（承包方证件号码），通常是户主的证件号码
        household_name_row = family_data[family_data['YHZGX'].isin(['01', '02', '户主'])]
        if not household_name_row.empty:
            # 尝试获取CBFZJHM或CYZJHM字段
            if 'CBFZJHM' in household_name_row.columns and not household_name_row['CBFZJHM'].empty:
                ws.cell(row=4, column=5).value = str(household_name_row['CBFZJHM'].iloc[0])
            elif 'CYZJHM' in household_name_row.columns and not household_name_row['CYZJHM'].empty:
                ws.cell(row=4, column=5).value = str(household_name_row['CYZJHM'].iloc[0])

        # I4: LXDH（联系电话）
        if 'LXDH_x' in family_data.columns and not family_data['LXDH_x'].dropna().empty:
            ws.cell(row=4, column=9).value = str(family_data['LXDH_x'].dropna().iloc[0])
            ws.cell(row=4, column=9).number_format = '@'  # 设置为文本格式

        # C5: CBFBM + "J"
        ws.cell(row=5, column=3).value = str(cbfbm) + 'J'

        # H5: CBQXQ-CBQXZ（承包期）
        # 获取合同期限信息
        cbqxq = family_data['CBQXQ'].iloc[0] if not family_data.empty and pd.notna(family_data['CBQXQ'].iloc[0]) else None
        cbqxz = family_data['CBQXZ'].iloc[0] if not family_data.empty and pd.notna(family_data['CBQXZ'].iloc[0]) else None
            
        # 格式化合同期限为字符串
        cbqxq_str = cbqxq.strftime('%Y年%m月%d日') if cbqxq is not None else ''
        cbqxz_str = cbqxz.strftime('%Y年%m月%d日') if cbqxz is not None else ''
        contract_period = f"{cbqxq_str} - {cbqxz_str}" if cbqxq_str and cbqxz_str else ''
        ws.cell(row=5, column=8).value = contract_period

        # C6: CBFDZ（承包方地址）
        if 'CBFDZ' in family_data.columns and not family_data['CBFDZ'].dropna().empty:
            ws.cell(row=6, column=3).value = str(family_data['CBFDZ'].dropna().iloc[0])

        # 计算家庭成员数量
        family_members = family_data
        n = len(family_members)
        land_rows = land_data
        m = len(land_rows)

        # 1. 从第9行开始插入新行，并复制第九行的样式，并写入家庭成员信息，成员个数n
        # 如果家庭成员个数n>1，则在第10行上方插入n-1行
        if n > 1:
            rows_to_insert = n - 1
            # 在第10行上方插入所需行数
            for i in range(rows_to_insert):
                ws.insert_rows(10)
                # 复制第10行的样式到新插入的行
                for col in range(1, 9):  # A-H列
                    source_cell = ws.cell(row=10 + rows_to_insert, column=col)
                    target_cell = ws.cell(row=10 + i, column=col)
                    target_cell.font = source_cell.font.copy()
                    target_cell.border = source_cell.border.copy()
                    target_cell.fill = source_cell.fill.copy()
                    target_cell.number_format = source_cell.number_format
                    target_cell.protection = source_cell.protection.copy()
                    target_cell.alignment = source_cell.alignment.copy()
                    # 清空目标单元格的值
                    target_cell.value = None

        # 从第9行开始写入家庭成员信息
        for i in range(n):
            row_num = 9 + i
            if i < len(family_members):
                member = family_members.iloc[i]
                # A列：CYXM
                ws.cell(row=row_num, column=1).value = member['CYXM']
                # B列：与户主关系（name）
                ws.cell(row=row_num, column=2).value = yhzgx_code_to_name(str(member['YHZGX']))
                # C列：CYZJHM
                ws.cell(row=row_num, column=3).value = member['CYZJHM']
                ws.merge_cells(start_row=row_num, start_column=3, end_row=row_num, end_column=6)
                # G列："是"
                ws.cell(row=row_num, column=7).value = "是"
                # H列："是"
                ws.cell(row=row_num, column=8).value = "是"
                
                # 设置单元格样式：字体大小10；行高：33.6；居中
                font = Font(name='宋体', size=10)
                alignment = Alignment(horizontal='center', vertical='center')
                for col in range(1, 9):  # A-H列
                    cell = ws.cell(row=row_num, column=col)
                    cell.font = font
                    cell.alignment = alignment
                    # 设置关键字段为文本格式
                    if col == 3:  # C列：CYZJHM（身份证号码）
                        cell.number_format = '@'
                # 行高设置为33.6
                ws.row_dimensions[row_num].height = 33.6

        # 2. 9+n行：农村土地承包经营权承包地块情况
        member_title_row = 9 + n
        ws.cell(row=member_title_row, column=1).value = "农村土地承包经营权承包地块情况"
        # 合并A-I列单元格
        ws.merge_cells(start_row=member_title_row, start_column=1, end_row=member_title_row, end_column=9)
        # 设置字体样式：居中、字体大小14、加粗
        font = Font(name='宋体', size=14, bold=True)
        alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=member_title_row, column=1).font = font
        ws.cell(row=member_title_row, column=1).alignment = alignment
        # 行高设置为25
        ws.row_dimensions[member_title_row].height = 25

        # 3. 9+n+1行和9+n+2行：地块相关信息表头
        land_header_row1 = 9 + n + 1
        land_header_row2 = 9 + n + 2
        
        # 9+n+1行内容
        ws.cell(row=land_header_row1, column=1).value = "地块代码"
        ws.cell(row=land_header_row1, column=2).value = "地块名称"
        
        # 计算确权面积总和
        total_area = round(land_rows['HTMJM'].sum(), 2) if not land_rows.empty else 0.00
        ws.cell(row=land_header_row1, column=7).value = f"确权面积   ({total_area}亩)"
        ws.cell(row=land_header_row1, column=8).value = "复核面积  (延包面积)"
        ws.cell(row=land_header_row1, column=9).value = "变更情况"
        
        # C-F列："坐落(四至)"
        ws.cell(row=land_header_row1, column=3).value = "坐落(四至)"
        # 合并C-F列单元格
        ws.merge_cells(start_row=land_header_row1, start_column=3, end_row=land_header_row1, end_column=6)
        # 设置单元格样式：居中，加粗，字体大小13，换行
        font = Font(name='宋体', size=13, bold=True)
        alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for col in range(1, 10):
            cell1 = ws.cell(row=land_header_row1, column=col)
            cell1.font = font
            cell1.alignment = alignment
        ws.row_dimensions[land_header_row1].height = 25 
        # 9+n+2行内容
        ws.cell(row=land_header_row2, column=3).value = "东至"
        ws.cell(row=land_header_row2, column=4).value = "南至"
        ws.cell(row=land_header_row2, column=5).value = "西至"
        ws.cell(row=land_header_row2, column=6).value = "北至"
        # 设置单元格样式：居中，加粗，字体大小13，换行
        for col in range(1, 10):
            cell2 = ws.cell(row=land_header_row2, column=col)
            if cell2.value is not None:
                cell2.font = font
                cell2.alignment = alignment
        ws.row_dimensions[land_header_row2].height = 25 
        # 合并9+n+1行和9+n+2行对应单元格
        # A列合并
        ws.merge_cells(start_row=land_header_row1, start_column=1, end_row=land_header_row2, end_column=1)
        # B列合并
        ws.merge_cells(start_row=land_header_row1, start_column=2, end_row=land_header_row2, end_column=2)
        # G列合并
        ws.merge_cells(start_row=land_header_row1, start_column=7, end_row=land_header_row2, end_column=7)
        # H列合并
        ws.merge_cells(start_row=land_header_row1, start_column=8, end_row=land_header_row2, end_column=8)
        # I列合并
        ws.merge_cells(start_row=land_header_row1, start_column=9, end_row=land_header_row2, end_column=9)

        # 4. 9+n+3行开始写入对应的地块信息，地块数m
        # 如果地块个数m>1，则在第(9+n+3+1)行上方插入m-1行
        land_start_row = 9 + n + 3
        if m > 1:
            rows_to_insert = m - 1
            # 在第(land_start_row+1)行上方插入所需行数
            for i in range(rows_to_insert):
                ws.insert_rows(land_start_row + 1)
                # 复制第(land_start_row+1)行的样式到新插入的行
                for col in range(1, 10):  # A-I列
                    source_cell = ws.cell(row=land_start_row + 1 + rows_to_insert, column=col)
                    target_cell = ws.cell(row=land_start_row + 1 + i, column=col)
                    target_cell.font = source_cell.font.copy()
                    target_cell.border = source_cell.border.copy()
                    target_cell.fill = source_cell.fill.copy()
                    target_cell.number_format = source_cell.number_format
                    target_cell.protection = source_cell.protection.copy()
                    target_cell.alignment = source_cell.alignment.copy()
                    # 清空目标单元格的值
                    target_cell.value = None
        
        # 从land_start_row行开始写入地块信息
        for i in range(m):
            row_num = land_start_row + i
            if i < len(land_rows):
                land = land_rows.iloc[i]
                # A列：DKBM
                ws.cell(row=row_num, column=1).value = land['DKBM']
                # B列：DKMC
                ws.cell(row=row_num, column=2).value = land['DKMC']
                # C列：DKDZ 东至
                ws.cell(row=row_num, column=3).value = land['DKDZ']
                # D列：DKNZ 南至
                ws.cell(row=row_num, column=4).value = land['DKNZ']
                # E列：DKXZ 西至
                ws.cell(row=row_num, column=5).value = land['DKXZ']
                # F列：DKBZ 北至
                ws.cell(row=row_num, column=6).value = land['DKBZ']
                # G列：HTMJM
                ws.cell(row=row_num, column=7).value = land['HTMJM']
                # H列：HTMJM（重复）
                ws.cell(row=row_num, column=8).value = land['HTMJM']
                # I列：DKBZXX（变更情况）
                ws.cell(row=row_num, column=9).value = land['DKBZXX'] if 'DKBZXX' in land else ''
                
                # 设置单元格样式：字体大小10、居中对齐
                font = Font(name='宋体', size=10)
                alignment = Alignment(horizontal='center', vertical='center',wrap_text=True)
                for col in range(1, 10):  # A-I列
                    cell = ws.cell(row=row_num, column=col)
                    cell.font = font
                    cell.alignment = alignment
                    # 设置关键字段为文本格式
                    if col == 1:  # A列：DKBM（地块编码）
                        cell.number_format = '@'
                # 行高设置为25
                ws.row_dimensions[row_num].height = 25

        # 5. 摸底情况和选项
        # (9+n+3+m行至9+n+3+m+2，A列)：合并单元格；content:"摸底情况"
        mdqk_start_row = land_start_row + m
        mdqk_end_row = mdqk_start_row + 2
        ws.cell(row=mdqk_start_row, column=1).value = "摸底情况"
        # A列合并单元格
        ws.merge_cells(start_row=mdqk_start_row, start_column=1, end_row=mdqk_end_row, end_column=1)
        # 设置单元格样式：居中，字体大小14，加粗
        font = Font(name='宋体', size=14, bold=True)
        alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=mdqk_start_row, column=1).font = font
        ws.cell(row=mdqk_start_row, column=1).alignment = alignment

        # (9+n+3+m,B至I列)：合并单元格；content："  □直接延包； □变更成员； □变更地块； □小调整； □消亡户； □分户； □合户 "
        ws.cell(row=mdqk_start_row, column=2).value = "  □直接延包； □变更成员； □变更地块； □小调整； □消亡户； □分户； □合户 "
        # B-I列合并单元格
        ws.merge_cells(start_row=mdqk_start_row, start_column=2, end_row=mdqk_start_row, end_column=9)
        # 设置字体大小为13
        font13 = Font(name='宋体', size=13)

        # 计算统计信息行
        # 统计信息对应行为mdqk_start_row+1 (即9+n+3+m+1)
        stats_row = mdqk_start_row + 1
        
        # 计算统计信息并添加在正确位置
        a = n  # 家庭人口数
        b = m  # 承包地块数
        c = round(land_rows['HTMJM'].sum(),2) if not land_rows.empty else 0.00  # 延包总面积
        ws.cell(row=stats_row, column=2).value = f"  该户现有家庭人口数：（{a}）人；承包地块数：（{b}）块；延包总面积：（{c}）亩"
        ws.merge_cells(start_row=stats_row, start_column=2, end_row=stats_row, end_column=9)
        # 设置单元格样式：居中、加粗、字体大小12、换行
        font12 = Font(name='宋体', size=12, bold=True)
        alignment12 = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.cell(row=stats_row, column=2).font = font12
        ws.cell(row=stats_row, column=2).alignment = alignment12
        ws.cell(row=stats_row+1, column=2).value = "  其他情况："
        ws.merge_cells(start_row=stats_row+1, start_column=2, end_row=stats_row+1, end_column=9)

        # 设置统计信息行及相邻行的行高为30
        ws.row_dimensions[stats_row-1].height = 30
        ws.row_dimensions[stats_row].height = 30
        ws.row_dimensions[stats_row+1].height = 30
        # 设置单元格样式：垂直居中、左对齐、字体大小13
        alignment_left_center = Alignment(horizontal='left', vertical='center')
        ws.cell(row=stats_row-1, column=2).font = font13
        ws.cell(row=stats_row-1, column=2).alignment = alignment_left_center
        ws.cell(row=stats_row, column=2).font = font13
        ws.cell(row=stats_row, column=2).alignment = alignment_left_center
        ws.cell(row=stats_row+1, column=2).font = font13
        ws.cell(row=stats_row+1, column=2).alignment = alignment_left_center
        ws.cell(row=stats_row+2, column=1).value = "  承包方（代表）确认"
        ws.merge_cells(start_row=stats_row+2, start_column=1, end_row=stats_row+2, end_column=2)
        ws.cell(row=stats_row+2, column=3).value = "          承包方（代表）：                      时间：    年   月   日"
        ws.merge_cells(start_row=stats_row+2, start_column=3, end_row=stats_row+2, end_column=9)
        ws.cell(row=stats_row+3, column=1).value = "  摸底结果审核意见"
        ws.merge_cells(start_row=stats_row+3, start_column=1, end_row=stats_row+5, end_column=1)
        ws.cell(row=stats_row+3, column=2).value = "延包工作小组审核意见"  
        ws.cell(row=stats_row+4, column=2).value = "村股份经济合作社审核意见" 
        ws.cell(row=stats_row+5, column=2).value = "乡镇审核意见"  
        ws.cell(row=stats_row+3, column=3).value = "                  审核人：                      时间：    年   月   日"  
        ws.merge_cells(start_row=stats_row+3, start_column=3, end_row=stats_row+3, end_column=9)
        ws.cell(row=stats_row+4, column=3).value = "                  审核人：                      时间：    年   月   日" 
        ws.merge_cells(start_row=stats_row+4, start_column=3, end_row=stats_row+4, end_column=9)
        ws.cell(row=stats_row+5, column=3).value = "                  审核人：                      时间：    年   月   日"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              
        ws.merge_cells(start_row=stats_row+5, start_column=3, end_row=stats_row+5, end_column=9)
        # 设置单元格样式：居中，字体大小14，加粗，换行
        font = Font(name='宋体', size=14, bold=True)
        alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.cell(row=stats_row+2, column=1).font = font
        ws.cell(row=stats_row+2, column=1).alignment = alignment
        ws.row_dimensions[stats_row+2].height = 57
        ws.cell(row=stats_row+3, column=1).font = font
        ws.cell(row=stats_row+3, column=1).alignment = alignment
        ws.cell(row=stats_row+3, column=2).font = font12
        ws.cell(row=stats_row+3, column=2).alignment = alignment12
        ws.cell(row=stats_row+4, column=2).font = font12
        ws.cell(row=stats_row+4, column=2).alignment = alignment12
        ws.cell(row=stats_row+5, column=2).font = font12
        ws.cell(row=stats_row+5, column=2).alignment = alignment12
        # 设置单元格样式：水平居中、底端对齐、换行、字体大小13
        font13 = Font(name='宋体', size=13)
        alignment13 = Alignment(horizontal='center', vertical='bottom', wrap_text=True)
        ws.cell(row=stats_row+2, column=3).font = font13
        ws.cell(row=stats_row+2, column=3).alignment = alignment13
        ws.cell(row=stats_row+3, column=3).font = font13
        ws.cell(row=stats_row+3, column=3).alignment = alignment13
        ws.cell(row=stats_row+4, column=3).font = font13
        ws.cell(row=stats_row+4, column=3).alignment = alignment13
        ws.cell(row=stats_row+5, column=3).font = font13
        ws.cell(row=stats_row+5, column=3).alignment = alignment13
        ws.row_dimensions[stats_row+3].height = 53
        ws.row_dimensions[stats_row+4].height = 53
        ws.row_dimensions[stats_row+5].height = 53
        # 为所有单元格添加边框
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 遍历工作表中的所有单元格并添加边框
        for row in ws.iter_rows(max_row=stats_row+5):
            for cell in row:
                cell.border = border
                
        # 保存工作簿
        wb.save(output_file)
       
        print(f"摸底调查表已保存到: {output_file}")


def generate_jinzhai_gsb(mdb_file_path, shp_file_path, output_directory, jtjjzz_name):
    """
    生成金寨县公示表
    使用金寨县调查公示表模板，内容从第5行开始填充
    """
    logger.debug(f"mdb_file_path: {mdb_file_path}")
    logger.debug(f"shp_file_path: {shp_file_path}")
    logger.debug(f"output_directory: {output_directory}")
    logger.debug(f"jtjjzz_name: {jtjjzz_name}")
    # 确保输出目录存在
    os.makedirs(output_directory, exist_ok=True)

    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

    # 创建MDB连接
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + mdb_file_path + ';'
    )
    logger.debug(f"conn_str: {conn_str}")
    conn = pyodbc.connect(conn_str)

    try:
        # 读取MDB文件中的表
        tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
        # 抑制pandas的SQLAlchemy连接警告
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}
        logger.debug(f"Tables read successfully: {tables}")

        # 处理CBF_JTCY表和CBF表
        cbf_jtcy = dfs['CBF_JTCY']
       
        cbf = dfs['CBF']
         # 假设'date_column'是包含日期的列
        cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
        cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
        # 处理CBHT表，提取CBQXQ和CBQXZ字段
        cbht = dfs['CBHT']
        cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
        cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
        # 选择所需字段
        cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]
        
        # 合并CBF_JTCY和CBF表，保留CBF表中的CBFBM字段
        merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left', suffixes=('', '_original'))
        logger.debug(f"Merged CBF_JTCY and CBF successfully")
        
        # 合并CBHT表
        merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
        logger.debug(f"Merged with CBHT successfully")

        # 用CBFBM的前14位挂接FBF表的FBFBM字段
        merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
        fbf = dfs['FBF']
        fbf['FBFDCRQ'] = pd.to_datetime(fbf['FBFDCRQ'], errors='coerce')
        merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')
        logger.debug(f"Merged with FBF successfully")

        # 处理CBDKXX表和矢量地块数据
        cbdxx = dfs['CBDKXX']
        # 查找名称中包含"DK"的shp文件
        shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
        logger.debug(f"Found SHP files: {shp_files_with_dk}")

        # 初始化一个空的GeoDataFrame来存储所有符合条件的矢量地块数据
        all_vector_data = gpd.GeoDataFrame()

        for shp_file in shp_files_with_dk:
            try:
                # 使用改进的SHP编码处理模块
                logger.debug(f"开始读取SHP文件: {shp_file}")
                vector_data = read_shp_safe(shp_file)
            except Exception as e:
                logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                raise
            all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)
        logger.debug(f"Read and concatenated SHP files successfully")

        try:
            # 使用改进的SHP编码处理模块
            logger.debug(f"开始读取主SHP文件: {shp_file_path}")
            vector_data = read_shp_safe(shp_file_path)
        except Exception as e:
            logger.error(f"读取主SHP文件失败 {shp_file_path}: {str(e)}")
            raise
        merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')
        logger.debug(f"Merged CBDKXX and vector data successfully")

        # 将数据写入两个Excel表格
        with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
            merged_all.to_excel(writer, sheet_name='Sheet1', index=False)
        logger.debug(f"Saved merged_all to Excel file: {excel_file1}")

        with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
            # 将GeoDataFrame转换为DataFrame以便写入Excel
            merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
            merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)
        logger.debug(f"Saved merged_cbdxx_vector to Excel file: {excel_file2}")

        merge_jinzhai_excel_files(excel_file1, excel_file2, output_directory, jtjjzz_name)
        logger.debug(f"Merge Jinzhai Excel files completed")

    except Exception as e:
        logger.error(f"发生错误: {e}")
        logger.error(traceback.format_exc())
        raise

    finally:
        # 关闭数据库连接
        conn.close()
        logger.debug("Database connection closed")


def merge_jinzhai_excel_files(family_info_path, land_info_path, output_dir, jtjjzz_name):
    """
    合并金寨县公示表Excel文件
    根据用户提供的表头图片格式：
    - 家庭成员：姓名、与承包方代表关系、身份证号
    - 地块具体情况：地块名称、地块编码、实测面积、确权面积、东至、南至、西至、北至
    - 公示备注
    """
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'YHZGX': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str})  # 强制将DKBM读取为字符串

    # 确保每个CBFBM对应的家庭成员信息和地块信息不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 将日期列转换为日期格式，确保它们可以被格式化
    family_info['CBQXQ'] = pd.to_datetime(family_info['CBQXQ'], errors='coerce')
    family_info['CBQXZ'] = pd.to_datetime(family_info['CBQXZ'], errors='coerce')
    
    # 按FBFMC分类
    def extract_group_name(fbfmc):
        if '合作社' in fbfmc:
            group = fbfmc.split('合作社')[-1].strip()
            return group if group else fbfmc
        elif '委员会' in fbfmc:
            group = fbfmc.split('委员会')[-1].strip()
            return group if group else fbfmc
        elif '经济组织' in fbfmc:
            group = fbfmc.split('集体经济组织')[-1].strip()
            return group if group else fbfmc
        elif '村' in fbfmc:
            group = fbfmc.split('村')[-1].strip()
            return group if group else fbfmc
        return fbfmc

    family_info['组名'] = family_info['FBFMC'].apply(extract_group_name)
    unique_groups = family_info['组名'].unique()

    for group in unique_groups:
        # 筛选出当前分组的家庭成员信息和地块信息
        group_family_info = family_info[family_info['组名'] == group]
        group_land_info = land_info[land_info['CBFBM'].isin(group_family_info['CBFBM'])]
        fbfmc = group_family_info['FBFMC'].iloc[0] 
        # 创建一个新的DataFrame来存储最终的结果
        result = []

        # 遍历每个CBFBM
        unique_codes = set(group_family_info['CBFBM']).union(set(group_land_info['CBFBM']))
        for code in unique_codes:
            # 获取当前CBFBM对应的家庭成员信息
            family_rows = group_family_info[group_family_info['CBFBM'] == code]
            # 获取当前CBFBM对应的地块信息
            land_rows = group_land_info[group_land_info['CBFBM'] == code]

            # 计算家庭成员信息和地块信息的数量
            num_family_rows = len(family_rows)
            num_land_rows = len(land_rows)

            # 确定最终的行数
            max_rows = max(num_family_rows, num_land_rows)

            # 构建最终的行数据
            for i in range(max_rows):
                family_row = family_rows.iloc[i] if i < num_family_rows else {'CBFBM': code, 'CYXM': '', 'YHZGX': '', 'CYZJHM': ''}
                land_row = land_rows.iloc[i] if i < num_land_rows else {'CBFBM': code, 'DKBM': '', 'DKMC': '', 'DKDZ': '', 'DKNZ': '', 'DKXZ': '', 'DKBZ': '', 'HTMJM': '', 'SCMJM': ''}

                # 合并行数据 - 按照金寨县公示表格式
                row = {
                    'CBFBM': str(code),  # 承包方编码(缩略码)
                    'CYXM': family_row['CYXM'],  # 姓名
                    'YHZGX': str(yhzgx_code_to_name(family_row['YHZGX'])),  # 与承包方代表关系
                    'CYZJHM': family_row['CYZJHM'],  # 身份证号
                    'DKMC': land_row['DKMC'],  # 地块名称
                    'DKBM': str(land_row['DKBM']),  # 地块编码
                    'SCMJM': land_row['SCMJM'] if 'SCMJM' in land_row and pd.notna(land_row['SCMJM']) else land_row['HTMJM'],  # 实测面积
                    'HTMJM': land_row['HTMJM'],  # 确权面积
                    'DKDZ': land_row['DKDZ'],  # 东至
                    'DKNZ': land_row['DKNZ'],  # 南至
                    'DKXZ': land_row['DKXZ'],  # 西至
                    'DKBZ': land_row['DKBZ'],  # 北至
                    'BZ': land_row['DKBZXX'] if 'DKBZXX' in land_row else '',  # 公示备注
                }
                result.append(row)

        # 将结果转换为DataFrame
        result_df = pd.DataFrame(result)

        # 设置承包方名称（用于标题）
        for code in unique_codes:
            # 过滤掉"与户主关系"为空的记录
            filtered_family_info = group_family_info[(group_family_info['CBFBM'] == code) & (group_family_info['YHZGX'].isin(['02', '01'])) & (group_family_info['YHZGX'].notna())]
            if not filtered_family_info.empty:
                household_name = filtered_family_info['CYXM'].iloc[0]
            else:
                household_name = ''  # 如果没有符合条件的记录，设置为空字符串
            result_df.loc[result_df['CBFBM'] == str(code), 'CBFMC'] = household_name

        # 重新排序列 - 按照金寨县公示表格式（不包含序号列）
        result_df = result_df[['CBFBM', 'CYXM', 'YHZGX', 'CYZJHM', 'DKMC', 'DKBM', 'SCMJM', 'HTMJM', 'DKDZ', 'DKNZ', 'DKXZ', 'DKBZ', 'BZ']]

        # 保存到新的Excel文件（不包含列名）
        output_file = f"{output_dir}/调查公示表（{group}）.xlsx"
        result_df.to_excel(output_file, index=False, sheet_name='Sheet1', header=False)

        # 使用openpyxl进行进一步处理
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
        keyword = "村民委员会"
        zm_name = fbfmc.split(keyword)[-1]
        
        # 插入标题行（第1-4行）
        ws.insert_rows(1, 4)
        
        # 第1行：主标题
        if group:
            ws['A1'] = f"农村土地承包经营权调查信息公示表({group})"
        else:
            ws['A1'] = f"农村土地承包经营权调查信息公示表"
        ws.merge_cells('A1:M1')  # 合并A-M列
        
        # 第2行：发包方
        ws['A2'] = f"发包方：{jtjjzz_name}"
        ws.merge_cells('A2:E2')
        ws['F2'] = "公示日期：      年   月   日至      年   月   日（共     天）(单位：亩、块）"
        ws.merge_cells('F2:M2')
        
        # 第3行：空行（用于合并表头）
        ws['A3'] = "承包方编码(缩略码)"
        ws['B3'] = "家庭成员"
        ws.merge_cells('B3:D3')  # 家庭成员合并3列：姓名、与承包方代表关系、身份证号
        ws['E3'] = "地块具体情况"
        ws.merge_cells('E3:L3')  # 地块具体情况合并7列：地块名称、地块编码、实测面积、确权面积、东至、南至、西至、北至
        ws['M3'] = "公示备注"
        
        # 第4行：详细列标题
        ws['A4'] = ""
        ws['B4'] = "姓名"
        ws['C4'] = "与承包方代表关系"
        ws['D4'] = "身份证号"
        ws['E4'] = "地块名称"
        ws['F4'] = "地块编码"
        ws['G4'] = "实测\n面积"
        ws['H4'] = "确权\n面积"
        ws['I4'] = "东至"
        ws['J4'] = "南至"
        ws['K4'] = "西至"
        ws['L4'] = "北至"
        ws['M4'] = "公示\n备注"

        # 合并A3:A4和M3:M4
        ws.merge_cells('A3:A4')
        ws.merge_cells('M3:M4')
        
        # 设置表头样式
        headers = [
            "承包方编码(缩略码)", "姓名", "与承包方代表关系", "身份证号",
            "地块名称", "地块编码", "实测面积", "确权面积", "东至", "南至", "西至", "北至", "公示备注"
        ]
        
        # 设置第3行（合并表头）样式
        for col_num in range(1, 14):
            cell = ws.cell(row=3, column=col_num)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.font = Font(bold=True)
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        ws.row_dimensions[3].height=18
        
        # 设置第4行（详细列标题）样式
        for col_num in range(1, 14):
            cell = ws.cell(row=4, column=col_num)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.font = Font(bold=True)
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        ws.row_dimensions[4].height=35
        # 设置单元格格式
        for row in ws.iter_rows(min_row=5, max_col=13, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # 设置第一行格式（主标题）
        for cell in ws[1]:
            cell.font = Font(size=18, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 32
        
        # 设置第二行格式
        for cell in ws[2]:
            cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.font = Font(size=10) 

        # 设置列宽（删除了序号列）
        column_widths = [15, 10, 10, 15, 10, 15, 8, 8, 9, 9, 9, 9, 8]
        for col_num, width in enumerate(column_widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = width
        
        # 合并CBFBM单元格
        for code in unique_codes:
            rows = result_df[result_df['CBFBM'] == str(code)].index.tolist()
            if len(rows) > 1:
                start_row = rows[0] + 5  # Excel索引从1开始，Pandas从0开始，加上前面的4行标题
                end_row = rows[-1] + 5
                ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)  # 承包方编码
                # 不合并姓名列，保持每个家庭成员的姓名独立显示
        
        # 设置关键字段的单元格格式为文本
        for row in ws.iter_rows(min_row=5, max_col=13, max_row=ws.max_row):
            row[1].number_format = '@'  # 承包方编码(CBFBM)
            row[3].number_format = '@'  # 与承包方代表关系(YHZGX)
            row[4].number_format = '@'  # 身份证号(CYZJHM)
            row[6].number_format = '@'  # 地块编码(DKBM)
        
        # 设置纸张大小、方向和页边距
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_margins = PageMargins(left=1.41/2.54, right=1.41/2.54, top=2.04/2.54, bottom=2.04/2.54)

        # 设置第3-4行为打印标题行
        ws.print_title_rows = '3:4'
        
        # 插入页脚页码
        ws.oddFooter.center.text = "第 &[Page] 页 共 &[Pages] 页"
        
        # 在最后一行下方添加审核信息
        ws.append(["延包小组审核：", "", "", "", "", "", "", "", "审核日期：    年   月   日"])
        # 合并I-M列
        last_row = ws.max_row
        ws.merge_cells(start_row=last_row, start_column=9, end_row=last_row, end_column=13)
        # 设置审核信息行的样式
        审核_cell = ws.cell(row=last_row, column=1)
        审核_cell.font = Font(bold=False)
        审核_cell.alignment = Alignment(horizontal='left', vertical='center')
        日期_cell = ws.cell(row=last_row, column=9)
        日期_cell.alignment = Alignment(horizontal='center', vertical='center')
        # 为审核信息行添加外侧边框
        for col in range(1, 14):
            cell = ws.cell(row=last_row, column=col)
            # 只设置外侧边框
            border = Border(
                left=Side(style='thin') if col == 1 else None,
                right=Side(style='thin') if col == 13 else None,
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            cell.border = border
        wb.save(output_file)
        # 保存工作簿
        print(f"金寨县公示表处理完成，已保存到 {output_file}")


def generate_jinzhai_mdb_table(family_info_path, land_info_path, output_directory):
    """
    生成金寨县摸底调查表（独立的函数，可不依赖公示表生成）
    接收excel文件路径参数，可以单独调用
    """
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'CBFBM':str,'YHZGX': str, 'CYXB': str, 'LXDH_x': str, 'CYZJHM': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str,'CBFBM':str})  # 强制将DKBM读取为字符串

    # 去重处理，确保数据不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 按FBFBM分组
    grouped = family_info.groupby('FBFBM')

    for fbfbm, group_family in grouped:
        # 获取组内的家庭成员信息和地块信息
        group_land = land_info[land_info['CBFBM'].isin(group_family['CBFBM'])]

        # 计算每个CBFBM的家庭成员数量
        cbfcy_count = group_family.groupby('CBFBM').size().reset_index(name='CBFCYSL')

        # 合并家庭成员数量
        group_family = pd.merge(group_family, cbfcy_count, on='CBFBM', how='left')

        # 按CBFBM升序排列
        group_family = group_family.sort_values('CBFBM')
        group_land = group_land.sort_values('CBFBM')

        # 获取FBFMC和FBFBM
        fbfmc = group_family['FBFMC'].iloc[0] if not group_family.empty else ''
        fbfbm_str = str(fbfbm)
        fbfmc_str = fbfmc.replace('安徽省六安市金寨县', '')
        # 生成文件名
        filename = f"{fbfmc_str}({fbfbm_str})-摸底调查表.xlsx"
        output_file = os.path.join(output_directory, filename)

        # 获取模板文件夹路径
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '单户摸底调查表_金寨县.xlsx')

        # 复制模板文件
        shutil.copy2(template_path, output_file)

        # 打开工作簿
        wb = openpyxl.load_workbook(output_file)
        ws = wb['Sheet1']  # 使用Sheet1作为工作表

        # A2: 单位：{FBFMC+\n+FBFBM}
        ws['A2'] = f"单位：{fbfmc}\n{fbfbm_str}"

        # 获取所有唯一的CBFBM
        unique_cbfbm = group_family['CBFBM'].unique()

        # 创建数据列表
        data_rows = []
        current_row = 6  # startrow: 第6行

        for cbfbm_val in unique_cbfbm:
            family_rows = group_family[group_family['CBFBM'] == cbfbm_val]
            land_rows = group_land[group_land['CBFBM'] == cbfbm_val]

            num_family = len(family_rows)
            num_land = len(land_rows)
            max_rows = max(num_family, num_land)

            for i in range(max_rows):
                family_row = family_rows.iloc[i] if i < num_family else None
                land_row = land_rows.iloc[i] if i < num_land else None

                row_data = {
                    'CBFBM': cbfbm_val,
                    'CBFMC': family_row['CBFMC'] if family_row is not None else '',
                    'CBFCYSL': family_row['CBFCYSL'] if family_row is not None else 0,
                    'CYXM': family_row['CYXM'] if family_row is not None else '',
                    'CYXB': xb_code_to_name(family_row['CYXB']) if family_row is not None and 'CYXB' in family_row.index else '',
                    'CYZJHM': family_row['CYZJHM'] if family_row is not None else '',
                    'YHZGX': yhzgx_code_to_name(family_row['YHZGX']) if family_row is not None else '',
                    'DKMC': land_row['DKMC'] if land_row is not None else '',
                    'DKBM': land_row['DKBM'] if land_row is not None else '',
                    'HTMJM': land_row['HTMJM'] if land_row is not None else '',
                    'DKDZ': land_row['DKDZ'] if land_row is not None else '',
                    'DKNZ': land_row['DKNZ'] if land_row is not None else '',
                    'DKXZ': land_row['DKXZ'] if land_row is not None else '',
                    'DKBZ': land_row['DKBZ'] if land_row is not None else '',
                    'LXDH': family_row['LXDH_x'] if family_row is not None else '',
                    'total_HTMJM': 0  # 将在后面计算
                }

                # 计算总面积
                total_area = land_rows['HTMJM'].sum() if not land_rows.empty else 0
                row_data['total_HTMJM'] = total_area

                data_rows.append(row_data)

        # 写入数据
        for i, row_data in enumerate(data_rows):
            row_num = current_row + i

            # A列：CBFBM
            ws.cell(row=row_num, column=1).value = row_data['CBFBM']
            # B列：CBFMC
            ws.cell(row=row_num, column=2).value = row_data['CBFMC']
            # C列："农户"
            ws.cell(row=row_num, column=3).value = "农户"
            # D列：CBFCYSL
            ws.cell(row=row_num, column=4).value = row_data['CBFCYSL']
            # E列：CYXM
            ws.cell(row=row_num, column=5).value = row_data['CYXM']
            # F列：CYXB
            ws.cell(row=row_num, column=6).value = row_data['CYXB']
            # G列："身份证"
            ws.cell(row=row_num, column=7).value = "身份证"
            # H列：CYZJHM
            ws.cell(row=row_num, column=8).value = row_data['CYZJHM']
            # I列：YHZGX
            ws.cell(row=row_num, column=9).value = row_data['YHZGX']
            # J列：""
            ws.cell(row=row_num, column=10).value = ""
            # K列：DKMC
            ws.cell(row=row_num, column=11).value = row_data['DKMC']
            # L列：DKBM
            ws.cell(row=row_num, column=12).value = row_data['DKBM']
            # M列：HTMJM
            ws.cell(row=row_num, column=13).value = row_data['HTMJM']
            # N列：sum{HTMJM}
            ws.cell(row=row_num, column=14).value = row_data['total_HTMJM']
            # O列：HTMJM
            ws.cell(row=row_num, column=15).value = row_data['HTMJM']
            # P列：sum{HTMJM}
            ws.cell(row=row_num, column=16).value = row_data['total_HTMJM']
            # Q列：DKDZ
            ws.cell(row=row_num, column=17).value = row_data['DKDZ']
            # R列：DKNZ
            ws.cell(row=row_num, column=18).value = row_data['DKNZ']
            # S列：DKXZ
            ws.cell(row=row_num, column=19).value = row_data['DKXZ']
            # T列：DKBZ
            ws.cell(row=row_num, column=20).value = row_data['DKBZ']
            # U列：LXDH
            ws.cell(row=row_num, column=21).value = row_data['LXDH']
            # V列：""
            ws.cell(row=row_num, column=22).value = ""

        # 合并单元格
        row_indices_per_cbfbm = {}
        for i, row_data in enumerate(data_rows):
            cbfbm_val = row_data['CBFBM']
            if cbfbm_val not in row_indices_per_cbfbm:
                row_indices_per_cbfbm[cbfbm_val] = []
            row_indices_per_cbfbm[cbfbm_val].append(current_row + i)

        for cbfbm_val, rows in row_indices_per_cbfbm.items():
            if len(rows) > 1:
                start_row = rows[0]
                end_row = rows[-1]
                # A列：CBFBM
                ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                # B列：CBFMC
                ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
                # C列："农户"
                ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)
                # D列：CBFCYSL
                ws.merge_cells(start_row=start_row, start_column=4, end_row=end_row, end_column=4)
                # N列：sum{HTMJM}
                ws.merge_cells(start_row=start_row, start_column=14, end_row=end_row, end_column=14)
                # P列：sum{HTMJM}
                ws.merge_cells(start_row=start_row, start_column=16, end_row=end_row, end_column=16)

        # 设置样式
        font = Font(name='宋体', size=10)
        alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        for row in ws.iter_rows(min_row=6, max_row=ws.max_row):
            for cell in row:
                cell.font = font
                cell.alignment = alignment
                cell.border = border

        # 设置行高
        for row_num in range(6, ws.max_row + 1):
            ws.row_dimensions[row_num].height = 13.5

        # 设置列宽（根据需要调整）
        #column_widths = [10, 10, 8, 8, 10, 8, 8, 15, 10, 5, 10, 15, 8, 8, 8, 8, 10, 10, 10, 10, 12, 5]
        #for col_num, width in enumerate(column_widths, 1):
            #ws.column_dimensions[get_column_letter(col_num)].width = width

        # 设置关键字段为文本格式
        for row in ws.iter_rows(min_row=6, max_row=ws.max_row):
            row[0].number_format = '@'  # A列：CBFBM
            row[7].number_format = '@'  # H列：CYZJHM
            row[11].number_format = '@'  # L列：DKBM
            row[20].number_format = '@'  # U列：LXDH

        # 保存工作簿
        wb.save(output_file)

        print(f"金寨县摸底调查表已保存到: {output_file}")


def generate_jinzhai_survey_info_table(family_info_path, land_info_path, output_directory, region_prefix=None):
    """
    生成金寨县土地承包经营权调查信息表
    与 generate_jinzhai_mdb_table 结构相同，但：
    - 模板不同
    - 无单位/日期行（row 2 已从模板中移除）
    - 数据从 row 5 开始（startrow=5）
    - row 1 标题添加 XX镇XX村 前缀
    - 末尾追加汇总行（复制模板汇总行样式）
    """
    def safe_read_excel(file_path, converters=None):
        try:
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            if os.path.exists(file_path):
                try:
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n1. 文件损坏或格式错误\n2. 文件不是有效的Excel文件\n3. 安装了xlrd包可能可以读取xls格式文件\n請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    family_info = safe_read_excel(family_info_path, converters={'CBFBM':str,'YHZGX': str, 'CYXB': str, 'LXDH_x': str, 'CYZJHM': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str,'CBFBM':str})
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    grouped = family_info.groupby('FBFBM')

    ownership_code_map = load_ownership_code_table(region_prefix)

    for fbfbm, group_family in grouped:
        group_land = land_info[land_info['CBFBM'].isin(group_family['CBFBM'])]
        cbfcy_count = group_family.groupby('CBFBM').size().reset_index(name='CBFCYSL')
        group_family = pd.merge(group_family, cbfcy_count, on='CBFBM', how='left')
        group_family = group_family.sort_values('CBFBM')
        group_land = group_land.sort_values('CBFBM')

        fbfbm_str = str(fbfbm).strip()
        town_name, village_name, group_name = parse_ownership_name(
            ownership_code_map.get(fbfbm_str, '')
        )
        if town_name or village_name or group_name:
            display_name = f'{town_name}{village_name}{group_name}'
            if group_name:
                header_name = f'{village_name}{group_name}'
            elif village_name:
                header_name = village_name
            else:
                header_name = town_name
        else:
            fbfmc = group_family['FBFMC'].iloc[0] if not group_family.empty else ''
            fbfmc_str = fbfmc.replace('安徽省六安市金寨县', '')
            display_name = fbfmc_str
            header_name = fbfmc_str
        filename = f"{display_name}土地承包经营权调查信息表.xlsx"
        output_file = os.path.join(output_directory, filename)

        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '土地承包经营权调查信息表_金寨县.xlsx')
        shutil.copy2(template_path, output_file)
        wb = openpyxl.load_workbook(output_file)
        ws = wb['Sheet1']

        ws['A1'] = f"{header_name}农村土地承包经营权调查信息表"

        # 保存模板汇总行样式（原始汇总行 = ws.max_row）
        from copy import copy as style_copy
        summary_row_idx = ws.max_row
        summary_styles = {}
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(summary_row_idx, col_idx)
            summary_styles[col_idx] = {
                'font': style_copy(cell.font),
                'border': style_copy(cell.border),
                'alignment': style_copy(cell.alignment),
                'fill': style_copy(cell.fill),
                'number_format': cell.number_format,
            }
        summary_merges = []
        for mr in ws.merged_cells.ranges:
            if mr.min_row == summary_row_idx:
                summary_merges.append((mr.min_col, mr.max_col))

        # 保存模板第5行（样表数据行）的样式用于后续复制
        template_row_styles = {}
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(5, col_idx)
            template_row_styles[col_idx] = {
                'font': style_copy(cell.font),
                'border': style_copy(cell.border),
                'alignment': style_copy(cell.alignment),
                'fill': style_copy(cell.fill),
                'number_format': cell.number_format,
            }

        # 删除模板数据区域（R5 ~ 末尾，含汇总行）
        ws.delete_rows(5, summary_row_idx - 4)

        # 重建合并单元格，仅保留表头合并（前4行），清除数据区残留合并
        header_merges = [str(mr) for mr in ws.merged_cells.ranges if mr.max_row < 5]
        ws.merged_cells = type(ws.merged_cells)()
        for hm in header_merges:
            ws.merge_cells(hm)

        # 清除模板残留的行维度
        for r in list(ws.row_dimensions.keys()):
            if r >= 5:
                del ws.row_dimensions[r]

        unique_cbfbm = group_family['CBFBM'].unique()
        data_rows = []
        current_row = 5

        for cbfbm_val in unique_cbfbm:
            family_rows = group_family[group_family['CBFBM'] == cbfbm_val]
            land_rows = group_land[group_land['CBFBM'] == cbfbm_val]
            num_family = len(family_rows)
            num_land = len(land_rows)
            max_rows = max(num_family, num_land)

            for i in range(max_rows):
                family_row = family_rows.iloc[i] if i < num_family else None
                land_row = land_rows.iloc[i] if i < num_land else None
                row_data = {
                    'CBFBM': cbfbm_val,
                    'CBFMC': family_row['CBFMC'] if family_row is not None else '',
                    'CBFCYSL': family_row['CBFCYSL'] if family_row is not None else 0,
                    'CYXM': family_row['CYXM'] if family_row is not None else '',
                    'CYXB': xb_code_to_name(family_row['CYXB']) if family_row is not None and 'CYXB' in family_row.index else '',
                    'CYZJHM': family_row['CYZJHM'] if family_row is not None else '',
                    'YHZGX': yhzgx_code_to_name(family_row['YHZGX']) if family_row is not None else '',
                    'DKMC': land_row['DKMC'] if land_row is not None else '',
                    'DKBM': land_row['DKBM'] if land_row is not None else '',
                    'HTMJM': land_row['HTMJM'] if land_row is not None else '',
                    'DKDZ': land_row['DKDZ'] if land_row is not None else '',
                    'DKNZ': land_row['DKNZ'] if land_row is not None else '',
                    'DKXZ': land_row['DKXZ'] if land_row is not None else '',
                    'DKBZ': land_row['DKBZ'] if land_row is not None else '',
                    'LXDH': family_row['LXDH_x'] if family_row is not None else '',
                    'total_HTMJM': 0
                }
                total_area = land_rows['HTMJM'].sum() if not land_rows.empty else 0
                row_data['total_HTMJM'] = total_area
                data_rows.append(row_data)

        for i, row_data in enumerate(data_rows):
            row_num = current_row + i
            ws.cell(row=row_num, column=1).value = row_data['CBFBM']
            ws.cell(row=row_num, column=2).value = row_data['CBFMC']
            ws.cell(row=row_num, column=3).value = "农户"
            ws.cell(row=row_num, column=4).value = row_data['CBFCYSL']
            ws.cell(row=row_num, column=5).value = row_data['CYXM']
            ws.cell(row=row_num, column=6).value = row_data['CYXB']
            ws.cell(row=row_num, column=7).value = "身份证"
            ws.cell(row=row_num, column=8).value = row_data['CYZJHM']
            ws.cell(row=row_num, column=9).value = row_data['YHZGX']
            ws.cell(row=row_num, column=10).value = ""
            ws.cell(row=row_num, column=11).value = row_data['DKMC']
            ws.cell(row=row_num, column=12).value = row_data['DKBM']
            ws.cell(row=row_num, column=13).value = row_data['HTMJM']
            ws.cell(row=row_num, column=14).value = row_data['total_HTMJM']
            ws.cell(row=row_num, column=15).value = row_data['HTMJM']
            ws.cell(row=row_num, column=16).value = row_data['total_HTMJM']
            ws.cell(row=row_num, column=17).value = row_data['DKDZ']
            ws.cell(row=row_num, column=18).value = row_data['DKNZ']
            ws.cell(row=row_num, column=19).value = row_data['DKXZ']
            ws.cell(row=row_num, column=20).value = row_data['DKBZ']
            ws.cell(row=row_num, column=21).value = row_data['LXDH']
            ws.cell(row=row_num, column=22).value = ""

        row_indices_per_cbfbm = {}
        for i, row_data in enumerate(data_rows):
            cbfbm_val = row_data['CBFBM']
            if cbfbm_val not in row_indices_per_cbfbm:
                row_indices_per_cbfbm[cbfbm_val] = []
            row_indices_per_cbfbm[cbfbm_val].append(current_row + i)

        for cbfbm_val, rows in row_indices_per_cbfbm.items():
            if len(rows) > 1:
                sr = rows[0]
                er = rows[-1]
                ws.merge_cells(start_row=sr, start_column=1, end_row=er, end_column=1)
                ws.merge_cells(start_row=sr, start_column=2, end_row=er, end_column=2)
                ws.merge_cells(start_row=sr, start_column=3, end_row=er, end_column=3)
                ws.merge_cells(start_row=sr, start_column=4, end_row=er, end_column=4)
                ws.merge_cells(start_row=sr, start_column=14, end_row=er, end_column=14)
                ws.merge_cells(start_row=sr, start_column=16, end_row=er, end_column=16)

        # 复用模板第5行样式应用到所有数据行
        last_data_row = current_row + len(data_rows) - 1
        for row_num in range(5, last_data_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row_num, col_idx)
                s = template_row_styles.get(col_idx)
                if s:
                    cell.font = style_copy(s['font'])
                    cell.border = style_copy(s['border'])
                    cell.alignment = style_copy(s['alignment'])
                    if s['fill']:
                        cell.fill = style_copy(s['fill'])
                    cell.number_format = s['number_format']

        # 追加汇总行（复制模板样式）
        sr = last_data_row + 1
        ws.cell(sr, 1).value = "汇审意见"
        ws.cell(sr, 9).value = "汇审人"
        ws.cell(sr, 14).value = "核实人"
        ws.cell(sr, 21).value = "单位盖章"
        for min_c, max_c in summary_merges:
            ws.merge_cells(start_row=sr, start_column=min_c, end_row=sr, end_column=max_c)
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(sr, col_idx)
            s = summary_styles.get(col_idx)
            if s:
                cell.font = s['font']
                cell.border = s['border']
                cell.alignment = s['alignment']
                if s['fill']:
                    cell.fill = s['fill']
                cell.number_format = s['number_format']
        ws.row_dimensions[sr].height = 30

        # 自适应列宽：H（证件号） L（地块编码）；A（编码）固定为19
        ws.column_dimensions['A'].width = 19
        for col_letter, col_idx in [('H', 8), ('L', 12)]:
            max_len = 0
            for row_cells in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                for cell in row_cells:
                    if cell.value:
                        val = str(cell.value)
                        length = sum(2 if ord(c) > 127 else 1 for c in val)
                        max_len = max(max_len, length)
            ws.column_dimensions[col_letter].width = max(max_len + 2, 8)

        wb.save(output_file)
        logger.info(f"金寨县调查信息表已保存到: {output_file}")


def generate_mdb_tables_only(mdb_file_path, shp_file_path, output_directory):
    """
    独立的摸底调查表生成功能，不依赖公示表生成过程
    直接处理MDB文件并生成摸底调查表
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating MDB tables from: {mdb_file_path}, {shp_file_path}")
    
    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')
    
    try:
        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)
        
        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}
            
            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]
            
            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')
            
            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)
            
            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')
            
            # 生成临时Excel文件作为数据源
            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)
            
            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # 调用openpyxl版本的摸底调查表生成功能
            generate_mdb_table(excel_file1, excel_file2, output_directory)
            
        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        # 清理临时文件夹
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")
    
    logger.debug("MDB table generation completed")


def generate_jinzhai_mdb_tables_only(mdb_file_path, shp_file_path, output_directory):
    """
    独立的金寨县摸底调查表生成功能，不依赖公示表生成过程
    直接处理MDB文件并生成金寨县摸底调查表
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating Jinzhai MDB tables from: {mdb_file_path}, {shp_file_path}")

    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

    try:
        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)

        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}

            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]

            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')

            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)

            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')

            # 生成临时Excel文件作为数据源
            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)

            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)

            # 调用金寨县摸底调查表生成功能
            generate_jinzhai_mdb_table(excel_file1, excel_file2, output_directory)

        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        # 清理临时文件夹
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")


def generate_jinzhai_survey_info_tables_only(mdb_file_path, shp_file_path, output_directory, region_prefix=None):
    """
    独立的金寨县土地承包经营权调查信息表生成功能
    直接处理MDB文件并生成调查信息表
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating Jinzhai survey info tables from: {mdb_file_path}, {shp_file_path}")

    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)

        try:
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}

            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]

            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')

            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)

            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')

            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)

            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)

            generate_jinzhai_survey_info_table(excel_file1, excel_file2, output_directory, region_prefix)

        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")


def generate_dkdc_table(family_info_path, land_info_path, output_directory, jtjjzz_name):
    """
    生成承包地块调查表（独立的函数，可不依赖其他表生成）
    接收excel文件路径参数，可以单独调用
    
    参数:
    family_info_path: 家庭信息excel文件路径
    land_info_path: 地块信息excel文件路径
    output_directory: 输出目录
    jtjjzz_name: 集体经济组织名称
    """
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'YHZGX': str, 'CBFBM': str, 'FBFBM': str, 'FBFMC': str, 'CBFMC': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str, 'CBFBM': str})  # 强制将DKBM读取为字符串

    # 去重处理，确保数据不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 获取CBFBM的唯一值列表
    unique_cbfbm = list(set(family_info['CBFBM'].unique()).union(set(land_info['CBFBM'].unique())))

    # 获取模板文件夹路径
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '承包地块调查表_南谯区.xlsx')

    for cbfbm in unique_cbfbm:
        # 过滤当前CBFBM的家庭成员信息和地块信息
        family_data = family_info[family_info['CBFBM'] == cbfbm]
        land_data = land_info[land_info['CBFBM'] == cbfbm]

        # 如果当前CBFBM没有有效数据，跳过
        if family_data.empty:
            continue

        # 获取承包方名称（CBFMC）
        cbfmc = family_data['CBFMC'].dropna().iloc[0] if 'CBFMC' in family_data.columns and not family_data['CBFMC'].dropna().empty else ''
        
        # 获取发包方信息
        fbfmc = family_data['FBFMC'].dropna().iloc[0] if 'FBFMC' in family_data.columns and not family_data['FBFMC'].dropna().empty else jtjjzz_name
        fbfbm = family_data['FBFBM'].dropna().iloc[0] if 'FBFBM' in family_data.columns and not family_data['FBFBM'].dropna().empty else ''
        # 截取"股份经济合作社"之后部分，但保留"股份经济合作社"
        if '股份经济合作社' in fbfmc:
            part_of_fbfmc = fbfmc.split('股份经济合作社')
            fbfmc_short= part_of_fbfmc[1] if len(part_of_fbfmc) > 1 else ""
        else:
            fbfmc_short = fbfmc
        # 生成文件名：CBFBM后6位_CBFMC承包地块调查表
        cbfbm_last6 = str(cbfbm)[-6:] if len(str(cbfbm)) >= 6 else str(cbfbm)  # 获取CBFBM后6位
        filename = f"{cbfbm}_{cbfmc}承包地块调查表.xlsx"
        # 获取发包方编码信息用于创建文件夹
        fbfbm_for_folder = family_data['FBFBM'].dropna().iloc[0] if 'FBFBM' in family_data.columns and not family_data['FBFBM'].dropna().empty else ''
        # 创建按村组分文件夹，文件夹名格式：FBFBM+FBFMC截取"股份经济合作社"之后部分
        folder_name = f"{fbfbm_for_folder}{fbfmc_short}"
        group_output_directory = os.path.join(output_directory, folder_name)
        os.makedirs(group_output_directory, exist_ok=True)

        output_file = os.path.join(group_output_directory, filename)

        # 复制模板文件
        shutil.copy2(template_path, output_file)

        # 打开工作簿
        wb = openpyxl.load_workbook(output_file)
        ws = wb['Sheet1']  # 使用Sheet1作为工作表

        # 填充固定信息
        # A2: FBFMC（发包方名称）
        ws['A2'] = jtjjzz_name if jtjjzz_name else fbfmc
        # B3: CBFMC（承包方名称）
        ws['B3'] = cbfmc
        # E3: CBFBM（承包方编码）
        ws['E3'] = str(cbfbm)
        # G3: FBFMC（发包方名称）
        ws['G3'] = jtjjzz_name if jtjjzz_name else fbfmc
        # L3: FBFBM（发包方编码）
        ws['L3'] = str(fbfbm) if fbfbm else ''

        # 获取地块信息
        land_rows = land_data
        m = len(land_rows)  # 地块数

        # 如果地块数>8，向excel表第6行之前插入【地块数-8】行
        if m > 8:
            rows_to_insert = m - 8
            for i in range(rows_to_insert):
                ws.insert_rows(6)
                # 复制第6行的样式到新插入的行
                source_row = 6 + rows_to_insert
                for col in range(1, 13):  # A-L列
                    source_cell = ws.cell(row=source_row, column=col)
                    target_cell = ws.cell(row=6 + i, column=col)
                    target_cell.font = source_cell.font.copy()
                    target_cell.border = source_cell.border.copy()
                    target_cell.fill = source_cell.fill.copy()
                    target_cell.number_format = source_cell.number_format
                    target_cell.protection = source_cell.protection.copy()
                    target_cell.alignment = source_cell.alignment.copy()
                    # 清空目标单元格的值
                    target_cell.value = None

        # 计算HTMJM汇总值
        total_htmjm = round(land_rows['HTMJM'].sum(), 2) if not land_rows.empty else 0.00

        # I列合并单元格
        if m > 8:
            # 地块数>8：startrow:6;endrow:地块数+6-1
            end_row = m + 6 - 1
        else:
            # 地块数≤8：startrow:6;endrow:8+6-1
            end_row = 8 + 6 - 1
        ws.merge_cells(start_row=6, start_column=9, end_row=end_row, end_column=9)  # I列合并
        ws.merge_cells(start_row=6, start_column=12, end_row=end_row, end_column=12)  # L列合并

        # 在I列的合并单元格中填入汇总值
        ws.cell(row=6, column=9).value = total_htmjm

        # 从endrow+1行开始插入4行，添加调查记事和审核意见
        if m > 8:
            # 地块数>8：startrow:6;endrow:地块数+6-1
            end_row = m + 6 - 1
        else:
            # 地块数≤8：startrow:6;endrow:8+6-1
            end_row = 8 + 6 - 1
        
        # 从endrow+1行开始插入4行
        for i in range(4):
            ws.insert_rows(end_row + 1)
        
        # 计算调查记事和审核意见的行号
        investigation_start = end_row + 1
        investigation_end = end_row + 2
        audit_start = end_row + 3
        audit_end = end_row + 4
        
        # 设置行高为36
        for row_num in [investigation_start, investigation_end, audit_start, audit_end]:
            ws.row_dimensions[row_num].height = 36
        
        # A列，investigation_start和investigation_end合并单元格，值设置为"调查记事"
        ws.merge_cells(start_row=investigation_start, start_column=1, end_row=investigation_end, end_column=1)
        ws.cell(row=investigation_start, column=1).value = "调查记事"
        ws.cell(row=investigation_start, column=1).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=investigation_start, column=1).font = Font(name='宋体', size=11, bold=True)
        
        # A列，audit_start和audit_end合并单元格，值设置为"审核意见"
        ws.merge_cells(start_row=audit_start, start_column=1, end_row=audit_end, end_column=1)
        ws.cell(row=audit_start, column=1).value = "审核意见"
        ws.cell(row=audit_start, column=1).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=audit_start, column=1).font = Font(name='宋体', size=11, bold=True)
        
        # 为A列的调查记事和审核意见单元格添加边框
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        ws.cell(row=investigation_start, column=1).border = border
        ws.cell(row=audit_start, column=1).border = border
        ws.cell(row=investigation_end, column=1).border = border
        ws.cell(row=audit_end, column=1).border = border
        # investigation_start行，B列-L列合并单元格，设置"经调查，情况属实"，居中对齐
        ws.merge_cells(start_row=investigation_start, start_column=2, end_row=investigation_start, end_column=12)
        ws.cell(row=investigation_start, column=2).value = "经调查，情况属实"
        ws.cell(row=investigation_start, column=2).alignment = Alignment(horizontal='center', vertical='center')
        
        # investigation_end行，B列—L列合并单元格，行高设置为28，值为"调查员："
        ws.row_dimensions[investigation_end].height = 28
        ws.merge_cells(start_row=investigation_end, start_column=2, end_row=investigation_end, end_column=12)
        ws.cell(row=investigation_end, column=2).value = "调查员：                                                                                            日期：       年   月   日"
        
        # audit_start行，B列—L列合并单元格
        ws.merge_cells(start_row=audit_start, start_column=2, end_row=audit_start, end_column=12)
        
        # audit_end行，B列—L列合并单元格，行高设置为28，值为"审核人："
        ws.row_dimensions[audit_end].height = 28
        ws.merge_cells(start_row=audit_end, start_column=2, end_row=audit_end, end_column=12)
        ws.cell(row=audit_end, column=2).value = "审核人：                                                                                            日期：       年   月   日"
        
        # 设置调查记事和审核意见的字体样式
        font = Font(name='宋体', size=11)
        alignment_left = Alignment(horizontal='left', vertical='center')
        alignment_center = Alignment(horizontal='center', vertical='center')
        
        for row_num in [investigation_start, investigation_end, audit_start, audit_end]:
            for col_num in range(1, 13):  # A-L列
                cell = ws.cell(row=row_num, column=col_num)
                cell.font = font
                if col_num == 1:  # A列左对齐
                    cell.alignment = alignment_center
                elif col_num == 2:  # B列
                    if row_num == investigation_start:
                        cell.alignment = alignment_center
                    else:
                        cell.alignment = alignment_left
                else:
                    cell.alignment = alignment_left

        # 从第6行开始填充地块数据
        for i in range(m):
            row_num = 6 + i
            if i < len(land_rows):
                land = land_rows.iloc[i]
                # A列：DKMC（地块名称）
                ws.cell(row=row_num, column=1).value = land['DKMC'] if 'DKMC' in land else ''
                # B列：DKBM（地块编码）
                ws.cell(row=row_num, column=2).value = str(land['DKBM']) if 'DKBM' in land else ''
                # D列：DKDZ（东至）
                ws.cell(row=row_num, column=4).value = land['DKDZ'] if 'DKDZ' in land else ''
                # E列：DKNZ（南至）
                ws.cell(row=row_num, column=5).value = land['DKNZ'] if 'DKNZ' in land else ''
                # F列：DKXZ（西至）
                ws.cell(row=row_num, column=6).value = land['DKXZ'] if 'DKXZ' in land else ''
                # G列：DKBZ（北至）
                ws.cell(row=row_num, column=7).value = land['DKBZ'] if 'DKBZ' in land else ''
                # H列：HTMJM（合同面积亩）
                ws.cell(row=row_num, column=8).value = land['HTMJM'] if 'HTMJM' in land else ''
                # J列：TDYT（土地用途）
                ws.cell(row=row_num, column=10).value = '种植业'
                # I列：不填入值（因为是合并单元格，汇总值已经填入第一行）

                # B列和C列合并单元格
                ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=3)

                # 设置单元格样式：字体大小11、居中对齐
                font = Font(name='宋体', size=11)
                alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                for col in range(1, 11):  # A-J列
                    cell = ws.cell(row=row_num, column=col)
                    cell.font = font
                    cell.alignment = alignment
                    # 设置关键字段为文本格式
                    if col == 2:  # B列：DKBM（地块编码）
                        cell.number_format = '@'

                # 行高设置为28（最小值）
                ws.row_dimensions[row_num].height = 28

        # 为所有单元格添加边框
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 遍历工作表中的所有单元格并添加边框
        for row in ws.iter_rows(min_row=1, max_row=end_row):
            for cell in row:
                cell.border = border

        # 为调查记事和审核意见部分设置特殊边框（仅外侧边线）
        # investigation_start-investigation_end行，B-L列设置外侧边框
        for row_num in [investigation_start, investigation_end]:
            for col_num in range(2, 13):  # B-L列
                cell = ws.cell(row=row_num, column=col_num)
                # 只设置外侧边框
                cell.border = Border(
                    left=Side(style='thin') if col_num == 2 else None,
                    right=Side(style='thin') if col_num == 12 else None,
                    top=Side(style='thin') if row_num == investigation_start else None,
                    bottom=Side(style='thin') if row_num == investigation_end else None
                )

        # audit_start-audit_end行，B-L列设置外侧边框
        for row_num in [audit_start, audit_end]:
            for col_num in range(2, 13):  # B-L列
                cell = ws.cell(row=row_num, column=col_num)
                # 只设置外侧边框
                cell.border = Border(
                    left=Side(style='thin') if col_num == 2 else None,
                    right=Side(style='thin') if col_num == 12 else None,
                    top=Side(style='thin') if row_num == audit_start else None,
                    bottom=Side(style='thin') if row_num == audit_end else None
                )

        # 保存工作簿
        wb.save(output_file)
        print(f"承包地块调查表已保存到: {output_file}")


def generate_dkdc_tables_only(mdb_file_path, shp_file_path, output_directory, jtjjzz_name):
    """
    独立的承包地块调查表生成功能，不依赖其他表生成过程
    直接处理MDB文件并生成承包地块调查表
    
    参数:
    mdb_file_path: MDB数据库文件路径
    shp_file_path: SHP矢量文件路径
    output_directory: 输出目录
    jtjjzz_name: 集体经济组织名称
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating DKDC tables from: {mdb_file_path}, {shp_file_path}")
    
    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')
    
    try:
        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)
        
        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}
            
            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]
            
            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ','CBFZJHM','CBFCYSL']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')
            
            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)
            
            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')
            
            # 生成临时Excel文件作为数据源
            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)
            
            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # 调用承包地块调查表生成功能
            generate_dkdc_table(excel_file1, excel_file2, output_directory, jtjjzz_name)
            
        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        # 清理临时文件夹
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")
    
    logger.debug("DKDC table generation completed")

    logger.debug("Jinzhai MDB table generation completed")


def generate_hzzz_sms(family_info_path, land_info_path, output_directory):
    """
    生成户主代表声明书（独立的函数，可不依赖其他表生成）
    接收excel文件路径参数，可以单独调用
    
    参数:
    family_info_path: 家庭信息excel文件路径
    land_info_path: 地块信息excel文件路径
    output_directory: 输出目录
    """
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 检查Document是否可用
    if Document is None:
        logger.error("python-docx库未安装，无法生成户主代表声明书")
        return

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'YHZGX': str, 'CYXB': str, 'CBFBM': str, 'FBFMC': str, 'CBFMC': str, 'CYZJHM': str, 'CBFZJHM': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str, 'CBFBM': str})  # 强制将DKBM读取为字符串

    # 去重处理，确保数据不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 获取CBFBM的唯一值列表
    unique_cbfbm = list(set(family_info['CBFBM'].unique()).union(set(land_info['CBFBM'].unique())))

    # 获取模板文件夹路径
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '户主代表声明书_南谯区.docx')

    for cbfbm in unique_cbfbm:
        # 过滤当前CBFBM的家庭成员信息和地块信息
        family_data = family_info[family_info['CBFBM'] == cbfbm]
        land_data = land_info[land_info['CBFBM'] == cbfbm]

        # 如果当前CBFBM没有有效数据，跳过
        if family_data.empty:
            continue

        # 获取承包方名称（CBFMC）
        cbfmc = family_data['CBFMC'].dropna().iloc[0] if 'CBFMC' in family_data.columns and not family_data['CBFMC'].dropna().empty else ''

        # 获取户主信息（YHZGX为01或02或户主的记录）
        household_info = family_data[family_data['YHZGX'].isin(['01', '02'])]
        if not household_info.empty:
            # 获取户主的性别
            cyxb = household_info['CYXB'].iloc[0] if 'CYXB' in household_info.columns else ''
            # 性别转换
            if str(cyxb) == '1':
                cyxb_name = '男'
            elif str(cyxb) == '2':
                cyxb_name = '女'
            else:
                cyxb_name = ''
        else:
            cyxb_name = ''

        # 获取发包方信息
        fbfmc = family_data['FBFMC'].dropna().iloc[0] if 'FBFMC' in family_data.columns and not family_data['FBFMC'].dropna().empty else ''
        # 截取"股份经济合作社"之后部分，但保留"股份经济合作社"
        if '股份经济合作社' in fbfmc:
            parts = fbfmc.split('股份经济合作社')
            fbfmc_short = parts[0]
            fbfmc_short1=parts[1] if len(parts)>1 else ''
        else:
            fbfmc_short = fbfmc
            fbfmc_short1 = ''  # 确保变量在所有情况下都被定义

        # 获取承包方证件号码
        cbfzm = ''
        if 'CBFZJHM' in household_info.columns and not household_info['CBFZJHM'].dropna().empty:
            cbfzm = str(household_info['CBFZJHM'].dropna().iloc[0])
        elif 'CYZJHM' in household_info.columns and not household_info['CYZJHM'].dropna().empty:
            cbfzm = str(household_info['CYZJHM'].dropna().iloc[0])

        # 获取家庭成员姓名列表，以"、"分割
        jtcy_names = family_data['CYXM'].dropna().tolist()
        jtcy = '、'.join(jtcy_names)

        # 计算家庭成员数量
        cbfcysl = len(family_data)

        # 获取发包方编码信息用于创建文件夹
        fbfbm_for_folder = family_data['FBFBM'].dropna().iloc[0] if 'FBFBM' in family_data.columns and not family_data['FBFBM'].dropna().empty else ''
        
        # 创建按村组分文件夹，文件夹名格式：FBFBM+FBFMC截取"股份经济合作社"之后部分
        folder_name = f"{fbfbm_for_folder}{fbfmc_short1}"
        group_output_directory = os.path.join(output_directory, folder_name)
        os.makedirs(group_output_directory, exist_ok=True)

        # 生成文件名：{CBFBM+CBFMC-户主代表声明书.docx}
        filename = f"{cbfbm}{cbfmc}-户主代表声明书.docx"
        output_file = os.path.join(group_output_directory, filename)

        try:
            # 检查模板文件格式，只支持.docx
            if not template_path.lower().endswith('.docx'):
                logger.error(f"模板文件格式错误，只支持.docx格式，当前文件：{template_path}")
                # 创建错误说明文档
                error_file = os.path.join(output_directory, f"{cbfbm}{cbfmc}-户主代表声明书-错误说明.txt")
                with open(error_file, 'w', encoding='utf-8') as f:
                    f.write(f"户主代表声明书生成失败\n")
                    f.write(f"错误原因：模板文件格式不兼容\n")
                    f.write(f"当前模板：{template_path}\n")
                    f.write(f"解决方案：请将模板文件转换为.docx格式\n")
                    f.write(f"将文件重命名为：Templates/户主代表声明书_凤阳县.docx\n")
                print(f"户主代表声明书错误说明已保存到: {error_file}")
                return
            
            # 复制模板文件
            import shutil
            temp_doc_path = os.path.join(output_directory, f"temp_{filename}")
            shutil.copy2(template_path, temp_doc_path)
            
            try:
                # 打开Word文档进行编辑
                doc = Document(temp_doc_path)
                
                # 替换文档中的标签，保持原有格式
                # 遍历所有run（具有相同格式的文本段）
                for paragraph in doc.paragraphs:
                    for run in paragraph.runs:
                        if 'CBFMC' in run.text:
                            run.text = run.text.replace('CBFMC', cbfmc)
                        if 'CYXB' in run.text:
                            run.text = run.text.replace('CYXB', cyxb_name)
                        if 'FBFMC' in run.text:
                            run.text = run.text.replace('FBFMC', fbfmc_short)
                        if 'CBFZJHM' in run.text:
                            run.text = run.text.replace('CBFZJHM', cbfzm)
                        if 'JTCY' in run.text:
                            run.text = run.text.replace('JTCY', jtcy)
                        if 'CBFCYSL' in run.text:
                            run.text = run.text.replace('CBFCYSL', str(cbfcysl))
                
                # 替换表格中的标签，保持原有格式
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    if 'CBFMC' in run.text:
                                        run.text = run.text.replace('CBFMC', cbfmc)
                                    if 'CYXB' in run.text:
                                        run.text = run.text.replace('CYXB', cyxb_name)
                                    if 'FBFMC' in run.text:
                                        run.text = run.text.replace('FBFMC', fbfmc_short)
                                    if 'CBFZJHM' in run.text:
                                        run.text = run.text.replace('CBFZJHM', cbfzm)
                                    if 'JTCY' in run.text:
                                        run.text = run.text.replace('JTCY', jtcy)
                                    if 'CBFCYSL' in run.text:
                                        run.text = run.text.replace('CBFCYSL', str(cbfcysl))
                
                # 保存修改后的文档
                doc.save(output_file)
                
                # 删除临时文件
                if os.path.exists(temp_doc_path):
                    os.remove(temp_doc_path)
                
                print(f"户主代表声明书已保存到: {output_file}")
                
            except Exception as e:
                logger.error(f"处理Word文档时发生错误 {output_file}: {str(e)}")
                # 如果处理失败，尝试删除临时文件
                if os.path.exists(temp_doc_path):
                    os.remove(temp_doc_path)
                raise
                
        except Exception as e:
            logger.error(f"处理户主代表声明书时发生错误: {str(e)}", exc_info=True)
            raise


def generate_hzzz_sms_only(mdb_file_path, shp_file_path, output_directory):
    """
    独立的户主代表声明书生成功能，不依赖其他表生成过程
    直接处理MDB文件并生成户主代表声明书

    参数:
    mdb_file_path: MDB数据库文件路径
    shp_file_path: SHP矢量文件路径
    output_directory: 输出目录
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating HZZZ SMS from: {mdb_file_path}, {shp_file_path}")

    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

    try:
        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)

        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}

            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]

            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ','CBFZJHM','CBFCYSL']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')

            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)

            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')

            # 生成临时Excel文件作为数据源
            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)

            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)

            # 调用户主代表声明书生成功能
            generate_hzzz_sms(excel_file1, excel_file2, output_directory)

        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        # 清理临时文件夹
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")

    logger.debug("HZZZ SMS generation completed")


def generate_yhzd_cover(family_info_path, land_info_path, output_directory):
    """
    生成一户一档封面（独立的函数，可不依赖其他表生成）
    接收excel文件路径参数，可以单独调用

    参数:
    family_info_path: 家庭信息excel文件路径
    land_info_path: 地块信息excel文件路径
    output_directory: 输出目录
    """
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 检查Document是否可用
    if Document is None:
        logger.error("python-docx库未安装，无法生成一户一档封面")
        return

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'YHZGX': str, 'CBFBM': str, 'FBFMC': str, 'CBFMC': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str, 'CBFBM': str})  # 强制将DKBM读取为字符串

    # 去重处理，确保数据不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 获取CBFBM的唯一值列表
    unique_cbfbm = list(set(family_info['CBFBM'].unique()).union(set(land_info['CBFBM'].unique())))

    # 获取模板文件夹路径
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '一户一档封面_南谯区.docx')

    for cbfbm in unique_cbfbm:
        # 过滤当前CBFBM的家庭成员信息和地块信息
        family_data = family_info[family_info['CBFBM'] == cbfbm]
        land_data = land_info[land_info['CBFBM'] == cbfbm]

        # 如果当前CBFBM没有有效数据，跳过
        if family_data.empty:
            continue

        # 获取承包方名称（CBFMC）
        cbfmc = family_data['CBFMC'].dropna().iloc[0] if 'CBFMC' in family_data.columns and not family_data['CBFMC'].dropna().empty else ''

        # 获取发包方信息
        fbfmc = family_data['FBFMC'].dropna().iloc[0] if 'FBFMC' in family_data.columns and not family_data['FBFMC'].dropna().empty else ''

        # 从FBFMC中解析镇乡名称
        def extract_town_name(fbfmc):
            """从FBFMC中解析镇乡名称和村名称"""
            if not fbfmc:
                return '', ''

            # 查找"镇"或"乡"的位置
            town_keywords = ['镇', '乡','街道']
            village_keywords = ['村', '社区']

            town_name = ''
            village_name = ''

            # 先找镇乡
            for keyword in town_keywords:
                if keyword in fbfmc:
                    parts = fbfmc.split(keyword)
                    if len(parts) > 1:
                        town_part = parts[0]
                        if town_part.strip():
                            town_name = town_part.strip() + keyword
                            remaining = keyword.join(parts[1:])
                            break

            # 再找村
            for keyword in village_keywords:
                if keyword in remaining:
                    parts = remaining.split(keyword)
                    if len(parts) > 1:
                        village_part = parts[0]
                        if village_part.strip():
                            village_name = village_part.strip() + keyword
                            break

            return town_name, village_name

        town_name, village_name = extract_town_name(fbfmc)
        xzmc = town_name.replace('滁州市南谯区','').replace('安徽省滁州市南谯区','')
        cmc = village_name
        # 获取发包方编码信息用于创建文件夹
        fbfbm_for_folder = family_data['FBFBM'].dropna().iloc[0] if 'FBFBM' in family_data.columns and not family_data['FBFBM'].dropna().empty else ''

        # 创建按村组分文件夹，文件夹名格式：FBFBM+FBFMC截取"股份经济合作社"之后部分
        if '股份经济合作社' in fbfmc:
            parts = fbfmc.split('股份经济合作社')
            fbfmc_short = parts[1] if len(parts) > 1 else ""
        else:
            fbfmc_short = fbfmc

        folder_name = f"{fbfbm_for_folder}{fbfmc_short}"
        group_output_directory = os.path.join(output_directory, folder_name)
        os.makedirs(group_output_directory, exist_ok=True)

        # 生成文件名：{CBFBM+CBFMC-一户一档封面.docx}
        filename = f"{cbfbm}{cbfmc}-一户一档封面.docx"
        output_file = os.path.join(group_output_directory, filename)

        try:
            # 检查模板文件格式，只支持.docx
            if not template_path.lower().endswith('.docx'):
                logger.error(f"模板文件格式错误，只支持.docx格式，当前文件：{template_path}")
                # 创建错误说明文档
                error_file = os.path.join(output_directory, f"{cbfbm}{cbfmc}-一户一档封面-错误说明.txt")
                with open(error_file, 'w', encoding='utf-8') as f:
                    f.write(f"一户一档封面生成失败\n")
                    f.write(f"错误原因：模板文件格式不兼容\n")
                    f.write(f"当前模板：{template_path}\n")
                    f.write(f"解决方案：请将模板文件转换为.docx格式\n")
                    f.write(f"将文件重命名为：Templates/一户一档封面_南谯区.docx\n")
                print(f"一户一档封面错误说明已保存到: {error_file}")
                return

            # 复制模板文件
            import shutil
            temp_doc_path = os.path.join(output_directory, f"temp_{filename}")
            shutil.copy2(template_path, temp_doc_path)

            try:
                # 打开Word文档进行编辑
                doc = Document(temp_doc_path)

                # 替换文档中的标签，保持原有格式
                # 遍历所有run（具有相同格式的文本段）
                for paragraph in doc.paragraphs:
                    for run in paragraph.runs:
                        if 'CBFMC' in run.text:
                            run.text = run.text.replace('CBFMC', cbfmc)
                        if 'XZMC' in run.text:
                            run.text = run.text.replace('XZMC', xzmc)
                        if 'FBFMC' in run.text:
                            run.text = run.text.replace('FBFMC', xzmc+cmc)

                # 替换表格中的标签，保持原有格式
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    if 'CBFMC' in run.text:
                                        run.text = run.text.replace('CBFMC', cbfmc)
                                    if 'XZMC' in run.text:
                                        run.text = run.text.replace('XZMC', xzmc)
                                    if 'FBFMC' in run.text:
                                        run.text = run.text.replace('FBFMC', xzmc+cmc)

                # 保存修改后的文档
                doc.save(output_file)

                # 删除临时文件
                if os.path.exists(temp_doc_path):
                    os.remove(temp_doc_path)

                print(f"一户一档封面已保存到: {output_file}")

            except Exception as e:
                logger.error(f"处理Word文档时发生错误 {output_file}: {str(e)}")
                # 如果处理失败，尝试删除临时文件
                if os.path.exists(temp_doc_path):
                    os.remove(temp_doc_path)
                raise

        except Exception as e:
            logger.error(f"处理一户一档封面时发生错误: {str(e)}", exc_info=True)
            raise


def generate_yhzd_cover_only(mdb_file_path, shp_file_path, output_directory):
    """
    独立的一户一档封面生成功能，不依赖其他表生成过程
    直接处理MDB文件并生成一户一档封面

    参数:
    mdb_file_path: MDB数据库文件路径
    shp_file_path: SHP矢量文件路径
    output_directory: 输出目录
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating YHZD Cover from: {mdb_file_path}, {shp_file_path}")

    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

    try:
        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)

        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}

            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]

            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ','CBFZJHM','CBFCYSL']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')

            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)

            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')

            # 生成临时Excel文件作为数据源
            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)

            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)

            # 调用一户一档封面生成功能
            generate_yhzd_cover(excel_file1, excel_file2, output_directory)

        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        # 清理临时文件夹
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")

    logger.debug("YHZD Cover generation completed")


def generate_jinzhai_gsb_only(mdb_file_path, shp_file_path, output_directory, jtjjzz_name):
    """
    独立的金寨县公示表生成功能，不依赖其他表生成过程
    直接处理MDB文件并生成金寨县公示表
    
    参数:
    mdb_file_path: MDB数据库文件路径
    shp_file_path: SHP矢量文件路径
    output_directory: 输出目录
    jtjjzz_name: 集体经济组织名称
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating Jinzhai GSB from: {mdb_file_path}, {shp_file_path}")
    
    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')
    
    try:
        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)
        
        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}
            
            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]
            
            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')
            
            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)
            
            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')
            
            # 生成临时Excel文件作为数据源
            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)
            
            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # 调用金寨县公示表生成功能
            merge_jinzhai_excel_files(excel_file1, excel_file2, output_directory, jtjjzz_name)
            
        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        # 清理临时文件夹
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")
    
    logger.debug("Jinzhai GSB generation completed")


def generate_yushe_mdb_table(family_info_path, land_info_path, output_directory):
    """
    生成榆社县摸底调查表（独立的函数，可不依赖其他表生成）
    接收excel文件路径参数，可以单独调用
    
    参数:
    family_info_path: 家庭信息excel文件路径
    land_info_path: 地块信息excel文件路径
    output_directory: 输出目录
    """
    def safe_read_excel(file_path, converters=None):
        """读取Excel文件，尝试不同引擎"""
        try:
            # 首先尝试使用openpyxl引擎
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {'YHZGX': str})
        except zipfile.BadZipFile:
            # 如果xlsx文件损坏，检查文件是否为xls格式
            if os.path.exists(file_path):
                try:
                    # 尝试xlrd引擎（用于Excel 2003 .xls文件）
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {'YHZGX': str})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")

            # 提供有用的错误信息
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            msg = f"Excel文件读取失败!\n\n文件: {file_path}\n"
            msg += "可能的錯誤原因:\n"
            msg += "1. 文件损坏或格式错误\n"
            msg += "2. 文件不是有效的Excel文件\n"
            msg += "3. 安装了xlrd包可能可以读取xls格式文件\n"
            msg += "請檢查文件是否有損壞，或者嘗試用其他Excel軟體打開驗證。"
            raise ValueError(msg)
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    # 读取Excel文件
    family_info = safe_read_excel(family_info_path, converters={'YHZGX': str, 'CBFBM': str, 'FBFBM': str, 'FBFMC': str, 'CBFMC': str})
    land_info = safe_read_excel(land_info_path, converters={'DKBM': str, 'CBFBM': str})  # 强制将DKBM读取为字符串

    # 去重处理，确保数据不重复
    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    # 按FBFBM（发包方编码）分组处理
    family_info['FBFBM_14'] = family_info['CBFBM'].str[:14]  # 提取前14位作为FBFBM
    fbf_grouped = family_info.groupby('FBFBM_14')

    # 获取模板文件路径
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '摸底调查表模版_榆社县.xlsx')

    # 按发包方分组处理
    for fbfbm, fbf_family_data in fbf_grouped:
        # 获取发包方名称
        fbfmc = fbf_family_data['FBFMC'].iloc[0] if 'FBFMC' in fbf_family_data.columns and not fbf_family_data['FBFMC'].empty else ''
        
        # 创建发包方文件夹
        fbf_folder_name = f"{fbfbm}_{fbfmc}"
        # 清理文件夹名称中的非法字符
        fbf_folder_name = fbf_folder_name.replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?', '').replace('"', '').replace('<', '').replace('>', '').replace('|', '')
        fbf_directory = os.path.join(output_directory, fbf_folder_name)
        os.makedirs(fbf_directory, exist_ok=True)
        
        # 按CBFBM分组处理当前发包方的承包方
        family_grouped = fbf_family_data.groupby('CBFBM')
        land_grouped = land_info.groupby('CBFBM')
        
        for cbfbm, family_data in family_grouped:
            # 获取当前承包方的基本信息
            cbf_info = family_data.iloc[0]
            cbf_name = cbf_info.get('CBFMC', '')
            cbfdz = cbf_info.get('CBFDZ', '')
            lxdh = cbf_info.get('LXDH_x', '')
            cbqxq = cbf_info.get('CBQXQ', '')
            cbqxz = cbf_info.get('CBQXZ', '')
            
            # 计算承包期限
            if pd.notna(cbqxq) and pd.notna(cbqxz):
                cbqxq_str = cbqxq.strftime('%Y年%m月%d日') if hasattr(cbqxq, 'strftime') else str(cbqxq)
                cbqxz_str = cbqxz.strftime('%Y年%m月%d日') if hasattr(cbqxz, 'strftime') else str(cbqxz)
                contract_term = f"{cbqxq_str} 至 {cbqxz_str}"
            else:
                contract_term = ""
            
            # 计算成员个数
            member_count = len(family_data)
            
            # 获取当前承包方的地块信息
            land_data = land_grouped.get_group(cbfbm) if cbfbm in land_grouped.groups else pd.DataFrame()
            land_count = len(land_data)
            # 计算合同总面积
            total_htmjm = land_data.get('HTMJM', pd.Series()).sum() if not land_data.empty else 0
            
            # 读取模板文件
            wb = openpyxl.load_workbook(template_path)
            
            # 处理承包方家庭成员摸底表
            for sheet_name in wb.sheetnames:
                if '承包方家庭成员摸底表' in sheet_name:
                    ws = wb[sheet_name]
                    # 计算页数：成员个数/6 向上取整
                    import math
                    member_pages = math.ceil(member_count / 6)
                    
                    # 填充每页的内容
                    for i in range(member_pages):
                        # 计算当前页的起始行
                        start_row = 20 * i
                        
                        # 填充页面头部信息
                        ws[f'A{start_row + 3}'] = f"发包方名称：{fbfmc}"  # 发包方名称
                        ws[f'D{start_row + 4}'] = cbf_name  # 承包方名称
                        ws[f'F{start_row + 4}'] = cbfdz  # 承包方地址
                        ws[f'H{start_row + 4}'] = lxdh  # 联系电话
                        ws[f'J{start_row + 4}'] = contract_term  # 承包期限
                        ws[f'G{start_row + 5}'] = f"{cbfbm}J"  # CBFBM+J
                        ws[f'L{start_row + 5}'] = f"{cbfbm}J"  # CBFBM+J
                        ws[f'C{start_row + 7}'] = f"{member_count}人"  # 成员个数+“人”
                        
                        # 填充家庭成员信息
                        start_member_row = start_row + 7
                        member_index = i * 6
                        
                        for j in range(6):
                            if member_index + j < member_count:
                                member = family_data.iloc[member_index + j]
                                row = start_member_row + j
                                ws[f'B{row}'] = member_index + j + 1  # 序号，递增1
                                ws[f'D{row}'] = member.get('CYXM', '')  # 姓名
                                ws[f'E{row}'] = xb_code_to_name(member.get('CYXB', ''))  # 性别
                                ws[f'F{row}'] = member.get('CYZJHM', '')  # 身份证号
                                ws[f'G{row}'] = yhzgx_code_to_name(member.get('YHZGX', ''))  # 与户主关系
                                ws[f'H{row}'] = "是"  # 是
                                ws[f'I{row}'] = member.get('CYBZSM', '') # 原二轮承包共有人
                    
                    # 删除多余的页
                    max_needed_rows = 20 * member_pages
                    start = max_needed_rows
                    end = ws.max_row
                    if start <= end:
                        ws.delete_rows(start, end - start + 1)
                    
                    # # 删除末尾的空白行
                    # row_to_check = ws.max_row
                    # while row_to_check > 0:
                    #     # 检查当前行是否为空（所有单元格都为None或空字符串或只包含空白字符）
                    #     is_empty = True
                    #     for cell in ws[row_to_check]:
                    #         if cell.value is not None and str(cell.value).strip() != '':
                    #             is_empty = False
                    #             break
                    #     if is_empty:
                    #         ws.delete_rows(row_to_check)
                    #         row_to_check -= 1
                    #     else:
                    #         break
                        
            # 处理承包方承包地块摸底表
            for sheet_name in wb.sheetnames:
                if '承包方承包地块摸底表' in sheet_name:
                    ws = wb[sheet_name]
                    # 计算页数：地块个数/6 向上取整
                    import math
                    land_pages = math.ceil(land_count / 6)
                    
                    # 填充每页的内容
                    for i in range(land_pages):
                        # 计算当前页的起始行
                        start_row = 20 * i
                        
                        # 填充页面头部信息
                        ws[f'A{start_row + 2}'] = f"发包方名称：{fbfmc}"  # 发包方名称
                        ws[f'D{start_row + 3}'] = cbf_name  # 承包方名称
                        ws[f'F{start_row + 3}'] = lxdh  # 联系电话
                        ws[f'K{start_row + 3}'] = total_htmjm  # 合同总面积亩
                        ws[f'N{start_row + 3}'] = land_count  # 地块总数
                        
                        # 填充地块信息
                        start_land_row = start_row + 6
                        land_index = i * 6
                        
                        for j in range(6):
                            if land_index + j < land_count:
                                land = land_data.iloc[land_index + j]
                                row = start_land_row + j
                                ws[f'B{row}'] = land_index + j + 1  # 序号，递增1
                                ws[f'C{row}'] = land.get('DKMC', '')  # 地块名称
                                ws[f'D{row}'] = land.get('DKBM', '')  # 地块编码
                                ws[f'E{row}'] = land.get('HTMJM', '')  # 合同面积亩
                                ws[f'G{row}'] = land.get('DKDZ', '')  # 地块东至
                                ws[f'H{row}'] = land.get('DKNZ', '')  # 地块南至
                                ws[f'I{row}'] = land.get('DKXZ', '')  # 地块西至
                                ws[f'J{row}'] = land.get('DKBZ', '')  # 地块北至
                    
                    # 删除多余的页
                    max_needed_rows = 20 * land_pages
                    start = max_needed_rows + 1
                    end = ws.max_row
                    if start <= end:
                        ws.delete_rows(start, end - start + 1)
                    
                    # # 删除末尾的空白行
                    # row_to_check = ws.max_row
                    # while row_to_check > 0:
                    #     # 检查当前行是否为空（所有单元格都为None或空字符串）
                    #     is_empty = True
                    #     for cell in ws[row_to_check]:
                    #         if cell.value is not None and str(cell.value).strip() != '':
                    #             is_empty = False
                    #             break
                    #     if is_empty:
                    #         ws.delete_rows(row_to_check)
                    #         row_to_check -= 1
                    #     else:
                    #         break
                               
            # 保存文件到发包方文件夹
            output_file = os.path.join(fbf_directory, f"{cbfbm}{cbf_name}_摸底调查表.xlsx")
            wb.save(output_file)
            print(f"榆社县摸底调查表处理完成，已保存到 {output_file}")


def generate_yushe_mdb_tables_only(mdb_file_path, shp_file_path, output_directory):
    """
    独立的榆社县摸底调查表生成功能，不依赖公示表生成过程
    直接处理MDB文件并生成榆社县摸底调查表
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating Yushe MDB tables from: {mdb_file_path}, {shp_file_path}")
    
    # 创建临时文件夹
    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')
    
    try:
        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'  
            r'DBQ=' + mdb_file_path + ';'  
        )
        conn = pyodbc.connect(conn_str)
        
        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}
            
            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]
            
            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')
            
            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)
            
            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')
            
            # 生成临时Excel文件作为数据源
            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)
            
            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # 调用榆社县摸底调查表生成功能
            generate_yushe_mdb_table(excel_file1, excel_file2, output_directory)
            
        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        # 清理临时文件夹
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")
    
    logger.debug("Yushe MDB table generation completed")


def _calc_age_from_id(zjhm):
    """
    根据证件号码推算年龄（周岁）。
    支持18位（出生日期为第7-14位）和15位（出生日期为第7-12位，年份前加19）证件号码。
    """
    if not zjhm:
        return ''
    zjhm = str(zjhm).strip()
    birth = ''
    if len(zjhm) == 18:
        birth = zjhm[6:14]
    elif len(zjhm) == 15:
        birth = '19' + zjhm[6:12]
    if len(birth) == 8 and birth.isdigit():
        from datetime import date
        year, month, day = int(birth[:4]), int(birth[4:6]), int(birth[6:8])
        today = date.today()
        age = today.year - year
        if (today.month, today.day) < (month, day):
            age -= 1
        return str(age) if 0 <= age <= 130 else ''
    return ''


def parse_yushe_fbfmc(fbfmc):
    """
    解析榆社县发包方名称，返回 (乡镇名, 村名, 组名)。
    乡镇名以 乡/镇/街道 结尾，村名以 村 结尾；
    剩余部分剥离开 股份经济合作社/经济合作社/合作社/村民委员会/村委会 后缀后，
    仅当包含"组"或以"队""社"结尾时才作为组名，否则返回空字符串。
    """
    if not fbfmc:
        return '', '', ''
    name = str(fbfmc).strip()
    if name.lower() == 'nan':
        return '', '', ''
    zhen = ''
    for sep in ('乡', '镇', '街道'):
        idx = name.find(sep)
        if idx != -1:
            zhen = name[:idx + len(sep)]
            name = name[idx + len(sep):]
            break
    cun = ''
    idx = name.find('村')
    if idx != -1:
        cun = name[:idx + 1]
        name = name[idx + 1:]
    for suffix in ('股份经济合作社', '集体经济组织','经济合作社', '合作社', '村民委员会', '村委会'):
        if name.endswith(suffix):
            name = name[:-len(suffix)]
            break
    zumin = ''
    if name and ('组' in name or name.endswith(('队', '社'))):
        zumin = name
    return zhen, cun, zumin


def _strip_admin_suffix(name):
    """去掉名称末尾的 街道/乡/镇/村/组 等行政区划后缀（最长优先），剥空则返回原值"""
    s = str(name)
    for suf in ('街道', '乡', '镇', '村', '组'):
        if s.endswith(suf) and len(s) > len(suf):
            return s[:-len(suf)]
    return s


def parse_yushe_parts(fbfmc):
    """
    解析榆社县发包方名称，返回剥离行政后缀的 (乡镇名, 村名, 组名)。
    例如：'榆社县社城镇官上村圪麻窊组集体经济组织' -> ('社城', '官上', '圪麻窊')
          '榆社县社城镇官上村集体经济组织' -> ('社城', '官上', '')
    """
    zhen_raw, cun_raw, zu_raw = parse_yushe_fbfmc(fbfmc)
    zhen = zhen_raw
    for pfx in ('山西省晋中市榆社县', '晋中市榆社县', '榆社县'):
        if zhen.startswith(pfx):
            zhen = zhen[len(pfx):]
            break
    for sfx in ('街道', '乡', '镇'):
        if zhen.endswith(sfx) and len(zhen) > len(sfx):
            zhen = zhen[:-len(sfx)]
            break
    cun = cun_raw
    for sfx in ('社区', '村'):
        if cun.endswith(sfx) and len(cun) > len(sfx):
            cun = cun[:-len(sfx)]
            break
    zu = zu_raw
    for sfx in ('村民小组', '小组', '组', '队', '社'):
        if zu.endswith(sfx) and len(zu) > len(sfx):
            zu = zu[:-len(sfx)]
            break
    return zhen, cun, zu


def parse_gender_from_id(id_str, fallback='男'):
    """根据身份证号解析性别，18位第17位或15位第15位奇数男、偶数女，解析失败默认'男'"""
    if not id_str:
        return fallback
    clean_id = str(id_str).strip()
    if len(clean_id) == 18 and clean_id[:17].isdigit():
        return '男' if int(clean_id[16]) % 2 != 0 else '女'
    elif len(clean_id) == 15 and clean_id.isdigit():
        return '男' if int(clean_id[14]) % 2 != 0 else '女'
    return fallback


def _clean_filename(name):
    """清理文件名中的非法字符"""
    import re
    return re.sub(r'[\\/:*?"<>|\r\n\t]', '', str(name)).strip()


# 常用纸张尺寸（paperSize 编码 -> (宽mm, 高mm)），用于计算分页容量
_PAPER_SIZE_MM = {
    1: (215.9, 279.4),    # Letter
    5: (215.9, 355.6),    # Legal
    8: (297, 420),        # A3
    9: (210, 297),        # A4
    11: (210, 297),       # A4 Small
    12: (210, 297),       # A4 Transverse
}


def _insert_page_breaks(ws, hh_spans):
    """
    在每户数据块（hh_spans，Excel行号区间列表）边界插入水平分页符，
    保证每户合并单元格不会跨页被切断。
    依据页面纸张、页边距与行高计算每页可容纳的数据行数：
    第1页还需扣除表头（第1-8行）高度；第2页起表头不再重复。
    """
    from openpyxl.worksheet.pagebreak import Break

    paper_size = ws.page_setup.paperSize
    if paper_size not in _PAPER_SIZE_MM:
        logger.warning(f"榆社县公示表：未知纸张尺寸(paperSize={paper_size})，跳过插入分页符")
        return
    w_mm, h_mm = _PAPER_SIZE_MM[paper_size]
    if ws.page_setup.orientation == 'landscape':
        w_mm, h_mm = h_mm, w_mm

    margins = ws.page_margins
    # 页面高度换算为磅(1mm = 72/25.4 pt)
    page_pt = h_mm / 25.4 * 72
    body_pt = page_pt - (margins.top + margins.bottom) * 72

    default_h = ws.sheet_format.defaultRowHeight or 15
    row_h = {r: (ws.row_dimensions[r].height if r in ws.row_dimensions else default_h) for r in range(1, 9)}
    header_pt = sum(row_h.values())                # 首页表头 rows1-8
    repeat_pt = sum(row_h[r] for r in (6, 7, 8))   # 重复标题行 rows6-8

    data_row_pt = ws.row_dimensions[9].height if 9 in ws.row_dimensions else 36
    if not data_row_pt:
        data_row_pt = 36

    cap1 = int((body_pt - header_pt) // data_row_pt)
    cap2 = int((body_pt - repeat_pt) // data_row_pt)
    if cap1 < 1 or cap2 < 1:
        logger.warning(f"榆社县公示表：计算出的分页容量异常(cap1={cap1}, cap2={cap2})，跳过插入分页符")
        return

    breaks = []
    used = 0
    cap = cap1
    for sr, er in hh_spans:
        size = er - sr + 1
        if used + size <= cap:
            used += size
            continue
        # 当前页放不下此户，需要开启新页
        if not (used == 0 and size > cap2):
            # 例外：首页第一户即超过单页容量时无法避免跨页，只能放在首页
            breaks.append(sr)
        used = size
        cap = cap2
        if size > cap2:
            logger.warning(f"榆社县公示表：第{sr}-{er}行户块共{size}行，超过单页容量({cap2}行)，该户合并单元格将跨页")
    for brk in breaks:
        # OOXML <brk> 的 id 为 0 基行号（分页符显示在 1 基行 id+1 上方），
        # 传入的 brk 是 1 基行号，需减 1 使分页符位于户块开始行的上方
        ws.row_breaks.append(Break(id=brk - 1))
    logger.debug(f"榆社县公示表：按每户边界插入 {len(breaks)} 个分页符，第1页容量={cap1}行，后续页容量={cap2}行")


def generate_yushe_gsb_table(family_info_path, land_info_path, output_directory):
    """
    生成榆社县农村土地承包经营权调查公示表（张贴版+签字版）。
    按 FBFBM 分组，解析 FBFMC 得到乡镇名/村名（组名仅用于文件命名），
    每组分别套用 Templates/榆社-公示表（张贴版）.xlsx 与 榆社-公示表（签字版）.xlsx 模板。
    表头：D5=榆社县、H5=乡镇名、J5=村名、O5=FBFBM（按模板标签结构填写）。
    """
    def safe_read_excel(file_path, converters=None):
        try:
            return pd.read_excel(file_path, engine='openpyxl', converters=converters or {})
        except zipfile.BadZipFile:
            if os.path.exists(file_path):
                try:
                    logger.warning(f"openpyxl引擎无法读取文件 {file_path}，尝试xlrd引擎（xls格式）")
                    return pd.read_excel(file_path, engine='xlrd', converters=converters or {})
                except ImportError:
                    logger.warning("xlrd包未安装，只能读取xlsx格式文件")
                except Exception as e:
                    logger.error(f"xlrd引擎也无法读取 {file_path}: {str(e)}")
            logger.error(f"Excel文件读取失败，文件可能是损坏的或者格式不正确: {file_path}")
            raise ValueError(f"Excel文件读取失败!\n\n文件: {file_path}\n1. 文件损坏或格式错误\n2. 文件不是有效的Excel文件\n3. 安装了xlrd包可能可以读取xls格式文件")
        except Exception as e:
            logger.error(f"读取Excel文件时发生错误 {file_path}: {str(e)}")
            raise

    def conv_str_none(v):
        if v is None:
            return None
        if isinstance(v, float) and pd.isna(v):
            return None
        s = str(v).strip()
        return None if s.lower() in ('nan', '') else s

    def to_num(v):
        try:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return None
            return round(float(v), 2)
        except (TypeError, ValueError):
            return None

    def to_int(v):
        try:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return None
            return int(float(v))
        except (TypeError, ValueError):
            return None

    def _fmt_phone(v):
        """格式化电话：float整数转str、去首尾空白；空值返回None"""
        if v is None:
            return None
        if isinstance(v, float):
            if pd.isna(v):
                return None
            v = int(v) if v.is_integer() else round(v)
        s = str(v).strip()
        if s.lower() in ('nan', 'none', ''):
            return None
        return s

    family_conv = {'CBFBM': str,
                   'CYXM': conv_str_none, 'CYXB': conv_str_none, 'YHZGX': conv_str_none,
                   'CYZJHM': conv_str_none, 'CYBZ': conv_str_none, 'CYBZSM': conv_str_none}
    land_conv = {'CBFBM': str, 'DKBM': str,
                 'DKMC': conv_str_none, 'DKLB': conv_str_none, 'TDLYLX': conv_str_none,
                 'DLDJ': conv_str_none, 'SFJBNT': conv_str_none,
                 'DKDZ': conv_str_none, 'DKNZ': conv_str_none, 'DKXZ': conv_str_none,
                 'DKBZ': conv_str_none, 'DKBZXX': conv_str_none}

    family_info = safe_read_excel(family_info_path, converters=family_conv)
    land_info = safe_read_excel(land_info_path, converters=land_conv)

    if family_info.empty or 'FBFBM' not in family_info.columns:
        logger.error("榆社县公示表：家庭信息为空或缺少FBFBM列")
        raise ValueError("榆社县公示表：家庭信息为空或缺少FBFBM列")

    family_info['CBFBM'] = family_info['CBFBM'].astype(str)
    land_info['CBFBM'] = land_info['CBFBM'].astype(str)
    land_info['DKBM'] = land_info['DKBM'].astype(str)

    family_info = family_info.drop_duplicates(subset=['CBFBM', 'CYXM'])
    land_info = land_info.drop_duplicates(subset=['CBFBM', 'DKBM'])

    family_info['FBFBM'] = family_info['FBFBM'].astype(str).str.strip()
    family_info = family_info[family_info['FBFBM'].str.lower() != 'nan']
    family_info = family_info[family_info['FBFBM'] != '']

    template_dir = os.path.dirname(os.path.abspath(__file__))
    versions = ['张贴版', '签字版']
    os.makedirs(output_directory, exist_ok=True)

    for fbfbm, group_family in family_info.groupby('FBFBM', sort=False):
        group_land = land_info[land_info['CBFBM'].isin(group_family['CBFBM'])]

        hh = group_family.groupby('CBFBM', as_index=False).first()
        hh_map = {r['CBFBM']: r for _, r in hh.iterrows()}
        hh_codes = sorted(set(hh['CBFBM'].tolist()) | set(group_land['CBFBM'].tolist()))

        if not hh_codes:
            continue

        fbfmc = hh['FBFMC'].iloc[0] if not hh.empty and pd.notna(hh['FBFMC'].iloc[0]) else ''
        zhen, cun, zumin = parse_yushe_fbfmc(fbfmc)

        # 构建数据行（每户先成员后地块，行数取成员数与地块数较大值）
        data_rows = []
        hh_spans = []
        total_member = 0
        total_land = 0
        total_p = 0.0
        total_r = 0.0

        # 记录户主/本人成员行（写入承包方电话），用于签字版备注列
        hz_phone_rows = []

        for code in hh_codes:
            fam = group_family[group_family['CBFBM'] == code]
            lnd = group_land[group_land['CBFBM'] == code]
            hr = hh_map.get(code)

            cbfmc_val = (hr['CBFMC'] or '') if hr is not None else ''
            cbcfysl = to_int(hr['CBFCYSL']) if hr is not None else None
            cbdxzs = to_int(hr['CBDKZS']) if hr is not None else None
            member_count = cbcfysl if cbcfysl is not None and cbcfysl >= 1 else len(fam)
            land_count = cbdxzs if cbdxzs is not None and cbdxzs >= 1 else len(lnd)
            max_rows = max(member_count, land_count, 1)

            scmjm_sum = None
            if 'SCMJM' in lnd.columns:
                scm_series = pd.to_numeric(lnd['SCMJM'], errors='coerce')
                if scm_series.notna().any():
                    scmjm_sum = to_num(scm_series.sum())
            htm_series = pd.to_numeric(lnd['HTMJM'], errors='coerce')
            if scmjm_sum is None:
                if htm_series.notna().any():
                    scmjm_sum = to_num(htm_series.sum())
            htmjm_sum = to_num(htm_series.sum()) if htm_series.notna().any() else None

            block_start = len(data_rows)
            for i in range(max_rows):
                f = fam.iloc[i] if i < len(fam) else None
                l = lnd.iloc[i] if i < len(lnd) else None
                row = {}
                if i == 0:
                    row[1] = str(code)[-4:]        # A 编号
                    row[2] = cbfmc_val             # B 承包方（代表）
                    row[3] = member_count          # C:D 成员数
                    row[12] = land_count           # L 地块数
                    row[17] = scmjm_sum            # Q 实测总面积
                    row[19] = htmjm_sum            # S 合同总面积
                if f is not None:
                    row[5] = f['CYXM'] or ''       # E 姓名
                    row[6] = f['CYZJHM'] or ''     # F 成员证件号
                    row[7] = xb_code_to_name(f['CYXB'])      # G 性别
                    row[8] = _calc_age_from_id(f['CYZJHM'])  # H 年龄
                    row[9] = yhzgx_code_to_name(f['YHZGX'])  # I 与承包方代表关系
                    row[10] = f['CYBZSM'] or f['CYBZ'] or ''  # J:K 成员备注
                    if f.get('YHZGX') in ('01', '02'):
                        phone_col = 'LXDH_y' if 'LXDH_y' in f.index else ('LXDH' if 'LXDH' in f.index else None)
                        if phone_col:
                            hz_phone_rows.append((9 + block_start + i, f[phone_col]))
                if l is not None:
                    row[13] = l['DKMC'] or ''      # M:N 地块名称
                    row[15] = str(l['DKBM'])[-5:]  # O 地块编码
                    scmjm = to_num(l['SCMJM']) if 'SCMJM' in l.index else None
                    if scmjm is None:
                        scmjm = to_num(l['HTMJM'])
                    row[16] = scmjm                # P 实测面积
                    row[18] = to_num(l['HTMJM'])   # R 合同面积
                    row[20] = "种植业"             # T 土地用途
                    row[21] = dklb_code_to_name(l['DKLB'])          # U 地块类别
                    row[22] = tdlylx_code_to_name(l['TDLYLX'])      # V 土地利用类型
                    row[23] = dldj_code_to_name(l['DLDJ'])          # W 地力等级
                    row[24] = sfjbnt_code_to_name(l['SFJBNT'])      # X 是否基本农田
                    row[25] = l['DKDZ'] or ''      # Y 东至
                    row[26] = l['DKNZ'] or ''      # Z 南至
                    row[27] = l['DKXZ'] or ''      # AA 西至
                    row[28] = l['DKBZ'] or ''      # AB 北至
                    row[29] = l['DKBZXX'] or ''    # AC 备注
                data_rows.append(row)
            hh_spans.append((9 + block_start, 9 + len(data_rows) - 1))
            total_member += member_count
            total_land += land_count
            total_p += scmjm_sum or 0
            total_r += htmjm_sum or 0

        if not data_rows:
            continue

        for version in versions:
            template_path = os.path.join(template_dir, 'Templates', f'榆社-公示表（{version}）.xlsx')
            if not os.path.exists(template_path):
                logger.error(f"榆社县公示表模板不存在: {template_path}")
                raise FileNotFoundError(f"榆社县公示表模板不存在: {template_path}")

            filename_base = f"{fbfbm}榆社县{zhen}{cun}"
            if zumin:
                filename_base += f"({zumin})"
            filename = _clean_filename(f"{filename_base}公示表（{version}）.xlsx")
            output_file = os.path.join(output_directory, filename)
            shutil.copy2(template_path, output_file)
            wb = openpyxl.load_workbook(output_file)
            ws = wb.active
            ncols = ws.max_column

            # 保存模板样式与值
            data_style = {}
            for c in range(1, ncols + 1):
                cell = ws.cell(9, c)
                data_style[c] = {
                    'font': copy(cell.font), 'border': copy(cell.border),
                    'alignment': copy(cell.alignment), 'fill': copy(cell.fill),
                    'number_format': cell.number_format,
                }
            tail_styles = {}
            tail_values = {}
            for r in (19, 20, 21):
                for c in range(1, ncols + 1):
                    cell = ws.cell(r, c)
                    tail_styles[(r, c)] = {
                        'font': copy(cell.font), 'border': copy(cell.border),
                        'alignment': copy(cell.alignment), 'fill': copy(cell.fill),
                        'number_format': cell.number_format,
                    }
                    tail_values[(r, c)] = cell.value
            header_merges = [str(m) for m in ws.merged_cells.ranges if m.max_row <= 8]
            tail_merges = []
            for m in ws.merged_cells.ranges:
                if m.min_row >= 19:
                    tail_merges.append((m.min_row - 19, m.min_col, m.max_row - 19, m.max_col))
            row_heights = {r: ws.row_dimensions[r].height for r in (9, 19, 20, 21) if r in ws.row_dimensions}

            # 删除模板数据区域（第9行及以下），重建表头合并
            ws.delete_rows(9, ws.max_row - 8)
            ws.merged_cells = type(ws.merged_cells)()
            for hm in header_merges:
                ws.merge_cells(hm)

            # 填写表头（模板行5重设计后：D5=乡镇名；I5=村名；K5=组名，默认"/"；P5=FBFBM）
            ws.cell(5, 4).value = _strip_admin_suffix(zhen)   # G5 乡镇名(镇(乡))
            ws.cell(5, 9).value = _strip_admin_suffix(cun)    # I5 村名
            ws.cell(5, 11).value = _strip_admin_suffix(zumin) if zumin else '/'   # K5 组名，默认"/"
            ws.cell(5, 16).value = str(fbfbm)       # P5 FBFBM

            # 写入数据
            current = 9
            for row in data_rows:
                for c, v in row.items():
                    ws.cell(current, c, v)
                current += 1
            data_end = current - 1

            # 每户合并 A/B/C:D/K/P/R；签字版额外纵向合并「签字确认」列(AD/30)。C:D 单行户也横跨。
            hh_merges = [(1, 1), (2, 2), (3, 4), (12, 12), (17, 17), (19, 19)]
            if version == '签字版':
                hh_merges.append((30, 30))
            for sr, er in hh_spans:
                if er > sr:
                    for cols in hh_merges:
                        ws.merge_cells(start_row=sr, start_column=cols[0], end_row=er, end_column=cols[1])
                else:
                    ws.merge_cells(start_row=sr, start_column=3, end_row=sr, end_column=4)
            # 每行合并 I:J 与 L:M
            for r in range(9, data_end + 1):
                ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)
                ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=14)

            # 应用数据行样式
            for r in range(9, data_end + 1):
                for c in range(1, ncols + 1):
                    cell = ws.cell(r, c)
                    s = data_style[c]
                    cell.font = s['font']
                    cell.border = s['border']
                    cell.alignment = s['alignment']
                    if s['fill']:
                        cell.fill = s['fill']
                    cell.number_format = s['number_format']
            # 编号/地块编码文本格式，面积保留两位小数
            for r in range(9, data_end + 1):
                ws.cell(r, 1).number_format = '@'
                ws.cell(r, 6).number_format = '@'   # 成员证件号 文本
                ws.cell(r, 15).number_format = '@'  # 地块编码 文本
                for c in (16, 17, 18, 19):
                    ws.cell(r, c).number_format = '0.00'

            # 签字版：备注列写入户主/本人行对应的承包方电话，字体设为白色（隐藏显示，数值保留）
            if version == '签字版' and hz_phone_rows:
                for _r, _v in hz_phone_rows:
                    p = _fmt_phone(_v)
                    if not p:
                        continue
                    cell = ws.cell(_r, 29)
                    cell.value = p
                    cell.number_format = '@'
                    cell.font = Font(color='FFFFFFFF')   # 白字隐藏

            # 合计行（第19行样式+值，重写汇总数值）
            sr_tail = data_end + 1
            for c in range(1, ncols + 1):
                cell = ws.cell(sr_tail, c)
                s = tail_styles[(19, c)]
                cell.font = s['font']
                cell.border = s['border']
                cell.alignment = s['alignment']
                if s['fill']:
                    cell.fill = s['fill']
                cell.number_format = s['number_format']
                cell.value = tail_values[(19, c)]
            ws.cell(sr_tail, 1).value = '合计'
            ws.cell(sr_tail, 3).value = total_member
            ws.cell(sr_tail, 12).value = total_land
            ws.cell(sr_tail, 17).value = round(total_p, 2)
            ws.cell(sr_tail, 19).value = round(total_r, 2)

            # 页脚行（第20-21行）
            f1, f2 = sr_tail + 1, sr_tail + 2
            for rr, src in ((f1, 20), (f2, 21)):
                for c in range(1, ncols + 1):
                    cell = ws.cell(rr, c)
                    s = tail_styles[(src, c)]
                    cell.font = s['font']
                    cell.border = s['border']
                    cell.alignment = s['alignment']
                    if s['fill']:
                        cell.fill = s['fill']
                    cell.number_format = s['number_format']
                    cell.value = tail_values[(src, c)]

            # 重建合计/页脚合并
            for rel1, mc1, rel2, mc2 in tail_merges:
                ws.merge_cells(start_row=sr_tail + rel1, start_column=mc1, end_row=sr_tail + rel2, end_column=mc2)

            # 行高
            for r in range(9, data_end + 1):
                ws.row_dimensions[r].height = row_heights.get(9)
            ws.row_dimensions[sr_tail].height = row_heights.get(19)
            ws.row_dimensions[f1].height = row_heights.get(20)
            ws.row_dimensions[f2].height = row_heights.get(21)

            # 在每户边界插入分页符，保证合并单元格不分页
            _insert_page_breaks(ws, hh_spans)

            # 设置第6-8行为重复标题行（每页顶部打印列头）
            ws.print_title_rows = '6:8'

            # 输出前对每个单元格再次设置一遍边框，确保合并单元格各边边框完整打印
            tail_src = {sr_tail: 19, f1: 20, f2: 21}
            for r in range(1, ws.max_row + 1):
                if 9 <= r <= data_end:
                    for c in range(1, ncols + 1):
                        ws.cell(r, c).border = data_style[c]['border']
                elif r in tail_src:
                    for c in range(1, ncols + 1):
                        ws.cell(r, c).border = tail_styles[(tail_src[r], c)]['border']

            wb.save(output_file)
            logger.info(f"榆社县公示表已保存到: {output_file}")
            print(f"榆社县公示表已保存到: {output_file}")


def generate_yushe_gsb_only(mdb_file_path, shp_file_path, output_directory):
    """
    独立的榆社县公示表生成功能（张贴版+签字版）。
    直接处理MDB与SHP，写出临时数据后调用 generate_yushe_gsb_table。
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating Yushe GSB tables from: {mdb_file_path}, {shp_file_path}")

    temp_dir = tempfile.mkdtemp()
    excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
    excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=' + mdb_file_path + ';'
        )
        conn = pyodbc.connect(conn_str)

        try:
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}

            # 家庭信息：成员 + 承包方（成员数）+ 合同（地块数）+ 发包方（FBFMC）
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbht = dfs['CBHT']
            cbht_selected = cbht[['CBFBM', 'CBDKZS']]
            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC', 'CBFBM', 'CBFCYSL', 'LXDH']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf[['FBFBM', 'FBFMC']], left_on='CBFBM_14', right_on='FBFBM', how='left')

            # 地块信息：CBDKXX 关联 SHP 矢量属性
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)

            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')

            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)
            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)

            generate_yushe_gsb_table(excel_file1, excel_file2, output_directory)

        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")
    finally:
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            logger.warning(f"无法删除临时文件夹 {temp_dir}: {e}")

    logger.debug("Yushe GSB table generation completed")


def replace_bookmark_in_doc(doc, bookmark_name, replacement_text):
    """替换Word文档中指定书签的内容"""
    from docx.oxml.ns import qn
    from copy import deepcopy

    def process_paragraph(para):
        para_xml = para._p
        children = list(para_xml)
        bm_start_tag = qn("w:bookmarkStart")
        bm_end_tag = qn("w:bookmarkEnd")

        for index, child in enumerate(children):
            if child.tag != bm_start_tag:
                continue
            bm_name = child.get(qn("w:name"))
            if not bm_name or (bm_name != bookmark_name and bm_name.upper() != bookmark_name.upper()):
                continue

            bookmark_id = child.get(qn("w:id"))
            end_index = None
            for search_index in range(index + 1, len(children)):
                end_child = children[search_index]
                if end_child.tag == bm_end_tag and end_child.get(qn("w:id")) == bookmark_id:
                    end_index = search_index
                    break

            if end_index is None:
                # 若段落内未找到 bookmarkEnd（例如 bookmarkEnd 位于 body 级），默认替换至段落末尾
                end_index = len(children)

            template_run = None
            for candidate in children[index + 1:end_index]:
                if candidate.tag == qn("w:r"):
                    template_run = deepcopy(candidate)
                    break

            for removable in children[index + 1:end_index]:
                para_xml.remove(removable)

            if template_run is not None:
                run_element = template_run
                for text_node in list(run_element.findall(qn("w:t"))):
                    run_element.remove(text_node)
            else:
                run_element = OxmlElement("w:r")
                rPr = OxmlElement("w:rPr")
                sz = OxmlElement("w:sz")
                sz.set(qn("w:val"), "24")
                szCs = OxmlElement("w:szCs")
                szCs.set(qn("w:val"), "24")
                rPr.append(sz)
                rPr.append(szCs)
                run_element.insert(0, rPr)

            text_element = OxmlElement("w:t")
            text_element.set(qn("xml:space"), "preserve")
            text_element.text = str(replacement_text)
            run_element.append(text_element)
            para_xml.insert(index + 1, run_element)
            return True
        return False

    for para in doc.paragraphs:
        process_paragraph(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    process_paragraph(para)


def generate_fbfdc_table(mdb_file_path, output_directory):
    """
    生成发包方调查表
    从FBF表读取数据，按FBFBM前12位分组，每组取第一条记录
    写入Templates/发包方调查表.docx模板中对应的书签位置
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating FBFDC table from: {mdb_file_path}")

    if Document is None:
        logger.error("python-docx库未安装，无法生成发包方调查表")
        return

    # 连接MDB
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + mdb_file_path + ';'
    )
    conn = pyodbc.connect(conn_str)

    try:
        # 读取FBF表
        fbf = pd.read_sql_query("SELECT FBFBM, FBFMC, FBFFZRXM, FZRZJHM, LXDH, FBFDZ, YZBM, FBFDCY, FBFDCRQ FROM [FBF]", conn)

        if fbf.empty:
            logger.error("FBF表中没有数据")
            return

        # 按FBFBM前12位分组，每组取第一条
        fbf['group_key'] = fbf['FBFBM'].str[:12]
        first_per_group = fbf.groupby('group_key', as_index=False).first()

        # 模板路径
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '发包方调查表.docx')

        for _, row in first_per_group.iterrows():
            # 构建数据字典
            fbfdcy = str(row['FBFDCY']) if pd.notna(row['FBFDCY']) else ''
            fbfdcrq = row['FBFDCRQ']
            if pd.notna(fbfdcrq):
                try:
                    fbfdcrq = pd.to_datetime(fbfdcrq).strftime('%Y年%m月%d日')
                except Exception:
                    fbfdcrq = str(fbfdcrq)
            else:
                fbfdcrq = ''

            data = {
                'FBFBM': str(row['FBFBM']) if pd.notna(row['FBFBM']) else '',
                'FBFMC': str(row['FBFMC']) if pd.notna(row['FBFMC']) else '',
                'FBFFZRXM': str(row['FBFFZRXM']) if pd.notna(row['FBFFZRXM']) else '',
                'FZRZJHM': str(row['FZRZJHM']) if pd.notna(row['FZRZJHM']) else '',
                'LXDH': str(row['LXDH']) if pd.notna(row['LXDH']) else '',
                'FBFDZ': str(row['FBFDZ']) if pd.notna(row['FBFDZ']) else '',
                'YZBM': str(row['YZBM']) if pd.notna(row['YZBM']) else '',
                'FBFDCY': fbfdcy,
                'FBFDCRQ': fbfdcrq,
            }

            # FBFMC书签替换为：截取股份经济合作社之前 + 股份经济合作社
            fbfmc_parsed = data['FBFMC'].split('股份经济合作社')[0] + "股份经济合作社"
            data['FBFMC'] = fbfmc_parsed

            # 文件名和调查记事用的FBFMC（去掉明光市和股份经济合作社）
            fbfmc_clean = fbfmc_parsed.replace('明光市', '').replace('股份经济合作社', '')

            filename = f"{fbfmc_clean}发包方调查表.docx"
            output_file = os.path.join(output_directory, filename)

            fbffzrxm = data['FBFFZRXM']
            fbfdcjs = f"经实地调查，发包土地为{fbfmc_clean}所有，发包方责任人为{fbffzrxm}，以上内容确认属实，无异议。"
            data['FBFDCJS'] = fbfdcjs

            try:
                shutil.copy2(template_path, output_file)
                doc = Document(output_file)

                for key, value in data.items():
                    replace_bookmark_in_doc(doc, key, value)

                doc.save(output_file)
                logger.info(f"发包方调查表已保存到: {output_file}")
                print(f"发包方调查表已保存到: {output_file}")

            except Exception as e:
                logger.error(f"生成发包方调查表失败 {output_file}: {str(e)}")
                raise

    except Exception as e:
        logger.error(f"发生错误: {e}")
        logger.error(traceback.format_exc())
        raise
    finally:
        conn.close()
        logger.debug("Database connection closed")

    logger.debug("FBFDC table generation completed")


def generate_fbfdc_table_only(mdb_file_path, output_directory):
    """
    独立的发包方调查表生成功能
    直接处理MDB文件并生成发包方调查表
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating FBFDC table from: {mdb_file_path}")
    generate_fbfdc_table(mdb_file_path, output_directory)
    logger.debug("FBFDC table generation completed")


def generate_yhqd_table_only(mdb_file_path, output_directory):
    """
    生成一户一档清单
    从CBF表读取数据，按CBFBM前14位分组，每组输出一个xlsx文件
    使用Templates/03户档清单.xlsx模板
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating YHQD table from: {mdb_file_path}")

    # 读取权属单位代码表（用于名称查询）
    resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
    code_table_path = os.path.join(resources_dir, '3411822026权属单位代码表.xls')
    qsdm_df = pd.read_excel(code_table_path)
    qsdm_df.columns = ['code', 'name']
    qsdm_df['code'] = qsdm_df['code'].astype(str)
    code_to_name = dict(zip(qsdm_df['code'], qsdm_df['name']))
    all_codes = qsdm_df['code'].tolist()

    # 连接MDB
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + mdb_file_path + ';'
    )
    conn = pyodbc.connect(conn_str)

    try:
        # 读取CBF表
        cbf = pd.read_sql_query("SELECT CBFBM, CBFMC, CBFZJHM, CBFCYSL FROM [CBF]", conn)
        if cbf.empty:
            logger.error("CBF表中没有数据")
            return

        cbf['CBFBM'] = cbf['CBFBM'].astype(str)
        cbf['group_code'] = cbf['CBFBM'].str[:14]

        # 按group_code分组（自动处理所有组）
        grouped = cbf.groupby('group_code')

        # 模板路径
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '03户档清单.xlsx')
        total_groups = 0
        skipped_groups = 0

        # 遍历每个小组代码
        for group_code, group_df in grouped:
            group_name = code_to_name.get(group_code, group_code)
            logger.info(f"正在生成: {group_name} ({group_code}), 共{len(group_df)}户")

            # 解析名称各层级用于标题和文件名
            # 从名称字符串解析镇/村/组层级（代码表中14位编码无前缀层级关系）
            name = group_name
            if name.startswith('明光市'):
                name = name[3:]
            zhen_name = ''
            zhen_suffix = ''
            cun_name = ''
            cun_suffix = ''
            zu_name = ''
            parse_name = name.replace('股份经济合作社', '').strip()
            # 找镇级边界
            rest = parse_name
            for s in ['街道', '镇', '乡']:
                idx = parse_name.find(s)
                if idx != -1:
                    zhen_name = parse_name[:idx]
                    zhen_suffix = s
                    rest = parse_name[idx + len(s):]
                    break
            # 找村级边界
            for s in ['村村', '社区', '村']:
                idx = rest.find(s)
                if idx != -1:
                    cun_name = rest[:idx]
                    zu_name = rest[idx + len(s):]
                    cun_suffix = s
                    break
            else:
                cun_name = rest
                zu_name = ''

            title_parts = []
            if zhen_name:
                title_parts.append(f"{zhen_name}{zhen_suffix}")
            if cun_name:
                title_parts.append(f"{cun_name}{cun_suffix}")
            if zu_name:
                if not zu_name.endswith('组'):
                    title_parts.append(f"{zu_name}组")
                else:
                    title_parts.append(zu_name)
            title_parts.append("一户一档清单")
            title = ''.join(title_parts)

            # 输出文件名：权属单位名称（去掉"股份经济合作社"）+ 14位代码 + 一户一档清单
            name_part = group_name.replace('股份经济合作社', '')
            filename = f"{name_part}({group_code})一户一档清单.xlsx"
            output_file = os.path.join(output_directory, filename)

            try:
                shutil.copy2(template_path, output_file)
                wb = openpyxl.load_workbook(output_file)
                ws = wb['Sheet1']

                ws['A1'] = title

                group_count = len(group_df)

                # === 预读模板 footer 行（18-19）样式/值/合并/行高 ===
                footer_data = {}
                for r in (18, 19):
                    cells = {}
                    for c in range(1, 26):
                        src = ws.cell(row=r, column=c)
                        cells[c] = {
                            'value': src.value,
                            'font': copy(src.font),
                            'alignment': copy(src.alignment),
                            'border': copy(src.border),
                            'fill': copy(src.fill),
                        }
                    footer_data[r] = {
                        'cells': cells,
                        'height': ws.row_dimensions[r].height,
                    }
                footer_merges = []
                for mr in list(ws.merged_cells.ranges):
                    if 18 <= mr.min_row <= 19 or 18 <= mr.max_row <= 19:
                        footer_merges.append((mr.min_col, mr.min_row, mr.max_col, mr.max_row))

                # 预读模板第一行数据样式（原 row4）
                template_style_refs = {}
                for c in range(1, 26):
                    src = ws.cell(row=4, column=c)
                    template_style_refs[c] = {
                        'font': copy(src.font),
                        'alignment': copy(src.alignment),
                        'border': copy(src.border),
                    }

                # 删除模板 data 区+footer（row4-19 共16行）
                ws.delete_rows(4, 16)

                # 清理 delete_rows 未能移除的 ghost merge ranges (openpyxl bug)
                ws.merged_cells.ranges -= {mr for mr in ws.merged_cells.ranges if 18 <= mr.min_row <= 19}

                # 清理 delete_rows 残留的 row_dimensions（模板 row 18-19 行高不随 delete_rows 移除）
                for r in range(4, 20):
                    try:
                        del ws.row_dimensions[r]
                    except (KeyError, TypeError):
                        pass

                # 写入数据行
                group_df_sorted = group_df.sort_values('CBFBM')
                for idx, (_, row) in enumerate(group_df_sorted.iterrows()):
                    row_num = 4 + idx
                    for c in range(1, 26):
                        dst = ws.cell(row=row_num, column=c)
                        dst.font = template_style_refs[c]['font']
                        dst.alignment = template_style_refs[c]['alignment']
                        dst.border = template_style_refs[c]['border']
                        dst.fill = PatternFill(fill_type=None)
                    ws.cell(row=row_num, column=1).value = idx + 1
                    ws.cell(row=row_num, column=2).value = str(row['CBFMC']) if pd.notna(row['CBFMC']) else ''
                    ws.cell(row=row_num, column=3).value = str(row['CBFZJHM']) if pd.notna(row['CBFZJHM']) else ''
                    ws.cell(row=row_num, column=4).value = '✓'
                    cy_raw = pd.to_numeric(row['CBFCYSL'], errors='coerce')
                    cy_count = int(cy_raw) if pd.notna(cy_raw) else 0
                    ws.cell(row=row_num, column=5).value = cy_count

                # 在数据行末尾追加 footer
                footer_start_row = 4 + group_count
                for template_row, actual_row in ((18, footer_start_row), (19, footer_start_row + 1)):
                    ws.row_dimensions[actual_row].height = footer_data[template_row]['height']
                    for c in range(1, 26):
                        dst = ws.cell(row=actual_row, column=c)
                        dst.value = footer_data[template_row]['cells'][c]['value']
                        dst.font = footer_data[template_row]['cells'][c]['font']
                        dst.alignment = footer_data[template_row]['cells'][c]['alignment']
                        dst.border = footer_data[template_row]['cells'][c]['border']
                        dst.fill = footer_data[template_row]['cells'][c]['fill']

                # 重新应用 footer 合并单元格
                for min_col, min_row, max_col, max_row in footer_merges:
                    new_min_row = min_row - 18 + footer_start_row
                    new_max_row = max_row - 18 + footer_start_row
                    ws.merge_cells(start_row=new_min_row, start_column=min_col, end_row=new_max_row, end_column=max_col)

                # 尝试更新表头（CBFCYSL > 10 时修改N~W列）
                try:
                    cbfcysl_numeric = pd.to_numeric(group_df['CBFCYSL'], errors='coerce')
                    has_cy_large = (cbfcysl_numeric.dropna() > 10).any()
                    if has_cy_large:
                        page_header_map = {
                            14: '家庭成员页1-2', 15: '家庭成员页3-4', 16: '家庭成员页5-6',
                            17: '家庭成员页7-8', 18: '家庭成员页9-10', 19: '家庭成员页11-12',
                            20: '家庭成员页13-14', 21: '家庭成员页15-16', 22: '家庭成员页17-18',
                            23: '家庭成员页19-20', 24: '家庭成员页21',
                        }
                        for col_idx, header_text in page_header_map.items():
                            ws.cell(row=3, column=col_idx).value = header_text
                except Exception as e:
                    logger.warning(f"更新表头失败（使用默认表头）: {e}")

                wb.save(output_file)
                logger.info(f"一户一档清单已保存到: {output_file}")
                total_groups += 1

            except Exception as e:
                logger.error(f"生成一户一档清单失败 {output_file}: {str(e)}")
                skipped_groups += 1
                continue

        logger.info(f"一户一档清单生成完成，共生成{total_groups}个文件，跳过{skipped_groups}个")

    except Exception as e:
        logger.error(f"发生错误: {e}")
        logger.error(traceback.format_exc())
        raise
    finally:
        conn.close()
        logger.debug("Database connection closed")

    logger.debug("YHQD table generation completed")


def generate_yjjsqd_table_only(mdb_file_path, output_directory):
    """
    生成移交技术资料清单
    从CBF表读取数据，按CBFBM前12位分组（村级），每组输出一个docx文件
    使用Templates/02第三方技术单位向乡镇移交技术资料清单.docx模板
    替换书签DJZQMC为"XX镇（街道、乡）XX村（社区、村村）"
    """
    os.makedirs(output_directory, exist_ok=True)
    logger.debug(f"Generating YJJSQD table from: {mdb_file_path}")

    # 读取权属单位代码表（用于名称查询）
    resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
    code_table_path = os.path.join(resources_dir, '3411822026权属单位代码表.xls')
    qsdm_df = pd.read_excel(code_table_path)
    qsdm_df.columns = ['code', 'name']
    qsdm_df['code'] = qsdm_df['code'].astype(str)
    code_to_name = dict(zip(qsdm_df['code'], qsdm_df['name']))

    # 连接MDB
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + mdb_file_path + ';'
    )
    conn = pyodbc.connect(conn_str)

    try:
        # 读取CBF表
        cbf = pd.read_sql_query("SELECT CBFBM FROM [CBF]", conn)
        if cbf.empty:
            logger.error("CBF表中没有数据")
            return

        cbf['CBFBM'] = cbf['CBFBM'].astype(str)
        cbf['group_code'] = cbf['CBFBM'].str[:12]

        # 按group_code分组（村级，12位编码）
        grouped = cbf.groupby('group_code')

        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '02第三方技术单位向乡镇移交技术资料清单.docx')
        total_groups = 0
        skipped_groups = 0

        for group_code, group_df in grouped:
            # 代码表中村级编码 = group_code + "00"（14位）
            full_code = group_code + "00"
            group_name = code_to_name.get(full_code, None)
            if group_name is None:
                logger.warning(f"未找到编码 {full_code} 对应的权属单位名称，跳过")
                skipped_groups += 1
                continue

            logger.info(f"正在生成: {group_name} ({full_code}), 共{len(group_df)}户")

            # 解析名称：明光市XX镇(乡/街道)XX村(社区/村村)股份经济合作社
            name = group_name
            if name.startswith('明光市'):
                name = name[3:]
            parse_name = name.replace('股份经济合作社', '').strip()

            zhen_name = ''
            zhen_suffix = ''
            cun_name = ''
            cun_suffix = ''
            rest = parse_name
            # 乡镇级关键字：乡、镇、街道
            for s in ['乡', '镇', '街道']:
                idx = rest.find(s)
                if idx != -1:
                    zhen_name = rest[:idx]
                    zhen_suffix = s
                    rest = rest[idx + len(s):]
                    break
            # 村级关键字：村村、社区、村（"村村"必须在"村"前）
            for s in ['村村', '社区', '村']:
                idx = rest.find(s)
                if idx != -1:
                    cun_name = rest[:idx]
                    cun_suffix = s
                    break

            djzqmc = f"{zhen_name}{zhen_suffix}{cun_name}{cun_suffix}"

            output_filename = f"第三方技术单位向乡镇移交技术资料清单({djzqmc}{group_code}).docx"
            output_file = os.path.join(output_directory, output_filename)

            try:
                # 复制模板
                shutil.copy2(template_path, output_file)

                # 打开docx并替换书签
                doc = Document(output_file)
                replace_bookmark_in_doc(doc, "DJZQMC", djzqmc)
                doc.save(output_file)

                logger.info(f"移交技术资料清单已保存到: {output_file}")
                total_groups += 1

            except Exception as e:
                logger.error(f"生成移交技术资料清单失败 {output_file}: {str(e)}")
                skipped_groups += 1
                continue

        logger.info(f"移交技术资料清单生成完成，共生成{total_groups}个文件，跳过{skipped_groups}个")

    except Exception as e:
        logger.error(f"发生错误: {e}")
        logger.error(traceback.format_exc())
        raise
    finally:
        conn.close()
        logger.debug("Database connection closed")

    logger.debug("YJJSQD document generation completed")


def _parse_cun_name(group_name):
    """从权属单位名称解析 镇名+村名，返回 (zhen_str, cun_str)"""
    name = group_name
    if name.startswith('明光市'):
        name = name[3:]
    parse_name = name.replace('股份经济合作社', '').strip()
    rest = parse_name
    zhen_name = ''
    zhen_suffix = ''
    cun_name = ''
    cun_suffix = ''
    for s in ['乡', '镇', '街道']:
        idx = rest.find(s)
        if idx != -1:
            zhen_name = rest[:idx]
            zhen_suffix = s
            rest = rest[idx + len(s):]
            break
    for s in ['村村', '社区', '村']:
        idx = rest.find(s)
        if idx != -1:
            cun_name = rest[:idx]
            cun_suffix = s
            break
    return f"{zhen_name}{zhen_suffix}", f"{cun_name}{cun_suffix}"


def _collect_tsqt_data(survey_dir):
    """收集摸底表数据，返回 village_data 列表和镇标题"""
    import pandas as pd
    resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
    code_table_path = os.path.join(resources_dir, '3411822026权属单位代码表.xls')
    qsdm_df = pd.read_excel(code_table_path)
    qsdm_df.columns = ['code', 'name']
    qsdm_df['code'] = qsdm_df['code'].astype(str)
    code_to_name = dict(zip(qsdm_df['code'], qsdm_df['name']))

    special_keyword_map = {
        "嫁": "出嫁", "入赘": "入赘婿", "女婿": "入赘婿",
        "新生": "新生婴儿", "嫁入": "婚入", "收养": "合法收养子女",
        "学生": "在校学生", "服刑": "服刑或刑满释放人员",
        "公务员": "公务员（正式事业编）", "事业编": "公务员（正式事业编）",
        "五保户": "五保户",
    }
    full_types = {
        "出嫁", "离婚", "丧偶", "入赘婿", "新生婴儿", "婚入",
        "合法收养子女", "全家进城落户", "现役军人", "在校学生",
        "服刑或刑满释放人员", "公务员（正式事业编）", "五保户",
    }

    def match_special_type(bgqk):
        if pd.isna(bgqk):
            return None
        bgqk = str(bgqk).strip()
        if not bgqk:
            return None
        if bgqk in full_types:
            return bgqk
        for keyword, group_type in special_keyword_map.items():
            if keyword in bgqk:
                return group_type
        return None

    village_data = []
    xlsx_files = [f for f in os.listdir(survey_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
    if not xlsx_files:
        logger.error(f"摸底表目录中没有找到xlsx文件: {survey_dir}")
        return [], ''

    for fname in sorted(xlsx_files):
        fpath = os.path.join(survey_dir, fname)
        try:
            df = pd.read_excel(fpath)
        except Exception as e:
            logger.warning(f"读取文件失败 {fname}: {e}")
            continue
        col_map = {}
        expected_cols = {'承包方编码': 0, '姓名': 1, '与户主关系': 2, '身份证号码': 3,
                        '联系电话': 4, '坐落': 5, '承包方名称': 6, '变更情况': 7}
        for col in list(df.columns):
            col_clean = col.strip()
            if col_clean in expected_cols:
                col_map[col_clean] = col

        if '变更情况' not in col_map:
            logger.warning(f"文件 {fname} 缺少'变更情况'列，跳过")
            continue

        members = []
        for _, row in df.iterrows():
            bgqk = row.get(col_map.get('变更情况'), None)
            group_type = match_special_type(bgqk)
            if group_type is None:
                continue
            cbfbm = str(row.get(col_map.get('承包方编码', ''), ''))
            code_12 = cbfbm[:12] if len(cbfbm) >= 12 else ''
            code_14 = cbfbm[:14] if len(cbfbm) >= 14 else ''
            cun_full_name = code_to_name.get(code_12 + '00', '')
            zu_full_name = code_to_name.get(code_14, '')
            zhen_str, _ = _parse_cun_name(cun_full_name)
            _, cun_str = _parse_cun_name(cun_full_name)
            zu_parsed = zu_full_name.replace('股份经济合作社', '').replace('明光市', '').strip() if zu_full_name else ''
            bare_cun = cun_str
            zu_prefix = f"{zhen_str}{cun_str}"
            zu_only = zu_parsed[len(zu_prefix):] if zu_parsed.startswith(zu_prefix) else zu_parsed
            members.append({
                'name': str(row.get(col_map.get('姓名', ''), '')),
                'id_number': str(row.get(col_map.get('身份证号码', ''), '')),
                'cun_name': bare_cun,
                'zu_name': zu_only,
                'cbfmc': str(row.get(col_map.get('承包方名称', ''), '')),
                'group_type': group_type,
                'code_12': code_12,
            })

        if members:
            cbfbm_first = str(df.iloc[0].get(col_map.get('承包方编码', ''), ''))
            code_12_first = cbfbm_first[:12] if len(cbfbm_first) >= 12 else ''
            cun_full = code_to_name.get(code_12_first + '00', '')
            zhen_str, cun_str = _parse_cun_name(cun_full)
            village_data.append({
                'cun_title': f"{zhen_str}{cun_str}",
                'zheh_str': zhen_str,
                'members': members,
            })
            logger.info(f"  文件 {fname}: 找到 {len(members)} 个特殊群体成员")

    zhen_title = village_data[0]['zheh_str'] if village_data else ''
    return village_data, zhen_title


def generate_tsqt_statistics_xlsx(survey_dir, output_path):
    """
    生成明光市特殊群体统计附表 (xlsx)
    按分组统计各特殊群体类型数量
    """
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side
    from copy import deepcopy

    village_data, zhen_title = _collect_tsqt_data(survey_dir)
    if not village_data:
        return

    thin_side = Side(style='thin', color='000000')
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    cell_align = Alignment(horizontal='center', vertical='center', wrapText=True)

    type_order = [
        "出嫁", "离婚", "丧偶", "入赘婿", "新生婴儿", "婚入",
        "合法收养子女", "全家进城落户", "现役军人", "在校学生",
        "服刑或刑满释放人员", "公务员（正式事业编）", "五保户",
    ]

    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 'Templates', 'XXXX镇特殊群体统计附表.xlsx')
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    ws.delete_rows(2, ws.max_row - 1)
    ws.row_dimensions[1].height = 18

    current_row = 2
    for vi, village in enumerate(village_data):
        type_counts = {}
        for m in village['members']:
            t = m['group_type']
            type_counts[t] = type_counts.get(t, 0) + 1
        sorted_types = sorted(type_counts.items(),
                              key=lambda x: type_order.index(x[0]) if x[0] in type_order else 999)
        start_row = current_row
        for ti, (type_name, count) in enumerate(sorted_types):
            ws.cell(current_row, 3, type_name)
            ws.cell(current_row, 4, count)
            for c in range(1, 5):
                cell = ws.cell(current_row, c)
                cell.border = cell_border
                cell.alignment = cell_align
            current_row += 1
        end_row = current_row - 1
        ws.cell(start_row, 1, vi + 1)
        ws.cell(start_row, 2, village['cun_title'])
        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
            ws.cell(start_row, 1).alignment = cell_align
            ws.cell(start_row, 2).alignment = cell_align
        for r in range(start_row, end_row + 1):
            ws.row_dimensions[r].height = 18

    output_path_file = os.path.join(output_path, f"02 {zhen_title}特殊群体统计附表.xlsx")
    wb.save(output_path_file)
    logger.info(f"特殊群体统计附表已保存到: {output_path_file}")
    logger.info(f"共处理 {len(village_data)} 个村")


def generate_tsqt_statistics(survey_dir, output_path, cover_date="    年  月  日"):
    """
    生成明光市二轮延包特殊群体统计表 (docx)
    survey_dir: 家庭成员摸底信息表目录（每个xlsx对应一个村）
    output_path: 输出目录（存放docx文件）
    cover_date: 封面日期，默认"    年  月  日"
    """
    from copy import deepcopy
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    village_data, zhen_title = _collect_tsqt_data(survey_dir)
    if not village_data:
        return

    # === 2. 构建 docx ===
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 'Templates', '09 XXXX镇特殊群体土地承包统计表.docx')

    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # 列宽定义：1.13, 1.45, 3.83, 1.60, 3.83, 1.62, 1.89, 1.39, 1.39, 1.67, 1.57, 1.41, 1.01, 0.93 cm
    column_widths_twips = [641, 822, 2172, 907, 2172, 919, 1072, 788, 788, 947, 890, 799, 573, 527]

    def set_table_title(table, cun_title):
        cell = table.rows[0].cells[0]
        cell.text = f"{cun_title}特殊群体土地延包统计表"
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.size = Pt(18)

    def format_table_layout(table):
        grid = table._tbl.find(f'{{{ns}}}tblGrid')
        grid_cols = grid.findall(f'{{{ns}}}gridCol')
        for i, col in enumerate(grid_cols):
            if i < len(column_widths_twips):
                col.set(qn('w:w'), str(column_widths_twips[i]))
        for ri, row in enumerate(table.rows):
            tr = row._tr
            tr_pr = tr.find(f'{{{ns}}}trPr')
            if tr_pr is None:
                tr_pr = OxmlElement('w:trPr')
                tr.insert(0, tr_pr)
            for existing_h in tr_pr.findall(f'{{{ns}}}trHeight'):
                tr_pr.remove(existing_h)
            tr_h = OxmlElement('w:trHeight')
            tr_h.set(qn('w:val'), '448')
            tr_h.set(qn('w:hRule'), 'atLeast')
            tr_pr.append(tr_h)
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if ri == 0:
                continue
            for ci, cell in enumerate(row.cells):
                if ci < len(column_widths_twips):
                    tc = cell._tc
                    tc_pr = tc.find(f'{{{ns}}}tcPr')
                    if tc_pr is None:
                        tc_pr = OxmlElement('w:tcPr')
                        tc.insert(0, tc_pr)
                    for existing_w in tc_pr.findall(f'{{{ns}}}tcW'):
                        tc_pr.remove(existing_w)
                    tcW = OxmlElement('w:tcW')
                    tcW.set(qn('w:w'), str(column_widths_twips[ci]))
                    tcW.set(qn('w:type'), 'dxa')
                    tc_pr.append(tcW)

    def _extract_id_birth_year(id_str):
        id_str = id_str.strip()
        if len(id_str) == 18:
            return id_str[6:10]
        elif len(id_str) == 15:
            return '19' + id_str[6:8]
        return ''

    def fill_table_rows(table, members, start_row_index=2):
        tr_elements = table._tbl.findall(f'{{{ns}}}tr')
        current_data_rows = len(tr_elements) - 2
        target_rows = len(members)
        if target_rows > current_data_rows:
            rows_to_add = target_rows - current_data_rows
            last_tr_proto = deepcopy(tr_elements[-1])
            for _ in range(rows_to_add):
                table._tbl.append(deepcopy(last_tr_proto))
        elif target_rows < current_data_rows:
            rows_to_remove = current_data_rows - target_rows
            for _ in range(rows_to_remove):
                all_tr = table._tbl.findall(f'{{{ns}}}tr')
                if len(all_tr) > 2:
                    table._tbl.remove(all_tr[-1])
        new_trs = table._tbl.findall(f'{{{ns}}}tr')
        for mi, member in enumerate(members):
            row_idx = start_row_index + mi
            if row_idx >= len(new_trs):
                break
            cells = table.rows[row_idx].cells
            cells[0].text = str(mi + 1)
            cells[1].text = member['name']
            cells[2].text = member['id_number']
            cells[3].text = member['cun_name']
            cells[4].text = member['zu_name']
            cells[5].text = member['cbfmc']
            cells[6].text = member['group_type']
            birth_year = _extract_id_birth_year(member['id_number'])
            cells[7].text = birth_year if member['group_type'] == '新生婴儿' else ''
            no_right = member['group_type'] == '公务员（正式事业编）'
            cells[8].text = '否' if no_right else '是'
            cells[9].text = '否'
            cells[10].text = '否'
            cells[11].text = '否'
            cells[12].text = '否' if no_right else '是'
            for c in cells:
                for para in c.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)

    def replace_rq_bookmark(doc, cover_date):
        for para in doc.paragraphs:
            bm_starts = para._p.findall(f'{{{ns}}}bookmarkStart')
            found_rq = False
            for bm in bm_starts:
                if bm.get(f'{{{ns}}}name') == 'RQ':
                    found_rq = True
                    break
            if not found_rq:
                continue
            children = list(para._p)
            bm_idx = children.index(bm)
            to_remove = [c for c in children[bm_idx + 1:] if c.tag != f'{{{ns}}}bookmarkEnd']
            for c in to_remove:
                para._p.remove(c)
            run_elem = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '36')
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), '36')
            rPr.append(sz)
            rPr.append(szCs)
            run_elem.append(rPr)
            t_elem = OxmlElement('w:t')
            t_elem.set(qn('xml:space'), 'preserve')
            t_elem.text = cover_date
            run_elem.append(t_elem)
            para._p.append(run_elem)
            break

    def _find_fill_shuoming(body):
        """查找填写说明段落，返回其 XML 元素或 None"""
        for child in body:
            if child.tag == qn('w:p'):
                texts = [n.text or '' for n in child.findall(f'.//{{{ns}}}t')]
                if ''.join(texts).strip().startswith('填写说明'):
                    return child
        return None

    # 封面文件（一个镇一份，不含填写说明）
    zhen_title = village_data[0]['zheh_str']
    cover_doc = Document(template_path)
    cover_body = cover_doc.element.body
    for child in list(cover_body):
        if child.tag == qn('w:tbl'):
            cover_body.remove(child)
    shuoming = _find_fill_shuoming(cover_body)
    if shuoming is not None:
        cover_body.remove(shuoming)
    for para in list(cover_body.iterchildren(f'{{{ns}}}p')):
        brs = para.findall(f'.//{{{ns}}}br')
        if any(br.get(f'{{{ns}}}type') == 'page' for br in brs):
            cover_body.remove(para)
    sect_prs = list(cover_body.iterchildren(f'{{{ns}}}sectPr'))
    for sp in sect_prs[:-1]:
        cover_body.remove(sp)
    cover_paras = list(cover_body.iterchildren(f'{{{ns}}}p'))
    for para in reversed(cover_paras):
        texts = [t.text or '' for t in para.findall(f'.//{{{ns}}}t')]
        if not any(text.strip() for text in texts):
            cover_body.remove(para)
        else:
            break
    replace_bookmark_in_doc(cover_doc, "XJQYMC1", zhen_title)
    replace_bookmark_in_doc(cover_doc, "XJQYMC2", zhen_title)
    replace_rq_bookmark(cover_doc, cover_date)
    cover_path = os.path.join(output_path, f"01 {zhen_title}特殊群体统计封面.docx")
    cover_doc.save(cover_path)
    logger.info(f"  封面已保存: {cover_path}")

    # 每个分组独立生成一个文件（不含封面，表格后附加填写说明）
    for vi, village in enumerate(village_data):
        doc = Document(template_path)
        body = doc.element.body
        fill_shuoming = _find_fill_shuoming(body)
        fill_shuoming_copy = deepcopy(fill_shuoming) if fill_shuoming is not None else None
        for child in list(body):
            if child.tag == qn('w:p'):
                body.remove(child)
        table = doc.tables[0]
        set_table_title(table, village['cun_title'])
        fill_table_rows(table, village['members'])
        format_table_layout(table)
        if fill_shuoming_copy is not None:
            body.append(fill_shuoming_copy)

        safe_name = village['cun_title'].replace('/', '_').replace('\\', '_').replace(':', '_')
        file_path = os.path.join(output_path, f"{vi + 3:02d} {safe_name}特殊群体土地延包统计表.docx")
        doc.save(file_path)
        logger.info(f"  已保存: {file_path}")

    logger.info(f"共处理 {len(village_data)} 个村（文件），{sum(len(v['members']) for v in village_data)} 个特殊群体成员")


def insert_rows_with_merges(ws, insert_pos, amount):
    """
    在指定位置插入新行，同时正确平移下方所有合并单元格（merged_cells）与行高设置（row_dimensions），
    避免 openpyxl 默认 insert_rows 不平移合并区域导致的样式冲突与单元格错位。
    """
    from openpyxl.worksheet.cell_range import CellRange

    # 1. 平移所有受影响的合并单元格
    new_ranges = []
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= insert_pos:
            ws.merged_cells.remove(rng)
            new_ranges.append(CellRange(
                min_col=rng.min_col,
                min_row=rng.min_row + amount,
                max_col=rng.max_col,
                max_row=rng.max_row + amount
            ))
        elif rng.max_row >= insert_pos:
            ws.merged_cells.remove(rng)
            new_ranges.append(CellRange(
                min_col=rng.min_col,
                min_row=rng.min_row,
                max_col=rng.max_col,
                max_row=rng.max_row + amount
            ))

    # 2. 平移所有受影响的行高
    new_row_dimensions = {}
    for r, dim in list(ws.row_dimensions.items()):
        if r >= insert_pos:
            del ws.row_dimensions[r]
            new_row_dimensions[r + amount] = dim
    for r, dim in new_row_dimensions.items():
        ws.row_dimensions[r] = dim

    # 3. 插入行（平移单元格数据）
    ws.insert_rows(insert_pos, amount)

    # 4. 重新挂载平移后的合并单元格
    for rng in new_ranges:
        ws.merged_cells.add(rng)


def chain_parcel_boundary(sub_lines, dkbm):
    """
    通过拓扑首尾相连算法对地块界址线进行闭环排序，
    并准确解析每个顶点的真实界址点号，彻底解决 Shapefile 矢量方向不一致、
    反向界址线导致点号相邻重复、以及宗地共享界址线乱序的问题。
    """
    import re
    lines = sub_lines.to_dict('records')
    if len(lines) < 3:
        return lines, [str(l.get('QJZDH', '')).strip() for l in lines]

    # 若存在宗地内细分线等非主边界线，优先提取包含 'T数字' 标号的主边界界址线
    t_lines = [l for l in lines if re.search(r'[由至到]T\d+|T\d+沿', str(l.get('JZXSM', '')))]
    if len(t_lines) >= 3:
        lines = t_lines
    n = len(lines)

    # 1. 寻找起始线（优先取 DKBM 完全匹配且包含 '由T1' 或 'T1沿' 的界址线）
    start_candidates = []
    for idx, l in enumerate(lines):
        sm = str(l.get('JZXSM', ''))
        if ('由T1' in sm or 'T1沿' in sm) and ('至T2' in sm or '到T2' in sm):
            start_candidates.append((l.get('DKBM') == dkbm, idx))

    if start_candidates:
        start_candidates.sort(key=lambda x: x[0], reverse=True)
        curr_line_idx = start_candidates[0][1]
    else:
        t1_cands = [i for i, l in enumerate(lines) if '由T1' in str(l.get('JZXSM', '')) or 'T1沿' in str(l.get('JZXSM', ''))]
        curr_line_idx = t1_cands[0] if t1_cands else 0

    curr_line = lines[curr_line_idx]

    # 2. 确定该起始线在闭环中哪一端是 T1
    t1_pt = None
    for l in lines:
        sm = str(l.get('JZXSM', ''))
        if ('至T1' in sm or '到T1' in sm) and (l.get('DKBM') == dkbm or len(lines) <= 4):
            inter = {str(curr_line['QJZDH']), str(curr_line['ZJZDH'])}.intersection({str(l['QJZDH']), str(l['ZJZDH'])})
            if len(inter) == 1:
                t1_pt = list(inter)[0]
                break

    if not t1_pt:
        t1_pt = str(curr_line['QJZDH']).strip()

    ordered_lines = [curr_line]
    ordered_pts = [t1_pt]
    used_line_indices = {curr_line_idx}

    curr_pt = t1_pt
    next_pt = (set([str(curr_line['QJZDH']).strip(), str(curr_line['ZJZDH']).strip()]) - {curr_pt}).pop()
    ordered_pts.append(next_pt)
    curr_pt = next_pt

    # 3. 沿环路拓扑依次寻径
    for _ in range(n - 1):
        found = False
        for idx, l in enumerate(lines):
            if idx in used_line_indices:
                continue
            s = set([str(l['QJZDH']).strip(), str(l['ZJZDH']).strip()])
            if curr_pt in s:
                used_line_indices.add(idx)
                ordered_lines.append(l)
                next_pt = (s - {curr_pt}).pop()
                if len(ordered_pts) < n:
                    ordered_pts.append(next_pt)
                curr_pt = next_pt
                found = True
                break
        if not found:
            return lines, [str(l.get('QJZDH', '')).strip() for l in lines]

    return ordered_lines, ordered_pts


def compute_1_2000_tfh_square(geom, crs=None):
    """
    根据 NY/T 2537 正方形分幅规则计算地块相交的所有 1:2000 图幅号。
    - 幅面规格：50cm×50cm，即地面实地 1km×1km 网格。
    - 图幅编号：按图廓西南角平面坐标公里数编制，X（北坐标/纵坐标）在前，Y（东坐标/横坐标）在后，中间用短横线连接，精确到 0.1km（例如 4125.0-403.0）。
    - 跨多个图幅：列出全部与地块重叠面积大于0的相交图幅号，按 X 从大到小（北到南）、Y 从小到大（西到东）排序，使用全角分号'；'分隔。
    """
    import math
    from shapely.geometry import box

    if geom is None or geom.is_empty:
        return ''

    # 若传入 CRS 且为地理坐标系（经纬度），正方形分幅无法直接按米/公里计算
    if crs is not None and getattr(crs, 'is_geographic', False):
        logger.warning(f"矢量地块坐标系为地理坐标系({crs})，无法直接计算1:2000正方形分幅图幅号")
        return ''

    minx, miny, maxx, maxy = geom.bounds

    # 1:2000 比例尺正方形分幅实地边长为 1000m
    sheet_size = 1000.0

    i_start = int(math.floor(miny / sheet_size))
    i_end = int(math.floor(maxy / sheet_size))
    j_start = int(math.floor(minx / sheet_size))
    j_end = int(math.floor(maxx / sheet_size))

    sheet_cells = []
    for i in range(i_start, i_end + 1):
        for j in range(j_start, j_end + 1):
            cell_poly = box(j * sheet_size, i * sheet_size, (j + 1) * sheet_size, (i + 1) * sheet_size)
            try:
                inter = geom.intersection(cell_poly)
                # 仅保留真正有面积重叠的图幅，排除仅点线接触
                if not inter.is_empty and getattr(inter, 'area', 0) > 1e-9:
                    sheet_cells.append((i, j))
            except Exception:
                pass

    if not sheet_cells:
        # 容错：若点或极端狭窄要素相交面积为0，取中心点所在图幅
        try:
            c = geom.centroid
            sheet_cells.append((int(math.floor(c.y / sheet_size)), int(math.floor(c.x / sheet_size))))
        except Exception:
            return ''

    # 排序：X（北坐标 i）从北到南（大到小）、Y（东坐标 j）从西到东（小到大）
    sheet_cells.sort(key=lambda item: (-item[0], item[1]))

    # 去重保留顺序
    seen = set()
    unique_cells = []
    for cell in sheet_cells:
        if cell not in seen:
            seen.add(cell)
            unique_cells.append(cell)

    # 格式化为 X.0-Y.0（西南角公里数，保留1位小数）
    labels = [f"{i * sheet_size / 1000.0:.1f}-{j * sheet_size / 1000.0:.1f}" for i, j in unique_cells]
    return '；'.join(labels)


def _parse_time_value(val):
    """
    规范化时间表单元格值；若为0、空值、非法格式或超出正常年份范围（1900-2100）一律返回 None。
    支持 datetime.date, datetime.datetime, pd.Timestamp, 标准日期字符串，
    以及处于合理区间（20000~80000，约1954~2119年）的 Excel 浮点/整数日期序列号。
    """
    import numbers
    import datetime

    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, float) and (pd.isna(val) or val in (float('inf'), float('-inf'))):
        return None

    dt = None
    if isinstance(val, str):
        s = val.strip()
        if not s or s in ('/', 'NaT', 'None', 'nan', '-', '--'):
            return None
        try:
            dt = pd.to_datetime(s, errors='coerce')
        except Exception:
            return None
    elif isinstance(val, numbers.Real):
        num = float(val)
        if num != num or num in (float('inf'), float('-inf')) or num <= 0:
            return None
        # Excel 日期序列号合理区间（约 1954-10-03 至 2119-01-05）
        if not (20000 <= num <= 80000):
            return None
        try:
            dt = pd.to_datetime(num, unit='D', origin='1899-12-30')
        except Exception:
            return None
    elif isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        dt = val
    else:
        try:
            dt = pd.to_datetime(val, errors='coerce')
        except Exception:
            return None

    if dt is None or pd.isna(dt):
        return None

    try:
        year = dt.year
    except Exception:
        return None

    if year < 1900 or year > 2100:
        return None

    return dt


def generate_yushe_yhzd_materials_only(mdb_file_path, shp_file_path, output_directory):
    """
    榆社县一户一档材料生成功能
    从MDB中读取 FBF, CBF, CBF_JTCY, CBHT 等表
    生成【发包方调查表】、【承包方调查表】、【承包经营权调查结果核实表】等一户一档材料
    """
    import copy
    import re
    os.makedirs(output_directory, exist_ok=True)
    logger.info(f"开始生成榆社县一户一档材料，输出目录: {output_directory}")

    template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Templates', '榆社资料模板')
    cbf_tpl_path = os.path.join(template_dir, '承包方调查表模版.xlsx')
    fbf_tpl_path = os.path.join(template_dir, '发包方调查表模版.xlsx')
    hs_tpl_path = os.path.join(template_dir, '承包经营权调查结果核实表模版.xlsx')
    gh_tpl_path = os.path.join(template_dir, '承包经营权公示结果归户表模版.xlsx')
    sm_tpl_path = os.path.join(template_dir, '公示无异议声明书模版.docx')
    hz_tpl_path = os.path.join(template_dir, '户主声明书模版.docx')
    sq_tpl_path = os.path.join(template_dir, '家庭承包申请书模版.docx')
    dk_tpl_path = os.path.join(template_dir, '承包地块调查表模版.xlsx')

    if not os.path.exists(cbf_tpl_path):
        logger.error(f"承包方调查表模版不存在: {cbf_tpl_path}")
        raise FileNotFoundError(f"承包方调查表模版不存在: {cbf_tpl_path}")
    if not os.path.exists(fbf_tpl_path):
        logger.error(f"发包方调查表模版不存在: {fbf_tpl_path}")
        raise FileNotFoundError(f"发包方调查表模版不存在: {fbf_tpl_path}")
    if not os.path.exists(hs_tpl_path):
        logger.error(f"承包经营权调查结果核实表模版不存在: {hs_tpl_path}")
        raise FileNotFoundError(f"承包经营权调查结果核实表模版不存在: {hs_tpl_path}")
    if not os.path.exists(gh_tpl_path):
        logger.error(f"承包经营权公示结果归户表模版不存在: {gh_tpl_path}")
        raise FileNotFoundError(f"承包经营权公示结果归户表模版不存在: {gh_tpl_path}")

    # 加载发包方与承包方变更调查统计表
    resources_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
    fbf_stat_path = os.path.join(resources_dir, '榆社县发包方信息变更调查统计.xlsx')
    cbf_stat_path = os.path.join(resources_dir, '榆社县承包方信息变更调查统计.xlsx')

    fbf_stat_dict = {}
    if os.path.exists(fbf_stat_path):
        try:
            logger.info(f"正在加载发包方变更调查统计: {fbf_stat_path}")
            df_fbf_stat = pd.read_excel(fbf_stat_path)
            for _, r in df_fbf_stat.iterrows():
                code = str(r['修改后代码']).strip()
                fbf_stat_dict[code] = r.to_dict()
        except Exception as e:
            logger.warning(f"读取发包方变更统计失败: {e}")
    else:
        logger.warning(f"发包方变更调查统计文件不存在: {fbf_stat_path}")

    cbf_stat_dict = {}
    if os.path.exists(cbf_stat_path):
        try:
            logger.info(f"正在加载承包方变更调查统计: {cbf_stat_path}")
            df_cbf_stat = pd.read_excel(cbf_stat_path)
            for _, r in df_cbf_stat.iterrows():
                code = str(r['承包方编码']).strip()
                cbf_stat_dict[code] = r.to_dict()
        except Exception as e:
            logger.warning(f"读取承包方变更统计失败: {e}")
    else:
        logger.warning(f"承包方变更调查统计文件不存在: {cbf_stat_path}")

    # 加载一户一档材料时间表
    time_table_path = os.path.join(resources_dir, '榆社县一户一档材料时间表.xlsx')
    yushe_time_dict = {}
    if os.path.exists(time_table_path):
        try:
            logger.info(f"正在加载一户一档材料时间表: {time_table_path}")
            wb_time = openpyxl.load_workbook(time_table_path, data_only=True)
            ws_time = wb_time.active
            for r in range(3, ws_time.max_row + 1):
                raw_code = ws_time.cell(r, 2).value
                if not raw_code:
                    continue
                try:
                    code_str = str(int(float(raw_code))).strip()
                except (ValueError, TypeError):
                    code_str = str(raw_code).strip()

                def get_dt_from_cell(col_idx):
                    val = ws_time.cell(r, col_idx).value
                    return _parse_time_value(val)

                yushe_time_dict[code_str] = {
                    'fbf_dc_dt': get_dt_from_cell(5),
                    'fbf_sh_dt': get_dt_from_cell(6),
                    'hz_rq_dt': get_dt_from_cell(7),
                    'cbf_db_dt': get_dt_from_cell(8),
                    'cbf_dc_dt': get_dt_from_cell(9),
                    'cbf_sh_dt': get_dt_from_cell(10),
                    'dk_db_dt': get_dt_from_cell(11),
                    'dk_dc_dt': get_dt_from_cell(12),
                    'dk_sh_dt': get_dt_from_cell(13),
                    'hs_dc_dt': get_dt_from_cell(14),
                    'hs_db_dt': get_dt_from_cell(15),
                    'hs_ry_dt': get_dt_from_cell(16),
                    'hs_sh_dt': get_dt_from_cell(17),
                    'sm_dt': get_dt_from_cell(18),
                    'gh_gs_dt': get_dt_from_cell(19),
                    'gh_db_dt': get_dt_from_cell(20),
                    'gh_sh_dt': get_dt_from_cell(21),
                    'sq_dt': get_dt_from_cell(22),
                }
        except Exception as e:
            logger.warning(f"读取一户一档材料时间表失败: {e}")
    else:
        logger.info(f"一户一档材料时间表不存在: {time_table_path}，将不进行任何日期写入操作")

    def format_date_str(dt):
        parsed = _parse_time_value(dt)
        return parsed.strftime('%Y年%m月%d日') if parsed is not None else ''

    def get_fbf_time_info(fbfbm_code):
        if not fbfbm_code or not yushe_time_dict:
            return None
        code_s = str(fbfbm_code).strip()
        return yushe_time_dict.get(code_s)

    def format_stat_num(val):
        if pd.isna(val) or str(val).strip() in ('/', '', 'nan', 'None'):
            return '/'
        try:
            f = float(val)
            if f.is_integer():
                return str(int(f))
            else:
                return f"{f:.2f}".rstrip('0').rstrip('.')
        except Exception:
            return str(val).strip()

    def clean_border_desc(text, direction_char):
        if pd.isna(text) or text is None:
            return ''
        s = str(text).strip()
        if not s or s in ('/', '无'):
            return ''
        for prefix in [f'{direction_char}：', f'{direction_char}:', f'{direction_char} :']:
            if s.startswith(prefix):
                s = s[len(prefix):].strip()
                break
        return s

    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + mdb_file_path + ';'
    )
    conn = pyodbc.connect(conn_str)

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=UserWarning)
            fbf_df = pd.read_sql_query('SELECT * FROM [FBF]', conn)
            cbf_df = pd.read_sql_query('SELECT * FROM [CBF]', conn)
            jtcy_df = pd.read_sql_query('SELECT * FROM [CBF_JTCY]', conn)
            try:
                cbht_df = pd.read_sql_query('SELECT * FROM [CBHT]', conn)
            except Exception:
                cbht_df = pd.DataFrame()
            try:
                cbdkxx_df = pd.read_sql_query('SELECT * FROM [CBDKXX]', conn)
            except Exception:
                cbdkxx_df = pd.DataFrame()

        # 读取矢量数据属性（四至等信息）以及 JZD / JZX
        merged_land_df = pd.DataFrame()
        vector_geom_by_dkbm = {}
        vector_crs = None
        jzd_df = gpd.GeoDataFrame()
        jzx_df = gpd.GeoDataFrame()
        if shp_file_path and os.path.exists(shp_file_path):
            try:
                shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
                all_vector_data = gpd.GeoDataFrame()
                for shp_file in shp_files_with_dk:
                    vdata = read_shp_safe(shp_file)
                    all_vector_data = pd.concat([all_vector_data, vdata], ignore_index=True)
                if not all_vector_data.empty:
                    vector_crs = getattr(all_vector_data, 'crs', None)
                    for _, vrow in all_vector_data.iterrows():
                        v_dkbm = str(vrow.get('DKBM', '')).strip()
                        v_geom = getattr(vrow, 'geometry', None)
                        if v_dkbm and v_geom is not None and not getattr(v_geom, 'is_empty', True):
                            if v_dkbm in vector_geom_by_dkbm:
                                try:
                                    vector_geom_by_dkbm[v_dkbm] = vector_geom_by_dkbm[v_dkbm].union(v_geom)
                                except Exception:
                                    pass
                            else:
                                vector_geom_by_dkbm[v_dkbm] = v_geom

                if not cbdkxx_df.empty and not all_vector_data.empty:
                    merged_land_df = pd.merge(cbdkxx_df, all_vector_data.drop(columns=['geometry'], errors='ignore'), on='DKBM', how='left')
                elif not cbdkxx_df.empty:
                    merged_land_df = cbdkxx_df.copy()
            except Exception as e:
                logger.warning(f"读取SHP矢量数据失败: {e}")
                if not cbdkxx_df.empty:
                    merged_land_df = cbdkxx_df.copy()

            # 读取 JZD 与 JZX 矢量数据
            try:
                for root_dir, _, filenames in os.walk(shp_file_path):
                    for fn in filenames:
                        fn_upper = fn.upper()
                        if fn_upper.endswith('.SHP'):
                            if fn_upper.startswith('JZD') or 'JZD' in fn_upper:
                                jzd_df = pd.concat([jzd_df, read_shp_safe(os.path.join(root_dir, fn))], ignore_index=True)
                            elif fn_upper.startswith('JZX') or 'JZX' in fn_upper:
                                jzx_df = pd.concat([jzx_df, read_shp_safe(os.path.join(root_dir, fn))], ignore_index=True)
            except Exception as e:
                logger.warning(f"读取JZD/JZX矢量数据失败: {e}")
        elif not cbdkxx_df.empty:
            merged_land_df = cbdkxx_df.copy()

        # -----------------------------
        # 1. 生成发包方调查表
        # -----------------------------
        logger.info(f"正在生成发包方调查表，共 {len(fbf_df)} 个发包方...")
        for _, fbf_row in fbf_df.iterrows():
            fbfbm = str(fbf_row['FBFBM']) if pd.notna(fbf_row['FBFBM']) else ''
            fbfmc = str(fbf_row['FBFMC']) if pd.notna(fbf_row['FBFMC']) else ''
            fbffzrxm = str(fbf_row['FBFFZRXM']) if pd.notna(fbf_row['FBFFZRXM']) else ''
            fzrzjhm = str(fbf_row['FZRZJHM']) if pd.notna(fbf_row['FZRZJHM']) else ''
            fbfdz = str(fbf_row['FBFDZ']) if pd.notna(fbf_row['FBFDZ']) else ''
            yzbm = str(fbf_row['YZBM']) if pd.notna(fbf_row['YZBM']) else ''
            lxdh = str(fbf_row['LXDH']) if pd.notna(fbf_row['LXDH']) else ''

            # 判断负责人性别：根据身份证号倒数第2位（第17位），奇数男，偶数女
            is_male = True
            clean_id = fzrzjhm.strip()
            if len(clean_id) == 18 and clean_id[:17].isdigit():
                gender_digit = int(clean_id[16])
                is_male = (gender_digit % 2 != 0)
            elif len(clean_id) == 15 and clean_id.isdigit():
                gender_digit = int(clean_id[14])
                is_male = (gender_digit % 2 != 0)

            sex_text = 'R男  £女' if is_male else '£男  R女'

            # 打开模版并填充数据
            wb_fbf = openpyxl.load_workbook(fbf_tpl_path)
            ws_fbf = wb_fbf.active

            ws_fbf['D4'].value = fbfmc
            ws_fbf['G4'].value = fbfbm
            ws_fbf['D5'].value = fbffzrxm
            ws_fbf['G5'].value = sex_text
            ws_fbf['G5'].font = openpyxl.styles.Font(name='Wingdings 2', size=11)
            ws_fbf['D6'].value = fbfdz
            ws_fbf['D7'].value = yzbm
            ws_fbf['G7'].value = lxdh
            ws_fbf['D9'].value = fzrzjhm

            # 填充发包方调查记事
            fbf_stat = fbf_stat_dict.get(fbfbm)
            if fbf_stat:
                y_fbfmc = str(fbf_stat.get('原发包方名称', '')).strip() if pd.notna(fbf_stat.get('原发包方名称')) else fbfmc
                x_fbfmc = str(fbf_stat.get('现发包方名称', '')).strip() if pd.notna(fbf_stat.get('现发包方名称')) else fbfmc
                y_fzr = str(fbf_stat.get('原发包方负责人', '')).strip() if pd.notna(fbf_stat.get('原发包方负责人')) else fbffzrxm
                x_fzr = str(fbf_stat.get('现发包方负责人', '')).strip() if pd.notna(fbf_stat.get('现发包方负责人')) else fbffzrxm

                y_cbf_gs = fbf_stat.get('原承包方个数', 0)
                x_cbf_gs = fbf_stat.get('现延包方个数', 0)
                try:
                    diff_num = int(x_cbf_gs) - int(y_cbf_gs)
                    if diff_num >= 0:
                        diff_text = f"分户新增{diff_num}个"
                    else:
                        diff_text = f"减少{-diff_num}个"
                except Exception:
                    diff_text = "分户新增0个"

                yb_td = fbf_stat.get('延包承包土地（亩）', 0)
                fyb_td = fbf_stat.get('非承包土地（亩）', 0)
                try:
                    total_land = float(yb_td) + float(fyb_td)
                    total_land_str = format_stat_num(total_land)
                except Exception:
                    total_land_str = format_stat_num(yb_td)

                yb_td_str = format_stat_num(yb_td)
                fyb_td_str = format_stat_num(fyb_td)
                y_cbf_str = format_stat_num(y_cbf_gs)
                x_cbf_str = format_stat_num(x_cbf_gs)

                fbfdcjs = (
                    f"1、原发包方名称:{y_fbfmc}，现发包方名称:{x_fbfmc}；\n"
                    f"2、原发包方负责人:{y_fzr}，现发包方负责人:{x_fzr}；\n"
                    f"3、原承包方{y_cbf_str}个，现延包方{x_cbf_str}个;{diff_text}；\n"
                    f"4、本次延包耕地共{total_land_str}亩，其中延包承包土地共{yb_td_str}亩，非承包土地{fyb_td_str}亩。"
                )
                ws_fbf['B10'].value = fbfdcjs

            # 填写发包方调查表时间（G11->调查记事时间XXXX年XX月XX日，G13->审核意见时间XXXX年XX月XX日）
            time_info_fbf = get_fbf_time_info(fbfbm)
            if time_info_fbf:
                dt_fbf_dc = format_date_str(time_info_fbf.get('fbf_dc_dt'))
                dt_fbf_sh = format_date_str(time_info_fbf.get('fbf_sh_dt'))
                if dt_fbf_dc:
                    ws_fbf['G11'].value = dt_fbf_dc
                if dt_fbf_sh:
                    ws_fbf['G13'].value = dt_fbf_sh

            # 导出文件名：{FBFBM}_{FBFFZRXM}_发包方调查表.xlsx
            fbf_filename = f"{fbfbm}_{fbffzrxm}_发包方调查表.xlsx"
            fbf_dir = os.path.join(output_directory, fbfbm) if fbfbm else output_directory
            os.makedirs(fbf_dir, exist_ok=True)
            fbf_output_path = os.path.join(fbf_dir, fbf_filename)
            wb_fbf.save(fbf_output_path)
            logger.debug(f"已生成发包方调查表: {fbf_filename}")

        # -----------------------------
        # 2. 生成承包方调查表
        # -----------------------------
        logger.info(f"正在生成承包方调查表，共 {len(cbf_df)} 户承包方...")

        # 整理合同期限与编号映射
        cbht_map = {}
        if not cbht_df.empty:
            for _, ht_row in cbht_df.iterrows():
                cbfbm_val = str(ht_row['CBFBM']) if pd.notna(ht_row['CBFBM']) else ''
                if not cbfbm_val:
                    continue
                cbqxq = ht_row.get('CBQXQ')
                cbqxz = ht_row.get('CBQXZ')
                cbqxq_str = pd.to_datetime(cbqxq).strftime('%Y年%m月%d日') if pd.notna(cbqxq) else ''
                cbqxz_str = pd.to_datetime(cbqxz).strftime('%Y年%m月%d日') if pd.notna(cbqxz) else ''
                q_dt = pd.to_datetime(cbqxq) if pd.notna(cbqxq) else None
                z_dt = pd.to_datetime(cbqxz) if pd.notna(cbqxz) else None
                cbht_map[cbfbm_val] = {
                    'CBQXQ': cbqxq_str,
                    'CBQXZ': cbqxz_str,
                    'Q_Y': str(q_dt.year) if q_dt else '2024',
                    'Q_M': f"{q_dt.month:02d}" if q_dt else '01',
                    'Q_D': f"{q_dt.day:02d}" if q_dt else '01',
                    'Z_Y': str(z_dt.year) if z_dt else '2053',
                    'Z_M': f"{z_dt.month:02d}" if z_dt else '12',
                    'Z_D': f"{z_dt.day:02d}" if z_dt else '31',
                }

        # 循环每户承包方
        for _, cbf_row in cbf_df.iterrows():
            cbfbm = str(cbf_row['CBFBM']) if pd.notna(cbf_row['CBFBM']) else ''
            cbfmc = str(cbf_row['CBFMC']) if pd.notna(cbf_row['CBFMC']) else ''
            cbfdz = str(cbf_row['CBFDZ']) if pd.notna(cbf_row['CBFDZ']) else ''
            yzbm = str(cbf_row['YZBM']) if pd.notna(cbf_row['YZBM']) else ''
            lxdh = str(cbf_row['LXDH']) if pd.notna(cbf_row['LXDH']) else ''
            cbfzjhm = str(cbf_row['CBFZJHM']) if pd.notna(cbf_row['CBFZJHM']) and str(cbf_row['CBFZJHM']).strip() != '' else '/'

            # 发包方编码：CBFBM 前 14 位
            fbfbm = cbfbm[:14] if len(cbfbm) >= 14 else ''
            # 承包方编码后4位
            cbfbm_last4 = cbfbm[-4:] if len(cbfbm) >= 4 else cbfbm
            # 承包合同编号：CBFBM + "J"
            cbhtbh = f"{cbfbm}J" if cbfbm else ''

            # 承包期限
            ht_info = cbht_map.get(cbfbm, {})
            cbqxq_str = ht_info.get('CBQXQ', '2024年01月01日')
            cbqxz_str = ht_info.get('CBQXZ', '2053年12月31日')
            cbqx_text = f"{cbqxq_str} 至 {cbqxz_str}"

            # 本户家庭成员
            members = jtcy_df[jtcy_df['CBFBM'] == cbfbm].copy()
            cbfcysl = len(members)

            # 打开承包方调查表模版
            wb_cbf = openpyxl.load_workbook(cbf_tpl_path)
            ws_cbf = wb_cbf.active

            # 填充基本信息单元格
            ws_cbf['D3'].value = fbfbm
            ws_cbf['K3'].value = cbfbm_last4
            ws_cbf['D4'].value = cbfmc
            ws_cbf['K4'].value = lxdh
            ws_cbf['D5'].value = cbfdz
            ws_cbf['M5'].value = yzbm
            ws_cbf['M7'].value = cbfzjhm
            ws_cbf['H9'].value = cbhtbh
            ws_cbf['H11'].value = cbhtbh
            ws_cbf['D13'].value = cbqx_text
            ws_cbf['Q14'].value = str(cbfcysl)

            # 处理成员列表行插入：
            # 模版初始预留了12行（17行~28行）
            # 如果成员个数 22 > cbfcysl > 12，则在17行位置插入10行
            # 如果成员个数 cbfcysl > 22，则在第17行位置插入 cbfcysl - 12 行
            num_insert = 0
            if 12 < cbfcysl < 22:
                num_insert = 10
            elif cbfcysl >= 22:
                num_insert = cbfcysl - 12

            if num_insert > 0:
                insert_rows_with_merges(ws_cbf, 17, num_insert)
                ref_row = 17 + num_insert
                for r in range(17, 17 + num_insert):
                    ws_cbf.row_dimensions[r].height = ws_cbf.row_dimensions[ref_row].height
                    for c in range(1, 19):
                        src_cell = ws_cbf.cell(ref_row, c)
                        tgt_cell = ws_cbf.cell(r, c)
                        if src_cell.has_style:
                            tgt_cell.font = copy.copy(src_cell.font)
                            tgt_cell.border = copy.copy(src_cell.border)
                            tgt_cell.fill = copy.copy(src_cell.fill)
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.protection = copy.copy(src_cell.protection)
                            tgt_cell.alignment = copy.copy(src_cell.alignment)
                    # 恢复该行合并单元格
                    ws_cbf.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
                    ws_cbf.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
                    ws_cbf.merge_cells(start_row=r, start_column=10, end_row=r, end_column=14)
                    ws_cbf.merge_cells(start_row=r, start_column=15, end_row=r, end_column=18)

            # 填充家庭成员数据
            # 成员排序：优先本人/户主排在第1位，其他成员按原有顺序
            def member_sort_key(m_row):
                y = str(m_row.get('YHZGX', '')).strip()
                return 0 if y in ('01', '02', '本人', '户主') else 1

            sorted_members = sorted(members.to_dict('records'), key=member_sort_key)

            for idx, m in enumerate(sorted_members):
                curr_row = 17 + idx
                cyxm = str(m.get('CYXM', '')) if pd.notna(m.get('CYXM')) else ''
                cyxb = xb_code_to_name(m.get('CYXB', ''))

                cyzjhm = str(m.get('CYZJHM', '')).strip() if pd.notna(m.get('CYZJHM')) else ''

                # 从CYZJHM解析出生日期
                csrq = ''
                if len(cyzjhm) == 18 and cyzjhm[:14].isdigit():
                    birth_year = cyzjhm[6:10]
                    birth_month = cyzjhm[10:12]
                    birth_day = cyzjhm[12:14]
                    csrq = f"{birth_year}年{birth_month}月{birth_day}日"
                elif len(cyzjhm) == 15 and cyzjhm[:12].isdigit():
                    birth_year = "19" + cyzjhm[6:8]
                    birth_month = cyzjhm[8:10]
                    birth_day = cyzjhm[10:12]
                    csrq = f"{birth_year}年{birth_month}月{birth_day}日"

                # 与户主关系
                yhzgx_raw = str(m.get('YHZGX', '')).strip() if pd.notna(m.get('YHZGX')) else ''
                yhzgx = yhzgx_code_to_name(yhzgx_raw) if yhzgx_raw.isdigit() else yhzgx_raw

                is_head = yhzgx in ('本人', '户主') or idx == 0

                # 联系电话：只填写户主或本人，其他为 "/"
                m_lxdh = lxdh if is_head else '/'

                # 证件名称：户主或本人为“居民身份证”，其他成员为“户口簿”
                zjmc = '居民身份证' if is_head else '户口簿'

                # 写入单元格
                ws_cbf.cell(curr_row, 1).value = cyxm
                ws_cbf.cell(curr_row, 2).value = cyxb
                ws_cbf.cell(curr_row, 3).value = csrq
                ws_cbf.cell(curr_row, 6).value = yhzgx
                ws_cbf.cell(curr_row, 7).value = m_lxdh
                ws_cbf.cell(curr_row, 9).value = zjmc
                ws_cbf.cell(curr_row, 10).value = cyzjhm

            # 填充承包方调查记事
            cbf_stat = cbf_stat_dict.get(cbfbm)
            if cbf_stat:
                y_cbf = str(cbf_stat.get('原承包方', '')).strip() if pd.notna(cbf_stat.get('原承包方')) else ''
                x_cbf = str(cbf_stat.get('现二轮延包承包方', '')).strip() if pd.notna(cbf_stat.get('现二轮延包承包方')) else cbfmc

                is_new_cbf = (not y_cbf) or (y_cbf in ('/', '无', '新增', 'nan', 'None'))

                y_mj_str = format_stat_num(cbf_stat.get('原承包土地面积（亩）'))
                y_ks_str = format_stat_num(cbf_stat.get('原承包土地地块数'))
                x_mj_str = format_stat_num(cbf_stat.get('延包土地面积（亩）'))
                x_ks_str = format_stat_num(cbf_stat.get('延包土地地块数'))

                if is_new_cbf:
                    cbfdcjs = (
                        f"1、原承包方不存在，为分户新增承包方，现二轮延包承包方为{x_cbf}。"
                        f"2、延包土地{x_mj_str}亩，{x_ks_str}块。"
                    )
                else:
                    cbfdcjs = (
                        f"1、原承包方为{y_cbf}，现二轮延包承包方为{x_cbf}。"
                        f"2、原承包土地{y_mj_str}亩，{y_ks_str}块；延包土地{x_mj_str}亩，{x_ks_str}块。"
                    )
            else:
                cbfdcjs = f"1、原承包方不存在，为分户新增承包方，现二轮延包承包方为{cbfmc}。"

            # 写入调查记事单元格：Row 29 + num_insert, Col 10 (J)
            ws_cbf.cell(29 + num_insert, 10).value = cbfdcjs

            # 承包方（代表）时间->承包方（代表）栏F列；调查记事时间->调查记事栏的P列；审核意见时间->审核意见栏的M列
            time_info = get_fbf_time_info(fbfbm)
            if time_info:
                dt_cbf_db = format_date_str(time_info.get('cbf_db_dt'))
                dt_cbf_dc = format_date_str(time_info.get('cbf_dc_dt'))
                dt_cbf_sh = format_date_str(time_info.get('cbf_sh_dt'))
                if dt_cbf_db:
                    ws_cbf.cell(29 + num_insert, 6).value = dt_cbf_db
                if dt_cbf_dc:
                    ws_cbf.cell(32 + num_insert, 16).value = dt_cbf_dc
                if dt_cbf_sh:
                    ws_cbf.cell(38 + num_insert, 13).value = dt_cbf_sh

            # 导出文件名：{CBFBM}_{CBFMC}_承包方调查表.xlsx
            cbf_filename = f"{cbfbm}_{cbfmc}_承包方调查表.xlsx"
            fbf_dir = os.path.join(output_directory, fbfbm) if fbfbm else output_directory
            safe_cbfmc = cbfmc.replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?', '').replace('\"', '').replace('<', '').replace('>', '').replace('|', '').strip()
            cbf_dir = os.path.join(fbf_dir, f"{cbfbm}_{safe_cbfmc}")
            os.makedirs(cbf_dir, exist_ok=True)
            cbf_output_path = os.path.join(cbf_dir, cbf_filename)
            wb_cbf.save(cbf_output_path)

        # -----------------------------
        # 3. 生成承包经营权调查结果核实表
        # -----------------------------
        logger.info(f"正在生成承包经营权调查结果核实表，共 {len(cbf_df)} 户承包方...")
        fbf_map = {}
        for _, f_row in fbf_df.iterrows():
            f_bm = str(f_row.get('FBFBM', '')).strip()
            fbf_map[f_bm] = {
                'FBFMC': str(f_row.get('FBFMC', '')) if pd.notna(f_row.get('FBFMC')) else '',
                'FBFFZRXM': str(f_row.get('FBFFZRXM', '')) if pd.notna(f_row.get('FBFFZRXM')) else '',
            }

        for _, cbf_row in cbf_df.iterrows():
            cbfbm = str(cbf_row['CBFBM']) if pd.notna(cbf_row['CBFBM']) else ''
            cbfmc = str(cbf_row['CBFMC']) if pd.notna(cbf_row['CBFMC']) else ''
            cbfdz = str(cbf_row['CBFDZ']) if pd.notna(cbf_row['CBFDZ']) else ''
            yzbm = str(cbf_row['YZBM']) if pd.notna(cbf_row['YZBM']) else ''
            lxdh = str(cbf_row['LXDH']) if pd.notna(cbf_row['LXDH']) else ''
            cbfzjhm = str(cbf_row['CBFZJHM']) if pd.notna(cbf_row['CBFZJHM']) and str(cbf_row['CBFZJHM']).strip() != '' else '/'

            # 发包方信息
            fbfbm = cbfbm[:14] if len(cbfbm) >= 14 else ''
            fbf_info = fbf_map.get(fbfbm, {})
            fbfmc = fbf_info.get('FBFMC', '')
            fbffzrxm = fbf_info.get('FBFFZRXM', '')
            time_info = get_fbf_time_info(fbfbm)

            # 合同编号与期限
            cbhtbh = f"{cbfbm}J" if cbfbm else ''
            ht_info = cbht_map.get(cbfbm, {})
            cbqxq_str = ht_info.get('CBQXQ', '2024年01月01日')
            cbqxz_str = ht_info.get('CBQXZ', '2053年12月31日')
            cbqx_text = f"{cbqxq_str} 至 {cbqxz_str}"

            # 成员信息
            members = jtcy_df[jtcy_df['CBFBM'] == cbfbm].copy()
            cbfcysl = len(members)

            # 生成该户的调查/公示记事内容（供核实表与归户表共用）
            cbf_stat = cbf_stat_dict.get(cbfbm)
            if cbf_stat:
                y_cbf = str(cbf_stat.get('原承包方', '')).strip() if pd.notna(cbf_stat.get('原承包方')) else ''
                x_cbf = str(cbf_stat.get('现二轮延包承包方', '')).strip() if pd.notna(cbf_stat.get('现二轮延包承包方')) else cbfmc
                is_new_cbf = (not y_cbf) or (y_cbf in ('/', '无', '新增', 'nan', 'None'))
                y_mj_str = format_stat_num(cbf_stat.get('原承包土地面积（亩）'))
                y_ks_str = format_stat_num(cbf_stat.get('原承包土地地块数'))
                x_mj_str = format_stat_num(cbf_stat.get('延包土地面积（亩）'))
                x_ks_str = format_stat_num(cbf_stat.get('延包土地地块数'))

                if is_new_cbf:
                    cbfdcjs_hs = (
                        f"1、原承包方不存在，为分户新增承包方，现二轮延包承包方为{x_cbf}。"
                        f"2、延包土地{x_mj_str}亩，{x_ks_str}块。"
                    )
                else:
                    cbfdcjs_hs = (
                        f"1、原承包方为{y_cbf}，现二轮延包承包方为{x_cbf}。"
                        f"2、原承包土地{y_mj_str}亩，{y_ks_str}块；延包土地{x_mj_str}亩，{x_ks_str}块。"
                    )
            else:
                cbfdcjs_hs = f"1、原承包方不存在，为分户新增承包方，现二轮延包承包方为{cbfmc}。"

            # 打开核实表模版
            wb_hs = openpyxl.load_workbook(hs_tpl_path)
            ws_hs = wb_hs.active

            # 锁定全表模板行高（未显式设行高的行赋默认行高17.5pt），确保保存时保留 customHeight="1"，避免 Excel AutoFit 跑版
            default_h_hs = ws_hs.sheet_format.defaultRowHeight or 17.5
            for r in range(1, ws_hs.max_row + 1):
                if ws_hs.row_dimensions[r].height is None:
                    ws_hs.row_dimensions[r].height = default_h_hs

            # 填充基本信息单元格
            ws_hs['E3'].value = fbfmc
            ws_hs['N3'].value = fbffzrxm
            ws_hs['E5'].value = cbhtbh
            ws_hs['E7'].value = cbqx_text
            ws_hs['E9'].value = cbfmc
            ws_hs['N9'].value = lxdh
            ws_hs['N10'].value = cbfzjhm
            ws_hs['E13'].value = cbfdz
            ws_hs['Q13'].value = yzbm
            ws_hs['O15'].value = f"共 {cbfcysl} 人"

            # 家庭成员信息：startrow=18;如果[cbfcysl]>25，则在43行粘贴标题行，之后在45行插入40行
            if cbfcysl > 25:
                insert_rows_with_merges(ws_hs, 43, 2)
                for r_offset in [0, 1]:
                    src_r = 16 + r_offset
                    tgt_r = 43 + r_offset
                    ws_hs.row_dimensions[tgt_r].height = ws_hs.row_dimensions[src_r].height
                    for c in range(1, 18):
                        src_cell = ws_hs.cell(src_r, c)
                        tgt_cell = ws_hs.cell(tgt_r, c)
                        tgt_cell.value = src_cell.value
                        if src_cell.has_style:
                            tgt_cell.font = copy.copy(src_cell.font)
                            tgt_cell.border = copy.copy(src_cell.border)
                            tgt_cell.fill = copy.copy(src_cell.fill)
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.protection = copy.copy(src_cell.protection)
                            tgt_cell.alignment = copy.copy(src_cell.alignment)

                ws_hs.merge_cells('A43:A44')
                ws_hs.merge_cells('B43:B44')
                ws_hs.merge_cells('C43:D44')
                ws_hs.merge_cells('E43:F44')
                ws_hs.merge_cells('G43:I44')
                ws_hs.merge_cells('J43:K44')
                ws_hs.merge_cells('L43:O44')
                ws_hs.merge_cells('P43:Q44')

                insert_rows_with_merges(ws_hs, 45, 40)
                for r in range(45, 85):
                    ws_hs.row_dimensions[r].height = ws_hs.row_dimensions[18].height
                    for c in range(1, 18):
                        src_cell = ws_hs.cell(18, c)
                        tgt_cell = ws_hs.cell(r, c)
                        if src_cell.has_style:
                            tgt_cell.font = copy.copy(src_cell.font)
                            tgt_cell.border = copy.copy(src_cell.border)
                            tgt_cell.fill = copy.copy(src_cell.fill)
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.protection = copy.copy(src_cell.protection)
                            tgt_cell.alignment = copy.copy(src_cell.alignment)
                    ws_hs.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
                    ws_hs.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
                    ws_hs.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
                    ws_hs.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)
                    ws_hs.merge_cells(start_row=r, start_column=12, end_row=r, end_column=15)
                    ws_hs.merge_cells(start_row=r, start_column=16, end_row=r, end_column=17)
                cy_endrow = 84
            else:
                cy_endrow = 42

            # 填充成员数据
            def member_sort_key_hs(m_row):
                y = str(m_row.get('YHZGX', '')).strip()
                return 0 if y in ('01', '02', '本人', '户主') else 1

            sorted_members_hs = sorted(members.to_dict('records'), key=member_sort_key_hs)

            for idx, m in enumerate(sorted_members_hs):
                if idx < 25:
                    curr_row = 18 + idx
                elif idx < 65:
                    curr_row = 45 + (idx - 25)
                else:
                    break

                cyxm = str(m.get('CYXM', '')) if pd.notna(m.get('CYXM')) else ''
                cyxb = xb_code_to_name(m.get('CYXB', ''))
                cyzjhm = str(m.get('CYZJHM', '')).strip() if pd.notna(m.get('CYZJHM')) else ''

                csrq = ''
                if len(cyzjhm) == 18 and cyzjhm[:14].isdigit():
                    birth_year = cyzjhm[6:10]
                    birth_month = cyzjhm[10:12]
                    birth_day = cyzjhm[12:14]
                    csrq = f"{birth_year}年{birth_month}月{birth_day}日"
                elif len(cyzjhm) == 15 and cyzjhm[:12].isdigit():
                    birth_year = "19" + cyzjhm[6:8]
                    birth_month = cyzjhm[8:10]
                    birth_day = cyzjhm[10:12]
                    csrq = f"{birth_year}年{birth_month}月{birth_day}日"

                yhzgx_raw = str(m.get('YHZGX', '')).strip() if pd.notna(m.get('YHZGX')) else ''
                yhzgx = yhzgx_code_to_name(yhzgx_raw) if yhzgx_raw.isdigit() else yhzgx_raw

                is_head = yhzgx in ('本人', '户主')
                m_lxdh = lxdh if is_head else '/'
                zjmc = '居民身份证' if is_head else '户口簿'

                ws_hs.cell(curr_row, 1).value = cyxm
                ws_hs.cell(curr_row, 2).value = cyxb
                ws_hs.cell(curr_row, 3).value = csrq
                ws_hs.cell(curr_row, 5).value = yhzgx
                ws_hs.cell(curr_row, 7).value = m_lxdh
                ws_hs.cell(curr_row, 10).value = zjmc
                ws_hs.cell(curr_row, 12).value = cyzjhm
                ws_hs.cell(curr_row, 16).value = '/'

            # 承包地块情况
            page1_start = cy_endrow + 1
            page_len = 38

            rel_merges = []
            for mr in ws_hs.merged_cells.ranges:
                if page1_start <= mr.min_row <= page1_start + page_len - 1 and page1_start <= mr.max_row <= page1_start + page_len - 1:
                    rel_merges.append((mr.min_col, mr.min_row - page1_start, mr.max_col, mr.max_row - page1_start))

            cbf_lands = merged_land_df[merged_land_df['CBFBM'] == cbfbm].sort_values('DKBM') if not merged_land_df.empty else pd.DataFrame()
            num_parcels = len(cbf_lands)
            extra_pages = (num_parcels - 1) // 8 if num_parcels > 0 else 0

            # 如果[CBDKZS]>8，增加整页“承包地块情况”行
            for ep in range(extra_pages):
                insert_pos = page1_start + page_len * (ep + 1)
                insert_rows_with_merges(ws_hs, insert_pos, page_len)
                for i in range(page_len):
                    src_r = page1_start + i
                    tgt_r = insert_pos + i
                    ws_hs.row_dimensions[tgt_r].height = ws_hs.row_dimensions[src_r].height
                    for c in range(1, 18):
                        src_cell = ws_hs.cell(src_r, c)
                        tgt_cell = ws_hs.cell(tgt_r, c)
                        if src_cell.has_style:
                            tgt_cell.font = copy.copy(src_cell.font)
                            tgt_cell.border = copy.copy(src_cell.border)
                            tgt_cell.fill = copy.copy(src_cell.fill)
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.protection = copy.copy(src_cell.protection)
                            tgt_cell.alignment = copy.copy(src_cell.alignment)
                        if i in [0, 1, 2, 3, 4] or src_cell.value in ['东 :', '南 :', '西 :', '北 :']:
                            tgt_cell.value = src_cell.value
                for min_col, min_r, max_col, max_r in rel_merges:
                    ws_hs.merge_cells(start_row=insert_pos + min_r, start_column=min_col, end_row=insert_pos + max_r, end_column=max_col)

            # 地块汇总：B{cy_endrow+2}：地块数量；G{cy_endrow+2}：合同总面积亩；N{cy_endrow+2}：实测总面积亩
            ht_rows = cbht_df[cbht_df['CBFBM'] == cbfbm]
            if not ht_rows.empty:
                cbdkzs = ht_rows.iloc[0].get('CBDKZS', num_parcels)
                htzmjm = ht_rows.iloc[0].get('HTZMJM', 0.0)
            else:
                cbdkzs = num_parcels
                htzmjm = cbf_lands['HTMJM'].sum() if not cbf_lands.empty and 'HTMJM' in cbf_lands.columns else 0.0

            total_pages = 1 + extra_pages
            for p in range(total_pages):
                p_sum_row = page1_start + p * page_len + 1
                ws_hs.cell(p_sum_row, 2).value = str(cbdkzs)
                ws_hs.cell(p_sum_row, 7).value = f"{float(htzmjm):.2f}亩"
                ws_hs.cell(p_sum_row, 14).value = f"{float(htzmjm):.2f}亩"

            # 填充地块槽位：每4行一个地块
            for idx, land in enumerate(cbf_lands.to_dict('records')):
                page_num = idx // 8
                slot_in_page = idx % 8
                r0 = page1_start + page_num * page_len + 5 + slot_in_page * 4

                dkbm_val = str(land.get('DKBM', '')).strip()
                dkbm_last5 = dkbm_val[-5:] if len(dkbm_val) >= 5 else dkbm_val
                dkmc = str(land.get('DKMC', '')).strip() if pd.notna(land.get('DKMC')) else ''

                clean_east = clean_border_desc(land.get('DKDZ'), '东')
                clean_south = clean_border_desc(land.get('DKNZ'), '南')
                clean_west = clean_border_desc(land.get('DKXZ'), '西')
                clean_north = clean_border_desc(land.get('DKBZ'), '北')

                htmjm_val = land.get('HTMJM')
                htmjm_str = f"{float(htmjm_val):.2f}" if pd.notna(htmjm_val) and str(htmjm_val).strip() != '' else ''

                scmjm_val = land.get('SCMJM')
                scmjm_str = f"{float(scmjm_val):.2f}" if pd.notna(scmjm_val) and str(scmjm_val).strip() != '' else htmjm_str

                tdyt_str = tdyt_code_to_name(land.get('TDYT', ''))

                dklb_str = dklb_code_to_name(land.get('DKLB', ''))
                tdlylx_str = tdlylx_code_to_name(land.get('TDLYLX', ''))
                dldj_str = dldj_code_to_name(land.get('DLDJ', ''))
                sfjbnt_str = sfjbnt_code_to_name(land.get('SFJBNT', ''))

                ws_hs.cell(r0, 1).value = dkbm_last5
                ws_hs.cell(r0, 2).value = dkmc
                ws_hs.cell(r0, 4).value = f"东：{clean_east}" if clean_east else "东 :"
                ws_hs.cell(r0 + 1, 4).value = f"南：{clean_south}" if clean_south else "南 :"
                ws_hs.cell(r0 + 2, 4).value = f"西：{clean_west}" if clean_west else "西 :"
                ws_hs.cell(r0 + 3, 4).value = f"北：{clean_north}" if clean_north else "北 :"
                ws_hs.cell(r0, 9).value = htmjm_str
                ws_hs.cell(r0, 10).value = scmjm_str
                ws_hs.cell(r0, 11).value = tdyt_str
                ws_hs.cell(r0, 12).value = dklb_str
                ws_hs.cell(r0, 13).value = tdlylx_str
                ws_hs.cell(r0, 14).value = dldj_str
                ws_hs.cell(r0, 15).value = sfjbnt_str
                ws_hs.cell(r0, 17).value = '家庭承包'

            # 填充“非承包地块情况”页的调查记事
            jishi_row = None
            for r in range(page1_start + total_pages * page_len, ws_hs.max_row + 1):
                if ws_hs.cell(r, 1).value == '调查\n记事':
                    jishi_row = r
                    break

            if jishi_row is not None:
                ws_hs.cell(jishi_row, 2).value = cbfdcjs_hs
                # 调查记事时间->调查记事栏的N列；调查人员时间->调查人员栏的N列；承包方（代表）时间->承包方（代表）栏的F列；审核意见时间->审核意见栏的N列
                if time_info:
                    dt_hs_dc = format_date_str(time_info.get('hs_dc_dt'))
                    dt_hs_db = format_date_str(time_info.get('hs_db_dt'))
                    dt_hs_ry = format_date_str(time_info.get('hs_ry_dt'))
                    dt_hs_sh = format_date_str(time_info.get('hs_sh_dt'))
                    if dt_hs_dc:
                        ws_hs.cell(jishi_row + 4, 14).value = dt_hs_dc
                    if dt_hs_db:
                        ws_hs.cell(jishi_row + 9, 6).value = dt_hs_db
                    if dt_hs_ry:
                        ws_hs.cell(jishi_row + 9, 14).value = dt_hs_ry
                    if dt_hs_sh:
                        ws_hs.cell(jishi_row + 14, 14).value = dt_hs_sh

            # 导出文件名：{CBFBM}_{CBFMC}_承包经营权调查结果核实表.xlsx
            hs_filename = f"{cbfbm}_{cbfmc}_承包经营权调查结果核实表.xlsx"
            fbf_dir = os.path.join(output_directory, fbfbm) if fbfbm else output_directory
            safe_cbfmc = cbfmc.replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?', '').replace('\"', '').replace('<', '').replace('>', '').replace('|', '').strip()
            cbf_dir = os.path.join(fbf_dir, f"{cbfbm}_{safe_cbfmc}")
            os.makedirs(cbf_dir, exist_ok=True)
            hs_output_path = os.path.join(cbf_dir, hs_filename)
            wb_hs.save(hs_output_path)

            # -----------------------------
            # 4. 生成承包经营权公示结果归户表
            # -----------------------------
            wb_gh = openpyxl.load_workbook(gh_tpl_path)
            ws_gh = wb_gh.active

            # 锁定全表模板行高（未显式设行高的行赋默认行高17.5pt），确保保存时保留 customHeight="1"，避免 Excel AutoFit 跑版
            default_h_gh = ws_gh.sheet_format.defaultRowHeight or 17.5
            for r in range(1, ws_gh.max_row + 1):
                if ws_gh.row_dimensions[r].height is None:
                    ws_gh.row_dimensions[r].height = default_h_gh

            # 填充基本信息单元格
            ws_gh['E3'].value = fbfmc
            ws_gh['L3'].value = fbffzrxm
            ws_gh['E5'].value = cbhtbh
            ws_gh['E7'].value = cbqx_text
            ws_gh['E9'].value = cbfmc
            ws_gh['L9'].value = lxdh
            ws_gh['L10'].value = cbfzjhm
            ws_gh['E13'].value = cbfdz
            ws_gh['O13'].value = yzbm
            ws_gh['M14'].value = f"共 {cbfcysl} 人"

            # 家庭成员信息：startrow=17;如果[cbfcysl]>26，则在43行粘贴标题行，之后在45行插入40行
            if cbfcysl > 26:
                insert_rows_with_merges(ws_gh, 43, 2)
                for r_offset in [0, 1]:
                    src_r = 15 + r_offset
                    tgt_r = 43 + r_offset
                    ws_gh.row_dimensions[tgt_r].height = ws_gh.row_dimensions[src_r].height
                    for c in range(1, 20):
                        src_cell = ws_gh.cell(src_r, c)
                        tgt_cell = ws_gh.cell(tgt_r, c)
                        tgt_cell.value = src_cell.value
                        if src_cell.has_style:
                            tgt_cell.font = copy.copy(src_cell.font)
                            tgt_cell.border = copy.copy(src_cell.border)
                            tgt_cell.fill = copy.copy(src_cell.fill)
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.protection = copy.copy(src_cell.protection)
                            tgt_cell.alignment = copy.copy(src_cell.alignment)

                ws_gh.merge_cells('A43:A44')
                ws_gh.merge_cells('B43:B44')
                ws_gh.merge_cells('C43:D44')
                ws_gh.merge_cells('E43:F44')
                ws_gh.merge_cells('G43:G44')
                ws_gh.merge_cells('H43:I44')
                ws_gh.merge_cells('J43:O44')
                ws_gh.merge_cells('P43:S44')

                insert_rows_with_merges(ws_gh, 45, 40)
                for r in range(45, 85):
                    ws_gh.row_dimensions[r].height = ws_gh.row_dimensions[17].height
                    for c in range(1, 20):
                        src_cell = ws_gh.cell(17, c)
                        tgt_cell = ws_gh.cell(r, c)
                        if src_cell.has_style:
                            tgt_cell.font = copy.copy(src_cell.font)
                            tgt_cell.border = copy.copy(src_cell.border)
                            tgt_cell.fill = copy.copy(src_cell.fill)
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.protection = copy.copy(src_cell.protection)
                            tgt_cell.alignment = copy.copy(src_cell.alignment)
                    ws_gh.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
                    ws_gh.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
                    ws_gh.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
                    ws_gh.merge_cells(start_row=r, start_column=10, end_row=r, end_column=15)
                    ws_gh.merge_cells(start_row=r, start_column=16, end_row=r, end_column=19)
                cy_endrow_gh = 84
            else:
                cy_endrow_gh = 42

            for idx, m in enumerate(sorted_members_hs):
                if idx < 26:
                    curr_row = 17 + idx
                elif idx < 66:
                    curr_row = 45 + (idx - 26)
                else:
                    break

                cyxm = str(m.get('CYXM', '')) if pd.notna(m.get('CYXM')) else ''
                cyxb = xb_code_to_name(m.get('CYXB', ''))
                cyzjhm = str(m.get('CYZJHM', '')).strip() if pd.notna(m.get('CYZJHM')) else ''

                csrq = ''
                if len(cyzjhm) == 18 and cyzjhm[:14].isdigit():
                    birth_year = cyzjhm[6:10]
                    birth_month = cyzjhm[10:12]
                    birth_day = cyzjhm[12:14]
                    csrq = f"{birth_year}年{birth_month}月{birth_day}日"
                elif len(cyzjhm) == 15 and cyzjhm[:12].isdigit():
                    birth_year = "19" + cyzjhm[6:8]
                    birth_month = cyzjhm[8:10]
                    birth_day = cyzjhm[10:12]
                    csrq = f"{birth_year}年{birth_month}月{birth_day}日"

                yhzgx_raw = str(m.get('YHZGX', '')).strip() if pd.notna(m.get('YHZGX')) else ''
                yhzgx = yhzgx_code_to_name(yhzgx_raw) if yhzgx_raw.isdigit() else yhzgx_raw

                is_head = yhzgx in ('本人', '户主')
                m_lxdh = lxdh if is_head else '/'
                zjmc = '居民身份证' if is_head else '户口簿'

                ws_gh.cell(curr_row, 1).value = cyxm
                ws_gh.cell(curr_row, 2).value = cyxb
                ws_gh.cell(curr_row, 3).value = csrq
                ws_gh.cell(curr_row, 5).value = yhzgx
                ws_gh.cell(curr_row, 7).value = m_lxdh
                ws_gh.cell(curr_row, 8).value = zjmc
                ws_gh.cell(curr_row, 10).value = cyzjhm
                ws_gh.cell(curr_row, 16).value = '/'

            # 承包地块情况 (归户表)
            page1_start_gh = cy_endrow_gh + 1

            rel_merges_gh = []
            for mr in ws_gh.merged_cells.ranges:
                if page1_start_gh <= mr.min_row <= page1_start_gh + page_len - 1 and page1_start_gh <= mr.max_row <= page1_start_gh + page_len - 1:
                    rel_merges_gh.append((mr.min_col, mr.min_row - page1_start_gh, mr.max_col, mr.max_row - page1_start_gh))

            # 如果[CBDKZS]>8，增加整页“承包地块情况”行
            for ep in range(extra_pages):
                insert_pos = page1_start_gh + page_len * (ep + 1)
                insert_rows_with_merges(ws_gh, insert_pos, page_len)
                for i in range(page_len):
                    src_r = page1_start_gh + i
                    tgt_r = insert_pos + i
                    ws_gh.row_dimensions[tgt_r].height = ws_gh.row_dimensions[src_r].height
                    for c in range(1, 20):
                        src_cell = ws_gh.cell(src_r, c)
                        tgt_cell = ws_gh.cell(tgt_r, c)
                        if src_cell.has_style:
                            tgt_cell.font = copy.copy(src_cell.font)
                            tgt_cell.border = copy.copy(src_cell.border)
                            tgt_cell.fill = copy.copy(src_cell.fill)
                            tgt_cell.number_format = src_cell.number_format
                            tgt_cell.protection = copy.copy(src_cell.protection)
                            tgt_cell.alignment = copy.copy(src_cell.alignment)
                        if i in range(6) or src_cell.value in ['东：', '南：', '西：', '北：']:
                            tgt_cell.value = src_cell.value
                for min_col, min_r, max_col, max_r in rel_merges_gh:
                    ws_gh.merge_cells(start_row=insert_pos + min_r, start_column=min_col, end_row=insert_pos + max_r, end_column=max_col)

            # 地块汇总：B{cy_endrow_gh+3}：地块数量；H{cy_endrow_gh+3}：合同总面积亩；O{cy_endrow_gh+3}：实测总面积亩
            for p in range(total_pages):
                p_sum_row = page1_start_gh + p * page_len + 2
                ws_gh.cell(p_sum_row, 2).value = str(cbdkzs)
                ws_gh.cell(p_sum_row, 8).value = f"{float(htzmjm):.2f}亩"
                ws_gh.cell(p_sum_row, 15).value = f"{float(htzmjm):.2f}亩"

            # 填充地块槽位：每4行一个地块 (归户表)
            for idx, land in enumerate(cbf_lands.to_dict('records')):
                page_num = idx // 8
                slot_in_page = idx % 8
                r0 = page1_start_gh + page_num * page_len + 6 + slot_in_page * 4

                dkbm_val = str(land.get('DKBM', '')).strip()
                dkbm_last5 = dkbm_val[-5:] if len(dkbm_val) >= 5 else dkbm_val
                dkmc = str(land.get('DKMC', '')).strip() if pd.notna(land.get('DKMC')) else ''

                clean_east = clean_border_desc(land.get('DKDZ'), '东')
                clean_south = clean_border_desc(land.get('DKNZ'), '南')
                clean_west = clean_border_desc(land.get('DKXZ'), '西')
                clean_north = clean_border_desc(land.get('DKBZ'), '北')

                htmjm_val = land.get('HTMJM')
                htmjm_str = f"{float(htmjm_val):.2f}" if pd.notna(htmjm_val) and str(htmjm_val).strip() != '' else ''

                scmjm_val = land.get('SCMJM')
                scmjm_str = f"{float(scmjm_val):.2f}" if pd.notna(scmjm_val) and str(scmjm_val).strip() != '' else htmjm_str

                tdyt_str = tdyt_code_to_name(land.get('TDYT', ''))
                dklb_str = dklb_code_to_name(land.get('DKLB', ''))
                tdlylx_str = tdlylx_code_to_name(land.get('TDLYLX', ''))
                dldj_str = dldj_code_to_name(land.get('DLDJ', ''))
                sfjbnt_str = sfjbnt_code_to_name(land.get('SFJBNT', ''))

                ws_gh.cell(r0, 1).value = dkbm_last5
                ws_gh.cell(r0, 2).value = dkmc
                ws_gh.cell(r0, 4).value = f"东：{clean_east}" if clean_east else "东："
                ws_gh.cell(r0 + 1, 4).value = f"南：{clean_south}" if clean_south else "南："
                ws_gh.cell(r0 + 2, 4).value = f"西：{clean_west}" if clean_west else "西："
                ws_gh.cell(r0 + 3, 4).value = f"北：{clean_north}" if clean_north else "北："
                ws_gh.cell(r0, 8).value = htmjm_str
                ws_gh.cell(r0, 9).value = scmjm_str
                ws_gh.cell(r0, 10).value = tdyt_str
                ws_gh.cell(r0, 11).value = dklb_str
                ws_gh.cell(r0, 13).value = tdlylx_str
                ws_gh.cell(r0, 15).value = dldj_str
                ws_gh.cell(r0, 17).value = sfjbnt_str
                ws_gh.cell(r0, 18).value = '家庭承包'

            # 填充“非承包地块情况”页的公示记事
            jishi_row_gh = None
            for r in range(page1_start_gh + total_pages * page_len, ws_gh.max_row + 1):
                val_1 = ws_gh.cell(r, 1).value
                if val_1 and '记事' in str(val_1):
                    jishi_row_gh = r
                    break

            if jishi_row_gh is not None:
                ws_gh.cell(jishi_row_gh, 2).value = cbfdcjs_hs
                # 公示记事时间->公示记事栏的N列；承包方(代表)时间->承包方(代表)对公示结果的意见 栏的N列；审核意见时间->公示结果审核意见 栏的N列
                if time_info:
                    dt_gh_gs = format_date_str(time_info.get('gh_gs_dt'))
                    dt_gh_db = format_date_str(time_info.get('gh_db_dt'))
                    dt_gh_sh = format_date_str(time_info.get('gh_sh_dt'))
                    if dt_gh_gs:
                        ws_gh.cell(jishi_row_gh + 4, 14).value = dt_gh_gs
                    if dt_gh_db:
                        ws_gh.cell(jishi_row_gh + 9, 14).value = dt_gh_db
                    if dt_gh_sh:
                        ws_gh.cell(jishi_row_gh + 14, 14).value = dt_gh_sh

            # 导出文件名：{CBFBM}_{CBFMC}_承包经营权公示结果归户表.xlsx
            gh_filename = f"{cbfbm}_{cbfmc}_承包经营权公示结果归户表.xlsx"
            gh_output_path = os.path.join(cbf_dir, gh_filename)
            wb_gh.save(gh_output_path)

            # 解析发包方名称各层级与户主性别
            zhen_clean, cun_clean, zu_clean = parse_yushe_parts(fbfmc)
            head_xb = sorted_members_hs[0].get('CYXB', '') if sorted_members_hs else ''
            xb_str = parse_gender_from_id(cbfzjhm, fallback=xb_code_to_name(head_xb) or '男')

            # -----------------------------
            # 5. 生成公示无异议声明书
            # -----------------------------
            if os.path.exists(sm_tpl_path) and Document is not None:
                doc_sm = Document(sm_tpl_path)
                replace_bookmark_in_doc(doc_sm, 'CBFMC', cbfmc)
                replace_bookmark_in_doc(doc_sm, 'CBFZJHM', cbfzjhm)
                replace_bookmark_in_doc(doc_sm, 'CJYQMC', cun_clean)
                replace_bookmark_in_doc(doc_sm, 'CJQYMC', cun_clean)
                replace_bookmark_in_doc(doc_sm, 'CJQYMC1', cun_clean)
                replace_bookmark_in_doc(doc_sm, 'XJQYMC', zhen_clean)
                replace_bookmark_in_doc(doc_sm, 'XJQYMC1', zhen_clean)
                replace_bookmark_in_doc(doc_sm, 'XB', xb_str)

                if zu_clean:
                    replace_bookmark_in_doc(doc_sm, 'ZJQYMC', zu_clean)
                    replace_bookmark_in_doc(doc_sm, 'ZJQYMC1', zu_clean)
                else:
                    replace_bookmark_in_doc(doc_sm, 'ZJQYMC', '')
                    replace_bookmark_in_doc(doc_sm, 'ZJQYMC1', '')
                    for p in doc_sm.paragraphs:
                        for i, r in enumerate(p.runs):
                            if '小组村民' in r.text or '组村民' in r.text:
                                r.text = r.text.replace('小组村民', '村民').replace('组村民', '村民')
                                if i >= 1 and p.runs[i - 1].text == ' ':
                                    p.runs[i - 1].text = ''
                                if i >= 3 and p.runs[i - 3].text == ' ':
                                    p.runs[i - 3].text = ''
                            if '组集体经济组织' in r.text:
                                r.text = r.text.replace(' 组集体经济组织', '集体经济组织').replace('组集体经济组织', '集体经济组织')
                                if i >= 1 and p.runs[i - 1].text == ' ':
                                    p.runs[i - 1].text = ''
                                if i >= 3 and p.runs[i - 3].text == ' ':
                                    p.runs[i - 3].text = ''

                # 填写公示无异议声明书时间（替换yyyy->年份，mm->月份，dd->天，RQ->XXXX年XX月XX日）
                if time_info and time_info.get('sm_dt'):
                    sm_dt = time_info['sm_dt']
                    replace_bookmark_in_doc(doc_sm, 'yyyy', str(sm_dt.year))
                    replace_bookmark_in_doc(doc_sm, 'mm', str(sm_dt.month))
                    replace_bookmark_in_doc(doc_sm, 'dd', str(sm_dt.day))
                    replace_bookmark_in_doc(doc_sm, 'RQ', format_date_str(sm_dt))

                sm_filename = f"{cbfbm}_{cbfmc}_公示无异议声明书.docx"
                doc_sm.save(os.path.join(cbf_dir, sm_filename))

            # -----------------------------
            # 6. 生成户主声明书
            # -----------------------------
            if os.path.exists(hz_tpl_path) and Document is not None:
                doc_hz = Document(hz_tpl_path)
                replace_bookmark_in_doc(doc_hz, 'CBFMC', cbfmc)
                replace_bookmark_in_doc(doc_hz, 'CBFZJHM', cbfzjhm)
                replace_bookmark_in_doc(doc_hz, 'CBFCYSL', str(cbfcysl))
                cyxms = '、'.join([str(m['CYXM']) for m in sorted_members_hs if pd.notna(m.get('CYXM')) and str(m.get('CYXM')).strip()])
                replace_bookmark_in_doc(doc_hz, 'CYXMs', cyxms)
                replace_bookmark_in_doc(doc_hz, 'FBFMC', fbfmc)
                replace_bookmark_in_doc(doc_hz, 'XB', xb_str)

                # 填写户主声明书时间（替换书签 RQ->XXXX年XX月XX日）
                if time_info and time_info.get('hz_rq_dt'):
                    replace_bookmark_in_doc(doc_hz, 'RQ', format_date_str(time_info['hz_rq_dt']))

                hz_filename = f"{cbfbm}_{cbfmc}_户主声明书.docx"
                doc_hz.save(os.path.join(cbf_dir, hz_filename))

            # -----------------------------
            # 7. 生成家庭承包申请书
            # -----------------------------
            if os.path.exists(sq_tpl_path) and Document is not None:
                doc_sq = Document(sq_tpl_path)
                replace_bookmark_in_doc(doc_sq, 'CBDKZS', str(cbdkzs))
                replace_bookmark_in_doc(doc_sq, 'CBFMC', cbfmc)
                replace_bookmark_in_doc(doc_sq, 'CBFZJHM', cbfzjhm)
                replace_bookmark_in_doc(doc_sq, 'CJQYMC', cun_clean)
                htzmjm_fmt = f"{float(htzmjm):.2f}"
                replace_bookmark_in_doc(doc_sq, 'HTZMJM', htzmjm_fmt)
                # 计算实测面积汇总及各土地利用类型面积
                # sczmjm->汇总地块SCMJM，stmj->汇总tdlylx为水田的地块的scmjm，sjdmj->汇总tdlylx为水浇地的地块的scmjm，hdmj->sczmjm-stmj-sjdmj
                sczmjm_val = 0.0
                stmj_val = 0.0
                sjdmj_val = 0.0
                if not cbf_lands.empty:
                    for _, land_r in cbf_lands.iterrows():
                        sc_val = land_r.get('SCMJM')
                        if pd.isna(sc_val) or str(sc_val).strip() == '':
                            sc_val = land_r.get('HTMJM', 0.0)
                        try:
                            sc_float = float(sc_val)
                        except (ValueError, TypeError):
                            sc_float = 0.0
                        sczmjm_val += sc_float

                        tdly_code = str(land_r.get('TDLYLX', '')).strip()
                        tdly_name = tdlylx_code_to_name(tdly_code)
                        if tdly_code in ('011', '11') or tdly_name == '水田':
                            stmj_val += sc_float
                        elif tdly_code in ('012', '12') or tdly_name == '水浇地':
                            sjdmj_val += sc_float

                hdmj_val = sczmjm_val - stmj_val - sjdmj_val

                replace_bookmark_in_doc(doc_sq, 'SCZMJM', f"{sczmjm_val:.2f}")
                replace_bookmark_in_doc(doc_sq, 'STMJ', f"{stmj_val:.2f}")
                replace_bookmark_in_doc(doc_sq, 'SJDMJ', f"{sjdmj_val:.2f}")
                replace_bookmark_in_doc(doc_sq, 'HDMJ', f"{hdmj_val:.2f}")
                replace_bookmark_in_doc(doc_sq, 'XJQYMC', zhen_clean)

                if zu_clean:
                    replace_bookmark_in_doc(doc_sq, 'ZJQYMC', ' ' + zu_clean)
                else:
                    replace_bookmark_in_doc(doc_sq, 'ZJQYMC', '')
                    for p in doc_sq.paragraphs:
                        for r in p.runs:
                            if '组农村耕地' in r.text:
                                r.text = r.text.replace('组农村耕地', '农村耕地')

                # 替换承包期限（使用书签替换：Qyyyy, Qmm, Qdd, Zyyyy, Zmm, Zdd）
                q_y = ht_info.get('Q_Y', '2024')
                q_m = ht_info.get('Q_M', '01')
                q_d = ht_info.get('Q_D', '01')
                z_y = ht_info.get('Z_Y', '2053')
                z_m = ht_info.get('Z_M', '12')
                z_d = ht_info.get('Z_D', '31')

                replace_bookmark_in_doc(doc_sq, 'Qyyyy', q_y)
                replace_bookmark_in_doc(doc_sq, 'Qmm', q_m)
                replace_bookmark_in_doc(doc_sq, 'Qdd', q_d)
                replace_bookmark_in_doc(doc_sq, 'Zyyyy', z_y)
                replace_bookmark_in_doc(doc_sq, 'Zmm', z_m)
                replace_bookmark_in_doc(doc_sq, 'Zdd', z_d)

                # 填写家庭承包申请书时间（替换书签 RQ->XXXX年XX月XX日）
                if time_info and time_info.get('sq_dt'):
                    replace_bookmark_in_doc(doc_sq, 'RQ', format_date_str(time_info['sq_dt']))

                sq_filename = f"{cbfbm}_{cbfmc}_家庭承包申请书.docx"
                doc_sq.save(os.path.join(cbf_dir, sq_filename))

            # -----------------------------
            # 8. 生成承包地块调查表（每个地块一个独立Excel文件）
            # -----------------------------
            if os.path.exists(dk_tpl_path) and not cbf_lands.empty:
                def set_check_dk(text, target_label, all_labels):
                    if not text:
                        return text
                    res = str(text)
                    for lbl in all_labels:
                        if lbl == target_label:
                            # 选中项：将未选中 £（或旧字符）替换为 R
                            res = res.replace(f'£{lbl}', f'R{lbl}').replace(f'□{lbl}', f'R{lbl}').replace(f'þ{lbl}', f'R{lbl}')
                        else:
                            # 未选中项：将 R（或旧字符）替换为未选中 £
                            res = res.replace(f'R{lbl}', f'£{lbl}').replace(f'þ{lbl}', f'£{lbl}').replace(f'□{lbl}', f'£{lbl}')
                    return res

                # 建立本批 JZD 字典（记录 JZDH、BSM、JBLX 等）
                jzd_dict = {}
                if not jzd_df.empty and 'JZDH' in jzd_df.columns:
                    for _, z in jzd_df.iterrows():
                        jzd_dict[str(z['JZDH'])] = {
                            'JZDH': str(z.get('JZDH', '')),
                            'BSM': str(z.get('BSM', '')),
                            'JBLX': str(z.get('JBLX', '')).strip(),
                            'JZDLX': str(z.get('JZDLX', '')).strip(),
                        }

                # 代码直接映射列号（严格按照字典值匹配对应表头列位）
                JBLX_CODE_TO_COL = {
                    '6': 5,   # 无标志 (E)
                    '7': 3,   # 木桩 (C)
                    '8': 4,   # 埋石 (D)
                }
                JZXLB_CODE_TO_COL = {
                    '01': 7,  # 田埂 (G)
                    '1': 7,
                    '02': 9,  # 沟渠 (I)
                    '2': 9,
                    '03': 10, # 道路 (J)
                    '3': 10,
                    '04': 11, # 行树 (K)
                    '4': 11,
                    '05': 12, # 围墙 (L)
                    '5': 12,
                    '06': 13, # 墙壁 (M)
                    '6': 13,
                    '07': 14, # 栅栏 (N)
                    '7': 14,
                    '08': 15, # 两点连线 (O)
                    '8': 15,
                }
                JZXWZ_CODE_TO_COL = {
                    '01': 19, # 内 (S)
                    '1': 19,
                    '02': 20, # 中 (T)
                    '2': 20,
                    '03': 21, # 外 (U)
                    '3': 21,
                }

                CHECK_CHAR = '\u00fc'  # 'ü'，在 Wingdings 字体下显示为对勾符号 (Unicode U+00FC)
                wingdings_font = openpyxl.styles.Font(name='Wingdings', size=9.0)

                def get_pt_info(pt_id):
                    info = jzd_dict.get(str(pt_id), {})
                    if isinstance(info, dict):
                        jzdh_val = info.get('JZDH') or str(pt_id)
                        jblx_code = str(info.get('JBLX', '6')).strip()
                    else:
                        jzdh_val = str(info) or str(pt_id)
                        jblx_code = '6'
                    jblx_col = JBLX_CODE_TO_COL.get(jblx_code, 5)
                    return jzdh_val, jblx_col

                def format_jzxsm_continuous(orig_sm, start_num, end_num):
                    """
                    规范化界址线说明，保证起点和终点为连续的 T1-T2, T2-T3...Tn-T1，
                    避免共享界址线或原始数据中保留邻宗地文本导致的编号跳跃或重复。
                    """
                    orig = str(orig_sm).strip() if orig_sm else ''
                    if not orig:
                        return f'由T{start_num}至T{end_num}'
                    if '沿' in orig:
                        m = re.search(r'(沿.*?)(?:[至到](?:T\d+|\d+)?|$)', orig)
                        if m:
                            middle = m.group(1).rstrip('至').rstrip('到')
                            return f'由T{start_num}{middle}至T{end_num}'
                        return f'由T{start_num}至T{end_num}'
                    return f'由T{start_num}{orig}至T{end_num}'

                for _, land_item in cbf_lands.iterrows():
                    dkbm_item = str(land_item.get('DKBM', '')).strip()
                    if not dkbm_item:
                        continue
                    dkmc_item = str(land_item.get('DKMC', '')).strip() if pd.notna(land_item.get('DKMC')) else ''
                    dkbm_last5 = dkbm_item[-5:] if len(dkbm_item) >= 5 else dkbm_item

                    htmjm_val = land_item.get('HTMJM')
                    htmjm_str = f"{float(htmjm_val):.2f}" if pd.notna(htmjm_val) and str(htmjm_val).strip() != '' else ''
                    scmjm_val = land_item.get('SCMJM')
                    scmjm_str = f"{float(scmjm_val):.2f}" if pd.notna(scmjm_val) and str(scmjm_val).strip() != '' else htmjm_str

                    # 载入模板
                    wb_dk = openpyxl.load_workbook(dk_tpl_path)
                    ws_dk = wb_dk.active

                    # 基本信息
                    ws_dk['D1'].value = fbfbm
                    ws_dk['O1'].value = cbfmc
                    tfh_raw = land_item.get('TFH')
                    if tfh_raw is None or pd.isna(tfh_raw) or str(tfh_raw).strip() in ('', 'nan', 'None'):
                        # 矢量地块数据中不存在TFH或为空，自动按NY/T 2537计算该地块所在1:2000正方形分幅图幅号
                        geom = vector_geom_by_dkbm.get(dkbm_item)
                        tfh_text = compute_1_2000_tfh_square(geom, vector_crs) if geom is not None else ''
                    else:
                        tfh_text = str(tfh_raw).strip()
                    ws_dk['V1'].value = tfh_text
                    ws_dk['C5'].value = dkbm_last5
                    ws_dk['M5'].value = dkmc_item
                    ws_dk['W5'].value = htmjm_str
                    ws_dk['AD5'].value = scmjm_str

                    # 四至
                    clean_east = clean_border_desc(land_item.get('DKDZ'), '东')
                    clean_south = clean_border_desc(land_item.get('DKNZ'), '南')
                    clean_west = clean_border_desc(land_item.get('DKXZ'), '西')
                    clean_north = clean_border_desc(land_item.get('DKBZ'), '北')
                    ws_dk['B6'].value = clean_east or (str(land_item.get('DKDZ', '')).strip() if pd.notna(land_item.get('DKDZ')) else '')
                    ws_dk['J6'].value = clean_south or (str(land_item.get('DKNZ', '')).strip() if pd.notna(land_item.get('DKNZ')) else '')
                    ws_dk['S6'].value = clean_west or (str(land_item.get('DKXZ', '')).strip() if pd.notna(land_item.get('DKXZ')) else '')
                    ws_dk['Z6'].value = clean_north or (str(land_item.get('DKBZ', '')).strip() if pd.notna(land_item.get('DKBZ')) else '')

                    # 复选框匹配
                    syqxz_raw = str(land_item.get('SYQXZ', '')).strip()
                    syqxz_target = SYQXZ_MAP.get(syqxz_raw, syqxz_code_to_name(syqxz_raw))
                    if not syqxz_target:
                        syqxz_target = '村民小组'

                    ws_dk['F2'].value = set_check_dk(ws_dk['F2'].value, syqxz_target, ['国有土地所有权', '集体土地所有权', '村民小组', '村级集体经济组织'])
                    ws_dk['F3'].value = set_check_dk(ws_dk['F3'].value, syqxz_target, ['乡级集体经济组织', '其他农民集体经济组织'])
                    ws_dk['F4'].value = set_check_dk(ws_dk['F4'].value, '家庭承包', ['家庭承包', '招标', '公开协商', '拍卖', '转让', '互换', '其他'])

                    tdyt_name = tdyt_code_to_name(land_item.get('TDYT', '')) or '种植业'
                    tdyt_labels = ['种植业', '林业', '畜牧业', '渔业', '其他']
                    for r_chk, lbl in [(7, '种植业'), (8, '林业'), (9, '畜牧业'), (10, '渔业'), (11, '其他')]:
                        ws_dk.cell(r_chk, 2).value = set_check_dk(ws_dk.cell(r_chk, 2).value, tdyt_name, tdyt_labels)

                    dklb_name = dklb_code_to_name(land_item.get('DKLB', '')) or '承包地块'
                    dklb_labels = ['承包地块', '自留地', '机动地', '开荒地', '其他集体土地']
                    for r_chk, lbl in [(7, '承包地块'), (8, '自留地'), (9, '机动地'), (10, '开荒地'), (11, '其他集体土地')]:
                        ws_dk.cell(r_chk, 8).value = set_check_dk(ws_dk.cell(r_chk, 8).value, dklb_name, dklb_labels)

                    dldj_name = dldj_code_to_name(land_item.get('DLDJ', ''))
                    # 确保 label 格式为 "X等地"
                    if dldj_name and not dldj_name.endswith('地'):
                        dldj_target = f"{dldj_name}地"
                    else:
                        dldj_target = dldj_name
                    for r_chk, pair in [(7, ['一等地', '二等地']), (8, ['三等地', '四等地']), (9, ['五等地', '六等地']), (10, ['七等地', '八等地']), (11, ['九等地', '十等地'])]:
                        ws_dk.cell(r_chk, 15).value = set_check_dk(ws_dk.cell(r_chk, 15).value, dldj_target, pair)

                    tdlylx_raw_name = tdlylx_code_to_name(land_item.get('TDLYLX', ''))
                    tdlylx_labels = ['水田', '旱地', '水浇地', '果园', '茶园', '其他']
                    tdlylx_name = tdlylx_raw_name if tdlylx_raw_name in tdlylx_labels else '其他'
                    for r_chk, lbl in [(7, '水田'), (8, '旱地'), (9, '水浇地'), (10, '果园'), (11, '茶园'), (12, '其他')]:
                        ws_dk.cell(r_chk, 24).value = set_check_dk(ws_dk.cell(r_chk, 24).value, tdlylx_name, tdlylx_labels)

                    sfjbnt_name = sfjbnt_code_to_name(land_item.get('SFJBNT', '')) or '否'
                    ws_dk['AB10'].value = set_check_dk(ws_dk['AB10'].value, sfjbnt_name, ['是', '否'])

                    # 提取界址线并进行拓扑闭环排序与点号解算
                    lines = []
                    ordered_pts = []
                    if not jzx_df.empty and 'DKBM' in jzx_df.columns:
                        jzx_sub = jzx_df[jzx_df['DKBM'].astype(str).str.contains(dkbm_item, na=False)].copy()
                        if not jzx_sub.empty:
                            lines, ordered_pts = chain_parcel_boundary(jzx_sub, dkbm_item)

                    num_lines = len(lines)

                    line_cols = [
                        (7, 7), (8, 8), (9, 9), (10, 10), (11, 11), (12, 12), (13, 13), (14, 14),
                        (15, 16), (17, 18), (19, 19), (20, 20), (21, 21), (22, 25), (26, 27), (28, 29)
                    ]

                    title_merges = [
                        (1, 0, 1, 2), (2, 0, 2, 2), (3, 0, 6, 0), (3, 1, 3, 2), (4, 1, 4, 2), (5, 1, 6, 2),
                        (7, 0, 18, 0), (7, 1, 7, 2), (8, 1, 8, 2), (9, 1, 9, 2), (10, 1, 10, 2), (11, 1, 11, 2),
                        (12, 1, 12, 2), (13, 1, 13, 2), (14, 1, 14, 2), (15, 1, 16, 2), (17, 1, 18, 2),
                        (19, 0, 21, 0), (19, 1, 19, 2), (20, 1, 20, 2), (21, 1, 21, 2), (22, 0, 25, 2),
                        (26, 0, 31, 0), (26, 1, 29, 1), (26, 2, 27, 2), (28, 2, 29, 2), (30, 1, 31, 2)
                    ]

                    thin_border_side = openpyxl.styles.Side(style="thin", color="000000")

                    def set_last_row_style(ws, last_row):
                        # 补全最后一行水平合并 (O:P, Q:R, V:Y, Z:AA, AB:AC)
                        for min_c, max_c in [(15, 16), (17, 18), (22, 25), (26, 27), (28, 29)]:
                            ws.merge_cells(start_row=last_row, start_column=min_c, end_row=last_row, end_column=max_c)
                        # 补齐整行 A~AE (1~31 列) 的边框，特别是 bottom=thin 下边框
                        for c in range(1, 32):
                            cell = ws.cell(last_row, c)
                            cur_b = cell.border
                            l = cur_b.left.style if cur_b and cur_b.left else "thin"
                            r = cur_b.right.style if cur_b and cur_b.right else "thin"
                            t = cur_b.top.style if cur_b and cur_b.top else "thin"
                            cell.border = openpyxl.styles.Border(
                                left=openpyxl.styles.Side(style=l, color="000000") if l else thin_border_side,
                                right=openpyxl.styles.Side(style=r, color="000000") if r else thin_border_side,
                                top=openpyxl.styles.Side(style=t, color="000000") if t else thin_border_side,
                                bottom=thin_border_side
                            )

                    # 分页与行布局处理：
                    # 第一页界址点区域严格使用模版默认行（16~40行），最多容纳 12 条界址线，
                    # 最后一个界址点（第13个点槽位，单页闭合点或跨页第13个点）最大放置在第 39~40 行的合并单元格中，不向下插入额外行；
                    # 若界址线 > 12 条，则第一页取前 12 条线后在第 40 行处分页，从第 41 行起生成第二页及后续页（每页容纳 17 条线，共 38 行）；
                    # 签名栏随之平移到最后一页底部。
                    if num_lines <= 12:
                        # -------------------------
                        # 单页情况（界址线 <= 12条，点数 <= 13）
                        # -------------------------
                        set_last_row_style(ws_dk, 40)

                        # 单页填充界址点与界址线
                        first_pt_id = ''
                        first_pt_pre = ''
                        first_jblx_col = 5
                        for k, line in enumerate(lines):
                            pt_row = 16 if k == 0 else (15 + k * 2)
                            line_row = 16 + k * 2

                            pt_id = ordered_pts[k] if k < len(ordered_pts) else str(line.get('QJZDH', '')).strip()
                            jzdh_val, jblx_col = get_pt_info(pt_id)
                            pre_num = f'J{k + 1}'
                            if k == 0:
                                first_pt_id = jzdh_val
                                first_pt_pre = pre_num
                                first_jblx_col = jblx_col

                            ws_dk.cell(pt_row, 1).value = jzdh_val
                            ws_dk.cell(pt_row, 2).value = pre_num
                            c_jblx = ws_dk.cell(pt_row, jblx_col)
                            c_jblx.value = CHECK_CHAR
                            c_jblx.font = wingdings_font

                            jzxlb_code = str(line.get('JZXLB', '01')).strip()
                            jzxlb_col = JZXLB_CODE_TO_COL.get(jzxlb_code, 7)
                            start_t = k + 1
                            end_t = 1 if (k + 1 == num_lines) else (k + 2)
                            jzxsm_text = format_jzxsm_continuous(line.get('JZXSM', ''), start_t, end_t)

                            c_jzxlb = ws_dk.cell(line_row, jzxlb_col)
                            c_jzxlb.value = CHECK_CHAR
                            c_jzxlb.font = wingdings_font

                            jzxwz_raw = str(line.get('JZXWZ', '02')).strip()
                            jzxwz_col = JZXWZ_CODE_TO_COL.get(jzxwz_raw, 20)

                            c_jzxwz = ws_dk.cell(line_row, jzxwz_col)
                            c_jzxwz.value = CHECK_CHAR
                            c_jzxwz.font = wingdings_font

                            ws_dk.cell(line_row, 22).value = jzxsm_text
                            ws_dk.cell(line_row, 26).value = line.get('PLDWQLR', cbfmc)

                        if lines:
                            close_pt_row = 15 + num_lines * 2
                            ws_dk.cell(close_pt_row, 1).value = first_pt_id
                            ws_dk.cell(close_pt_row, 2).value = first_pt_pre
                            c_close = ws_dk.cell(close_pt_row, first_jblx_col)
                            c_close.value = CHECK_CHAR
                            c_close.font = wingdings_font

                    else:
                        # -------------------------
                        # 多页情况（界址线 > 12条）
                        # 第一页固定放置 12 条线，第 13 个点最大放置在第 39~40 行合并单元格
                        # -------------------------
                        pages_lines = [lines[:12]]
                        rem_lines = lines[12:]
                        while rem_lines:
                            pages_lines.append(rem_lines[:17])
                            rem_lines = rem_lines[17:]

                        set_last_row_style(ws_dk, 40)
                        # 在第一页末尾（第 40 行下方，即第 41 行前）插入分页符
                        ws_dk.row_breaks.append(Break(id=40))

                        insert_pos_cur = 41
                        for page_idx in range(1, len(pages_lines)):
                            # 每一页固定插入满 18 行（18个界址点位置/17条线，35行点线槽位 + 3行标题 = 38行）
                            p_total_rows = 3 + 1 + 17 * 2  # 固定 38 行

                            insert_rows_with_merges(ws_dk, insert_pos_cur, p_total_rows)

                            # 复制 3 行标题行（源自 13~15）
                            for i in range(3):
                                src_r = 13 + i
                                tgt_r = insert_pos_cur + i
                                ws_dk.row_dimensions[tgt_r].height = ws_dk.row_dimensions[src_r].height
                                for c in range(1, 32):
                                    src_cell = ws_dk.cell(src_r, c)
                                    tgt_cell = ws_dk.cell(tgt_r, c)
                                    tgt_cell.value = src_cell.value
                                    if src_cell.has_style:
                                        tgt_cell.font = copy.copy(src_cell.font)
                                        tgt_cell.border = copy.copy(src_cell.border)
                                        tgt_cell.fill = copy.copy(src_cell.fill)
                                        tgt_cell.alignment = copy.copy(src_cell.alignment)

                            for min_c, min_r, max_c, max_r in title_merges:
                                ws_dk.merge_cells(start_row=insert_pos_cur + min_r, start_column=min_c, end_row=insert_pos_cur + max_r, end_column=max_c)

                            pt_start_row = insert_pos_cur + 3
                            pt_end_row = insert_pos_cur + p_total_rows - 1
                            for r_ins in range(pt_start_row, pt_end_row + 1):
                                ws_dk.row_dimensions[r_ins].height = ws_dk.row_dimensions[16].height
                                src_r = 16 if ((r_ins - pt_start_row) % 2 == 0) else 17
                                for c in range(1, 32):
                                    src_cell = ws_dk.cell(src_r, c)
                                    tgt_cell = ws_dk.cell(r_ins, c)
                                    if src_cell.has_style:
                                        tgt_cell.font = copy.copy(src_cell.font)
                                        tgt_cell.border = copy.copy(src_cell.border)
                                        tgt_cell.fill = copy.copy(src_cell.fill)
                                        tgt_cell.alignment = copy.copy(src_cell.alignment)

                            ws_dk.merge_cells(start_row=pt_start_row, start_column=30, end_row=pt_end_row, end_column=31)
                            ws_dk.merge_cells(start_row=pt_start_row, start_column=5, end_row=pt_start_row, end_column=6)

                            for k in range(17):
                                r_line = pt_start_row + k * 2
                                for min_c, max_c in line_cols:
                                    ws_dk.merge_cells(start_row=r_line, start_column=min_c, end_row=r_line + 1, end_column=max_c)
                                r_pt = (pt_start_row - 1) + (k + 1) * 2
                                for col in [1, 2, 3, 4]:
                                    ws_dk.merge_cells(start_row=r_pt, start_column=col, end_row=r_pt + 1, end_column=col)
                                ws_dk.merge_cells(start_row=r_pt, start_column=5, end_row=r_pt + 1, end_column=6)

                            set_last_row_style(ws_dk, pt_end_row)

                            # 若还有下一页，在该页末尾插入分页符
                            if page_idx < len(pages_lines) - 1:
                                ws_dk.row_breaks.append(Break(id=pt_end_row))

                            insert_pos_cur += p_total_rows

                        # 跨页填充数据
                        first_pt_id = ''
                        first_pt_pre = ''
                        first_jblx_col = 5
                        cur_row_base = 16
                        global_line_idx = 0

                        for page_idx, p_lines in enumerate(pages_lines):
                            cur_pt_start = 16 if page_idx == 0 else (cur_row_base + 3)

                            # 填充该页顶部的起始界址点
                            top_pt_id = ordered_pts[global_line_idx] if global_line_idx < len(ordered_pts) else str(p_lines[0].get('QJZDH', '')).strip()
                            top_jzdh, top_jblx_col = get_pt_info(top_pt_id)
                            top_pre_num = f'J{global_line_idx + 1}'

                            if page_idx == 0:
                                first_pt_id = top_jzdh
                                first_pt_pre = top_pre_num
                                first_jblx_col = top_jblx_col

                            ws_dk.cell(cur_pt_start, 1).value = top_jzdh
                            ws_dk.cell(cur_pt_start, 2).value = top_pre_num
                            c_jblx = ws_dk.cell(cur_pt_start, top_jblx_col)
                            c_jblx.value = CHECK_CHAR
                            c_jblx.font = wingdings_font

                            # 依次填充界址线及后续界址点
                            for k, line in enumerate(p_lines):
                                line_row = cur_pt_start + k * 2
                                g_idx = global_line_idx + k

                                jzxlb_code = str(line.get('JZXLB', '01')).strip()
                                jzxlb_col = JZXLB_CODE_TO_COL.get(jzxlb_code, 7)
                                start_t = g_idx + 1
                                end_t = 1 if (g_idx + 1 == num_lines) else (g_idx + 2)
                                jzxsm_text = format_jzxsm_continuous(line.get('JZXSM', ''), start_t, end_t)

                                c_jzxlb = ws_dk.cell(line_row, jzxlb_col)
                                c_jzxlb.value = CHECK_CHAR
                                c_jzxlb.font = wingdings_font

                                jzxwz_raw = str(line.get('JZXWZ', '02')).strip()
                                jzxwz_col = JZXWZ_CODE_TO_COL.get(jzxwz_raw, 20)
                                c_jzxwz = ws_dk.cell(line_row, jzxwz_col)
                                c_jzxwz.value = CHECK_CHAR
                                c_jzxwz.font = wingdings_font

                                ws_dk.cell(line_row, 22).value = jzxsm_text
                                ws_dk.cell(line_row, 26).value = line.get('PLDWQLR', cbfmc)

                                # 线的下一端点
                                next_pt_row = (cur_pt_start - 1) + (k + 1) * 2
                                next_g_idx = g_idx + 1
                                if next_g_idx == num_lines:
                                    # 闭合回起始点
                                    next_jzdh = first_pt_id
                                    next_pre = first_pt_pre
                                    next_col = first_jblx_col
                                else:
                                    next_pt_id = ordered_pts[next_g_idx] if next_g_idx < len(ordered_pts) else str(line.get('ZJZDH', '')).strip()
                                    next_jzdh, next_col = get_pt_info(next_pt_id)
                                    next_pre = f'J{next_g_idx + 1}'

                                ws_dk.cell(next_pt_row, 1).value = next_jzdh
                                ws_dk.cell(next_pt_row, 2).value = next_pre
                                c_next = ws_dk.cell(next_pt_row, next_col)
                                c_next.value = CHECK_CHAR
                                c_next.font = wingdings_font

                            global_line_idx += len(p_lines)
                            if page_idx == 0:
                                cur_row_base = 41
                            else:
                                cur_row_base += 38

                    # 填写地块调查表时间（承包方（代表）时间->承包方（代表）栏的H列；调查记事时间->调查记事栏的Y列；审核意见时间->审核意见栏的Y列）
                    if time_info:
                        sign_r = None
                        audit_r = None
                        for r_chk in range(16, ws_dk.max_row + 1):
                            v_c3 = str(ws_dk.cell(r_chk, 3).value or '')
                            if '签字' in v_c3:
                                sign_r = r_chk
                            elif '审核人' in v_c3:
                                audit_r = r_chk

                        dt_dk_db = format_date_str(time_info.get('dk_db_dt'))
                        dt_dk_dc = format_date_str(time_info.get('dk_dc_dt'))
                        dt_dk_sh = format_date_str(time_info.get('dk_sh_dt'))

                        if sign_r is not None:
                            if dt_dk_db:
                                ws_dk.cell(sign_r, 8).value = dt_dk_db
                            if dt_dk_dc:
                                ws_dk.cell(sign_r, 25).value = dt_dk_dc
                        if audit_r is not None and dt_dk_sh:
                            ws_dk.cell(audit_r, 25).value = dt_dk_sh

                    # 保存地块调查表
                    dk_filename = f"{cbfbm}_{cbfmc}_地块调查表{dkbm_item}.xlsx"
                    ws_dk.title = "承包地块调查表"
                    wb_dk.save(os.path.join(cbf_dir, dk_filename))

        logger.info(f"榆社县一户一档材料生成完成！共处理 {len(fbf_df)} 个发包方，{len(cbf_df)} 户承包方。")

    finally:
        conn.close()


# 示例调用
if __name__ == "__main__":
    print("选择功能：")
    print("1. 生成公示表")
    print("2. 生成摸底调查表（openpyxl版本）")
    print("3. 生成金寨县摸底调查表")
    print("4. 生成承包地块调查表")
    print("5. 生成户主代表声明书")
    print("6. 生成金寨县公示表")
    print("7. 生成榆社县摸底调查表")
    print("8. 生成榆社县一户一档材料")

    choice = input("请输入选择 (1/2/3/4/5/6/7/8): ")

    if choice == "1":
        print("\n=== 生成公示表 ===")
        jtjjzz_name = input("请输入集体经济组织名称: ")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")
        generate_gsb(mdb_file_path, shp_file_path, output_directory, jtjjzz_name)

    elif choice == "2":
        print("\n=== 生成摸底调查表（openpyxl版本） ===")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")

        # 创建临时文件夹处理数据
        temp_dir = tempfile.mkdtemp()

        # 创建MDB连接并处理数据，得到family_info和land_info
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'  
            r'DBQ=' + mdb_file_path + ';'  
        )
        conn = pyodbc.connect(conn_str)

        try:
            # 读取并处理数据，按照公示表的逻辑
            tables = ['CBDKXX', 'CBF', 'CBF_JTCY', 'FBF', 'CBHT']
            # 抑制pandas的SQLAlchemy连接警告
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                dfs = {table: pd.read_sql_query(f'SELECT * FROM [{table}]', conn) for table in tables}

            # 数据处理逻辑
            cbf_jtcy = dfs['CBF_JTCY']
            cbf = dfs['CBF']
            cbf['GSSHRQ'] = pd.to_datetime(cbf['GSSHRQ'], errors='coerce')
            cbf['CBFDCRQ'] = pd.to_datetime(cbf['CBFDCRQ'], errors='coerce')
            cbht = dfs['CBHT']
            cbht['CBQXQ'] = pd.to_datetime(cbht['CBQXQ'], errors='coerce')
            cbht['CBQXZ'] = pd.to_datetime(cbht['CBQXZ'], errors='coerce')
            cbht_selected = cbht[['CBFBM', 'CBQXQ', 'CBQXZ']]

            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC','CBFBM','LXDH','CBFDZ']], on='CBFBM', how='left')
            merged_cbf_jtcy_cbf = pd.merge(merged_cbf_jtcy_cbf, cbht_selected, on='CBFBM', how='left')
            merged_cbf_jtcy_cbf['CBFBM_14'] = merged_cbf_jtcy_cbf['CBFBM'].str[:14]
            fbf = dfs['FBF']
            merged_all = pd.merge(merged_cbf_jtcy_cbf, fbf, left_on='CBFBM_14', right_on='FBFBM', how='left')

            # 处理Shp文件得到vector_data
            shp_files_with_dk = find_shp_files_with_dk(shp_file_path)
            all_vector_data = gpd.GeoDataFrame()
            for shp_file in shp_files_with_dk:
                try:
                    # 使用改进的SHP编码处理模块
                    logger.debug(f"开始读取SHP文件: {shp_file}")
                    vector_data = read_shp_safe(shp_file)
                except Exception as e:
                    logger.error(f"读取SHP文件失败 {shp_file}: {str(e)}")
                    raise
                all_vector_data = pd.concat([all_vector_data, vector_data], ignore_index=True)

            cbdxx = dfs['CBDKXX']
            merged_cbdxx_vector = pd.merge(cbdxx, all_vector_data, left_on='DKBM', right_on='DKBM', how='left')

            # 生成临时Excel文件作为数据源
            excel_file1 = os.path.join(temp_dir, 'family_info_temp.xlsx')
            excel_file2 = os.path.join(temp_dir, 'land_info_temp.xlsx')

            with pd.ExcelWriter(excel_file1, engine='openpyxl') as writer:
                merged_all.to_excel(writer, sheet_name='Sheet1', index=False)

            with pd.ExcelWriter(excel_file2, engine='openpyxl') as writer:
                merged_cbdxx_vector_df = merged_cbdxx_vector.drop(columns=['geometry'])
                merged_cbdxx_vector_df.to_excel(writer, sheet_name='Sheet1', index=False)

            # 调用openpyxl版本的摸底调查表生成功能
            generate_mdb_table(excel_file1, excel_file2, output_directory)

        except Exception as e:
            logger.error(f"发生错误: {e}")
            logger.error(traceback.format_exc())
            raise
        finally:
            conn.close()
            logger.debug("Database connection closed")

    elif choice == "3":
        print("\n=== 生成金寨县摸底调查表 ===")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")
        generate_jinzhai_mdb_tables_only(mdb_file_path, shp_file_path, output_directory)

    elif choice == "4":
        print("\n=== 生成承包地块调查表 ===")
        jtjjzz_name = input("请输入集体经济组织名称: ")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")
        generate_dkdc_tables_only(mdb_file_path, shp_file_path, output_directory, jtjjzz_name)

    elif choice == "5":
        print("\n=== 生成户主代表声明书 ===")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")
        generate_hzzz_sms_only(mdb_file_path, shp_file_path, output_directory)

    elif choice == "6":
        print("\n=== 生成金寨县公示表 ===")
        jtjjzz_name = input("请输入集体经济组织名称: ")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")
        generate_jinzhai_gsb_only(mdb_file_path, shp_file_path, output_directory, jtjjzz_name)

    elif choice == "7":
        print("\n=== 生成榆社县摸底调查表 ===")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")
        generate_yushe_mdb_tables_only(mdb_file_path, shp_file_path, output_directory)

    elif choice == "8":
        print("\n=== 生成榆社县一户一档材料 ===")
        mdb_file_path = input("请输入mdb文件: ")
        shp_file_path = input("请输入shp文件的路径: ")
        output_directory = input("请输入输出目录的路径: ")
        generate_yushe_yhzd_materials_only(mdb_file_path, shp_file_path, output_directory)

    else:
        print("无效选择，请重新运行程序")