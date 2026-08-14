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


def dklb_code_to_name(code):
    """
    将 DKLB 代码转换为对应的地块类别名称。
    """
    code_name_map = {
        '10': "承包地块",
        '21': "自留地",
        '22': "机动地",
        '23': "开荒地",
        '99': "其他集体土地",
    }
    return code_name_map.get(str(code), str(code)) if code else ""


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
    code_name_map = {
        '01': "一等地",
        '02': "二等地",
        '03': "三等地",
        '04': "四等地",
        '05': "五等地",
        '06': "六等地",
    }
    return code_name_map.get(str(code), str(code)) if code else ""


def sfjbnt_code_to_name(code):
    """
    将 SFJBNT 代码转换为对应的是否基本农田名称。
    """
    code_name_map = {
        '1': "是",
        '2': "否",
    }
    return code_name_map.get(str(code), str(code)) if code else ""


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
    for suffix in ('股份经济合作社', '经济合作社', '合作社', '村民委员会', '村委会'):
        if name.endswith(suffix):
            name = name[:-len(suffix)]
            break
    zumin = ''
    if name and ('组' in name or name.endswith(('队', '社'))):
        zumin = name
    return zhen, cun, zumin


def _clean_filename(name):
    """清理文件名中的非法字符"""
    import re
    return re.sub(r'[\\/:*?"<>|\r\n\t]', '', str(name)).strip()


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
                    row[11] = land_count           # K 地块数
                    row[16] = scmjm_sum            # P 实测总面积
                    row[18] = htmjm_sum            # R 合同总面积
                if f is not None:
                    row[5] = f['CYXM'] or ''       # E 姓名
                    row[6] = xb_code_to_name(f['CYXB'])      # F 性别
                    row[7] = _calc_age_from_id(f['CYZJHM'])  # G 年龄
                    row[8] = yhzgx_code_to_name(f['YHZGX'])  # H 与承包方代表关系
                    row[9] = f['CYBZSM'] or f['CYBZ'] or ''  # I:J 成员备注
                if l is not None:
                    row[12] = l['DKMC'] or ''      # L:M 地块名称
                    row[14] = str(l['DKBM'])[-5:]  # N 地块编码
                    scmjm = to_num(l['SCMJM']) if 'SCMJM' in l.index else None
                    if scmjm is None:
                        scmjm = to_num(l['HTMJM'])
                    row[15] = scmjm                # O 实测面积
                    row[17] = to_num(l['HTMJM'])   # Q 合同面积
                    row[19] = "种植业"             # S 土地用途
                    row[20] = dklb_code_to_name(l['DKLB'])          # T 地块类别
                    row[21] = tdlylx_code_to_name(l['TDLYLX'])      # U 土地利用类型
                    row[22] = dldj_code_to_name(l['DLDJ'])          # V 地力等级
                    row[23] = sfjbnt_code_to_name(l['SFJBNT'])      # W 是否基本农田
                    row[24] = l['DKDZ'] or ''      # X 东至
                    row[25] = l['DKNZ'] or ''      # Y 南至
                    row[26] = l['DKXZ'] or ''      # Z 西至
                    row[27] = l['DKBZ'] or ''      # AA 北至
                    row[28] = l['DKBZXX'] or ''    # AB 备注
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

            filename_base = f"榆社县{zhen}{cun}"
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

            # 填写表头（按md文档单元格标记：D5=乡镇名；H5=村名；J5=组名，默认"/"；O5=FBFBM）
            ws.cell(5, 4).value = zhen              # D5 乡镇名
            ws.cell(5, 8).value = cun               # H5 村名
            ws.cell(5, 10).value = zumin if zumin else '/'   # J5 组名，默认"/"
            ws.cell(5, 15).value = str(fbfbm)       # O5 FBFBM

            # 写入数据
            current = 9
            for row in data_rows:
                for c, v in row.items():
                    ws.cell(current, c, v)
                current += 1
            data_end = current - 1

            # 每户合并 A/B/C:D/K/P/R
            for sr, er in hh_spans:
                if er > sr:
                    for cols in [(1, 1), (2, 2), (3, 4), (11, 11), (16, 16), (18, 18)]:
                        ws.merge_cells(start_row=sr, start_column=cols[0], end_row=er, end_column=cols[1])
            # 每行合并 I:J 与 L:M
            for r in range(9, data_end + 1):
                ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
                ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=13)

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
                ws.cell(r, 14).number_format = '@'
                for c in (15, 16, 17, 18):
                    ws.cell(r, c).number_format = '0.00'

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
            ws.cell(sr_tail, 11).value = total_land
            ws.cell(sr_tail, 16).value = round(total_p, 2)
            ws.cell(sr_tail, 18).value = round(total_r, 2)

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
            merged_cbf_jtcy_cbf = pd.merge(cbf_jtcy, cbf[['CBFMC', 'CBFBM', 'CBFCYSL']], on='CBFBM', how='left')
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
            if child.get(qn("w:name")) != bookmark_name:
                continue

            bookmark_id = child.get(qn("w:id"))
            end_index = None
            for search_index in range(index + 1, len(children)):
                end_child = children[search_index]
                if end_child.tag == bm_end_tag and end_child.get(qn("w:id")) == bookmark_id:
                    end_index = search_index
                    break

            if end_index is None:
                return False

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

    choice = input("请输入选择 (1/2/3/4/5/6/7): ")

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

    else:
        print("无效选择，请重新运行程序")