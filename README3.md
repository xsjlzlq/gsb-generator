# 农村土地二轮承包到期后再延长三十年承包信息公示表生成工具

## 功能介绍

本工具用于自动生成农村土地二轮承包到期后再延长三十年的承包信息公示表。通过读取指定的MDB数据库和SHP矢量地块文件，生成按村组分类的Excel公示表格。

主要功能包括：
- 读取MDB数据库中的承包方信息、家庭成员、地块信息等数据
- 读取SHP矢量地块文件的空间信息
- 将数据合并整理并按村组生成公示表格
- 自动设置公示表格式，包括标题、表头、单元格合并等
- 提供图形界面，便于操作

## 系统要求

- Windows 操作系统 (Windows 7/8/10/11)
- Microsoft Access 数据库引擎（用于读取MDB文件）
- 需要安装以下Python依赖包（如使用可执行文件则无需安装）：
  - pandas
  - geopandas
  - pyodbc
  - openpyxl
  - PyQt5
  - pyinstaller（仅用于构建可执行文件）

## 使用说明

### 使用预编译的可执行文件

1. 双击运行`公示表生成工具.exe`或`公示表生成工具_单文件版.exe`
2. 在界面中填写以下信息：
   - 集体经济组织名称（如XX村民委员会）
   - 选择MDB数据库文件（包含承包方、家庭成员、地块信息的Access数据库）
   - 选择SHP矢量文件所在的文件夹（包含地块空间信息的矢量文件）
   - 选择输出目录（生成的公示表将保存在此目录）
3. 点击"生成公示表"按钮开始处理
4. 查看日志信息区域的处理进度和结果
5. 处理完成后，可在指定的输出目录中找到生成的公示表Excel文件

### 使用源代码运行

1. 确保已安装所需的Python依赖包
2. 运行`GenerateGSB_GUI_Wrapper.py`启动应用程序
3. 或者在命令行中直接使用原始脚本：
   ```
   python GenerateGSB.py
   ```
   并按提示输入所需信息

## 目录结构

- `GenerateGSB.py` - 核心数据处理模块
- `GenerateGSB_GUI.py` - 图形界面模块
- `GenerateGSB_GUI_Wrapper.py` - 界面启动器和异常处理
- `logger_util.py` - 日志工具类
- `build.bat` - 一键打包脚本
- `打包说明.md` - 打包和发布指南
- `logs/` - 日志文件目录

## 数据库结构要求

工具需要以下数据表：
- `CBDKXX` - 承包地块信息表
- `CBF` - 承包方信息表
- `CBF_JTCY` - 承包方家庭成员表
- `FBF` - 发包方信息表

SHP矢量文件应包含地块编码(DKBM)字段，用于与CBDKXX表关联。

## 问题排查

如果遇到问题，请检查：
1. 日志文件夹（logs目录）中的详细日志
2. 确认MDB数据库和SHP文件格式是否正确
3. 确认Microsoft Access Database Engine是否正确安装
4. 检查SHP文件名是否包含"DK"关键字

## 构建可执行文件

项目提供了两种打包方式：

### 使用一键打包脚本（推荐）

直接运行`build.bat`批处理文件，它将自动：
1. 清理之前的构建文件
2. 创建必要的打包配置
3. 生成文件夹版本和单文件版本的可执行文件
4. 显示打包结果的位置

### 手动打包

如需手动构建可执行文件，请运行：
```
pip install pyinstaller
pyinstaller --name="公示表生成工具" --windowed --hidden-import=pandas --hidden-import=geopandas --hidden-import=pyodbc --hidden-import=openpyxl --hidden-import=PyQt5 --hidden-import=numpy --hidden-import=shapely --hidden-import=fiona --hidden-import=rtree --hidden-import=pyproj GenerateGSB_GUI_Wrapper.py
```

打包完成后，可执行文件将位于`dist`目录中。详细说明请参阅`打包说明.md`文件。

## 版权信息

© 2025 农村土地承包信息管理系统 