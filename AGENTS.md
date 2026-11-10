# AGENTS.md — 生成调查公示表格

## 项目概述

Windows-only Python 工具（Python 3.13），读取 MDB (Access) 数据库 + SHP 矢量文件，生成农村土地承包经营权调查/公示 Excel/Word 表格。PyQt5 图形界面。支持 南谯区 / 金寨县 / 榆社县 / 明光市 四个地区的不同表格模板。源码全部为中文注释与日志。

## 入口

| 用途 | 命令 |
|------|------|
| 启动 GUI | `python GenerateGSB_GUI_Wrapper.py` |
| CLI 调试 | `python GenerateGSB.py`（交互式 1-8 功能菜单） |

GUI 启动器 `GenerateGSB_GUI_Wrapper.py` 设置日志与全局异常处理器（弹 Windows MessageBox）后再导入 `GenerateGSB_GUI.py` 的 `MainWindow`。核心逻辑全在 `GenerateGSB.py`。

## 核心文件职责

- `GenerateGSB.py` — 全部业务逻辑：MDB 读取、SHP 编码读取、Excel/Word 表格生成、`generate_*_only()` 入口函数（GUI 直接调用这些）
- `GenerateGSB_GUI.py` — PyQt5 界面，每个功能一个 `QThread` 子线程类避免界面冻结
- `GenerateGSB_GUI_Wrapper.py` — 启动器，设置日志、全局异常处理器、MessageBox
- `shp_encoding_handler.py` — SHP 中文编码检测；`setup_chinese_encoding()` 在 GenerateGSB 导入时即被调用，会设置 GDAL_DATA 等环境变量。不要改动此调用顺序
- `logger_util.py` — `GUILogger` 通过 `pyqtSignal` 发日志到 GUI
- `Templates/` — Excel/Word 模板；`resources/` — 权属单位代码表

## GUI 功能选择 (标签页)

| 标签页 | 功能 | 需要发包方名称 | 输入 |
|--------|------|:--------:|------|
| 南谯区 | 公示表 / 摸底调查表 / 承包地块调查表 / 户主代表声明书 / 一户一档封面 | 公示表、承包地块调查表需要 | MDB+SHP |
| 金寨县 | 摸底调查表 / 公示表 / 土地承包经营权调查信息表 | 公示表需要 | MDB+SHP |
| 榆社县 | 摸底调查表 / 公示表 / 一户一档材料 | 否 | MDB+SHP |
| 明光市 | 发包方调查表 / 一户一档清单 / 移交技术资料清单 / 特殊群体统计表 | 否 | 前三者仅 MDB；特殊群体统计表仅需摸底表目录 + 封面日期 |

## 数据流 (公示表为例)

MDB → pandas 读 5 张表合并（`CBFBM` 前 14 位 join `FBFBM`）→ SHP 读取 → 写出临时 Excel（`tempfile.mkdtemp()` 下 `family_info_temp.xlsx` / `land_info_temp.xlsx`）→ `merge_excel_files()` 按 `FBFMC` 中的组名（切分"合作社/委员会"）分组，每组输出一个 `公示表（组名）.xlsx`。中间产物不落盘到输出目录。

## MDB 数据库要求

必需表：`CBDKXX`（承包地块）、`CBF`（承包方）、`CBF_JTCY`（家庭成员）、`FBF`（发包方）、`CBHT`（承包合同）。连接串使用 `Microsoft Access Driver (*.mdb, *.accdb)`（pyodbc）。

SHP 输入是**文件夹**：`find_shp_files_with_dk()` 递归找文件名含 `DK` 的 `.shp`，其余文件被忽略。

## resources 权属代码表

- `3415242026权属单位代码表.xls` → 金寨县（`region_prefix='341524'` 是 `GenerateJinzhaiSurveyInfoThread` 的默认值，`load_ownership_code_table()` 按前缀搜索）
- `3411822026权属单位代码表.xls` → 明光市（在 fbfdc/yhqd/yjjsqd/tsqt 中硬编码路径）

新增地区需要 `resources/` 下有对应前缀的代码表，否则回退到原始 `FBFMC` 逻辑。

## 隐式依赖 (不在 requirements.txt 中)

- `xlrd` — 读取 `.xls` 代码表和 Excel 降级读取；缺失时相关功能会失败或降级
- `python-docx` — docx 模板（声明书/封面等）；导入失败仅打印警告并禁用功能
- `xlsxwriter` — 部分摸底表实现；缺失时回退 openpyxl

requirements.txt 只列 `pandas, geopandas, pyodbc, openpyxl, PyQt5, cx_Freeze`。

## 构建 (PyInstaller)

- 仓库内当前**没有 `.spec` 文件**，也没有 `build.bat` / `打包说明.md`（README3.md 仍引用它们，已过时）
- `build/` 与 `dist/` 是历史打包残留：`dist/农村土地承包信息公示表生成工具.exe`（文件夹版，Python 3.13）
- 打包必须带数据文件：`Templates/`、`gdal-data/`、`resources/`；否则模板与编码检测失败
- README3.md 给出了手动 `pyinstaller --name=... --hidden-import=...` 命令

## 已知陷阱

- **numpy/pickle 循环导入**：`GenerateGSB.py`、`GenerateGSB_GUI.py`、`GenerateGSB_GUI_Wrapper.py` 顶部都预加载 `numpy.random._pickle`、`pickle`、`_pickle`。不要移除，新入口文件也要带上
- **SHP 编码**：优先级 `.cpg` 文件 > GBK > GB18030 > GB2312 > UTF-8；`shp_encoding_handler.py` 的 `install_missing_gdal_data()` 按 脚本目录 gdal-data → cwd gdal-data → fiona/gdal-data 找数据
- **Microsoft Access 驱动**：需要系统安装 "Microsoft Access Database Engine"；编译后的 .exe 同样依赖
- **日志**：写入 `logs/gsb_*.log`（核心逻辑）与 `logs/gsb_gui_*.log`（启动器）；GUI 内嵌实时彩色日志框
- **读取 Excel**：优先 openpyxl，`BadZipFile` 时回退 xlrd（旧 .xls 格式）
