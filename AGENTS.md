# AGENTS.md — 生成调查公示表格

## 项目概述

Windows-only Python 工具（Python 3.13），读取 MDB (Access) 数据库 + SHP 矢量文件，生成农村土地承包经营权调查/公示 Excel/Word 表格。PyQt5 图形界面。支持 南谯区 / 金寨县 / 榆社县 / 明光市 四个地区的不同表格模板。源码全部为中文注释与日志。

## 入口与快速验证

| 用途 | 命令 |
|------|------|
| 启动 GUI | `python GenerateGSB_GUI_Wrapper.py` |
| CLI 调试 | `python GenerateGSB.py`（交互式 1-8 功能菜单） |
| 语法/编译验证 | `python -m py_compile GenerateGSB.py GenerateGSB_GUI.py` |

GUI 启动器 `GenerateGSB_GUI_Wrapper.py` 设置日志与全局异常处理器（弹 Windows MessageBox）后再导入 `GenerateGSB_GUI.py` 的 `MainWindow`。核心业务逻辑全在 `GenerateGSB.py`。

## 核心文件职责

- `GenerateGSB.py` — 全部业务逻辑：MDB 读取、SHP 读取与拓扑闭环、Excel/Word 表格生成、`generate_*_only()` 入口函数（GUI 直接调用这些）
- `GenerateGSB_GUI.py` — PyQt5 界面，每个功能一个 `QThread` 子线程类避免界面冻结
- `GenerateGSB_GUI_Wrapper.py` — 启动器，设置日志、全局异常处理器、MessageBox
- `shp_encoding_handler.py` — SHP 中文编码检测；`setup_chinese_encoding()` 在 GenerateGSB 导入时即被调用，会设置 GDAL_DATA 等环境变量。不要改动此调用顺序
- `logger_util.py` — `GUILogger` 通过 `pyqtSignal` 发日志到 GUI
- `Templates/` — Excel/Word 模板；`resources/` — 权属单位代码表与调查统计表

## GUI 功能选择 (标签页)

| 标签页 | 功能 | 需要发包方名称 | 输入 |
|--------|------|:--------:|------|
| 南谯区 | 公示表 / 摸底调查表 / 承包地块调查表 / 户主代表声明书 / 一户一档封面 | 公示表、承包地块调查表需要 | MDB+SHP |
| 金寨县 | 摸底调查表 / 公示表 / 土地承包经营权调查信息表 | 公示表需要 | MDB+SHP |
| 榆社县 | 摸底调查表 / 公示表 / 一户一档材料 | 否 | MDB+SHP |
| 明光市 | 发包方调查表 / 一户一档清单 / 移交技术资料清单 / 特殊群体统计表 | 否 | 前三者仅 MDB；特殊群体统计表仅需摸底表目录 + 封面日期 |

## MDB 数据库与 SHP 矢量要求

- **MDB 必需表**：`CBDKXX`（承包地块）、`CBF`（承包方）、`CBF_JTCY`（家庭成员）、`FBF`（发包方）、`CBHT`（承包合同）。连接串使用 `Microsoft Access Driver (*.mdb, *.accdb)`（pyodbc）。
- **SHP 文件夹输入**：
  - 地块图层：`find_shp_files_with_dk()` 递归查找文件名含 `DK` 的 `.shp`。
  - 一户一档地块调查表额外依赖同目录下的 `JZD.shp`（界址点）与 `JZX.shp`（界址线）。

## 榆社县一户一档材料核心规则与避坑 (高频易错)

- **外部统计表依赖 (`resources/`)**：
  - `榆社县发包方信息变更调查统计.xlsx`（匹配修改后代码）
  - `榆社县承包方信息变更调查统计.xlsx`（匹配承包方编码，用于调查记事生成）
  - `榆社县一户一档材料时间表.xlsx`（匹配发包方编码，各表单各签名栏时间，缺失则不写时间）
- **调查结果核实表 / 公示结果归户表**：
  - 表头汇总行中的“实测总面积”必须取承包方所有地块实测面积求和 `sum(SCMJM)`，**严禁直接赋值为合同总面积 `HTZMJM`**。
- **承包地块调查表 (单地块独立 Excel)**：
  - **界址点拓扑排序**：必须调用 `chain_parcel_boundary(sub_lines, dkbm)` 进行首尾闭环排序，严禁仅用 `JZXSM` 正则数字排序（共享界址线如 `DKBM_A/DKBM_B` 会造成顺序错乱）。
  - **界址点号提取**：必须由闭环拓扑链顶点提取点号，严禁直接取界址线要素的 `QJZDH`（反向数字化的界址线起点实为 `ZJZDH`，直接取会导致相邻两行点号重复并遗漏真实点号）。
  - **第一页排版铁律**：**第一页界址点区域严格使用模版默认行（16~40行），最多容纳 12 条线，第 13 个点（单页闭合点或跨页第 13 个点）最大落在第 39~40 行合并单元格 (`A39:A40`)。第一页严禁插行扩展！**
  - **多页分页规则**：界址线 > 12 条时在第 40 行下方插入分页符（`Break(id=40)`），第二页及后续页从第 41 行起插入（每页 3 行表头 + 35 行点线槽位 = 38 行，容纳 17 条线），签名栏平移至最后一页底部。
  - **界址线说明连续编号**：必须调用 `format_jzxsm_continuous(orig_sm, start_num, end_num)` 统一格式化为连续的 `由T1...至T2, 由T2...至T3`，消除共享边保留邻宗地视角文本的问题。
- **Word 模板处理**：
  - `公示无异议声明书`、`户主声明书`、`家庭承包申请书` 均为 docx，使用 `replace_bookmark_in_doc()` 替换书签或精确 run 文本替换。

## openpyxl 插行与合并单元格陷阱

- openpyxl 原生 `ws.insert_rows()` **不会自动平移下方已存在的合并单元格范围与行高**，直接调用会导致严重的单元格错位与样式崩塌。
- 在任何需要动态插行的位置，必须使用 `insert_rows_with_merges(ws, insert_pos, amount)` 函数。

## resources 权属代码表

- `3415242026权属单位代码表.xls` → 金寨县（`region_prefix='341524'`）
- `3411822026权属单位代码表.xls` → 明光市（在 fbfdc/yhqd/yjjsqd/tsqt 中硬编码）
- 新增地区需要 `resources/` 下有对应前缀的代码表，否则回退到原始 `FBFMC` 逻辑。

## 隐式依赖 (不在 requirements.txt 中)

- `xlrd` — 读取 `.xls` 代码表和 Excel 降级读取；缺失时相关功能会失败或降级
- `python-docx` — docx 模板（声明书/封面等）；导入失败仅打印警告并禁用功能
- `xlsxwriter` — 部分摸底表实现；缺失时回退 openpyxl

requirements.txt 只列 `pandas, geopandas, pyodbc, openpyxl, PyQt5, cx_Freeze`。

## 已知陷阱

- **numpy/pickle 循环导入**：`GenerateGSB.py`、`GenerateGSB_GUI.py`、`GenerateGSB_GUI_Wrapper.py` 顶部都预加载 `numpy.random._pickle`、`pickle`、`_pickle`。不要移除，新入口文件也要带上
- **SHP 编码**：优先级 `.cpg` 文件 > GBK > GB18030 > GB2312 > UTF-8；`shp_encoding_handler.py` 的 `install_missing_gdal_data()` 按 脚本目录 gdal-data → cwd gdal-data → fiona/gdal-data 找数据
- **Microsoft Access 驱动**：需要系统安装 "Microsoft Access Database Engine" (64位)；编译后的 .exe 同样依赖
- **日志**：写入 `logs/gsb_*.log`（核心逻辑）与 `logs/gsb_gui_*.log`（启动器）；GUI 内嵌实时彩色日志框
- **读取 Excel**：优先 openpyxl，`BadZipFile` 时回退 xlrd（旧 .xls 格式）
