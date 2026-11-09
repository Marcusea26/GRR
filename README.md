# GRR 报告生成器 · Gauge R&R Report Generator

> 📌 **自用项目** · GRR 数据处理及报告生成
>
> 基于 **AIAG MSA 4th Edition** 的量具能力分析（MSA）工具集。
> 读取 Excel 宽表形式的重复性测量数据，自动完成 **ANOVA 法** 与 **Xbar-R 法** 双算法 GRR 计算，
> 并输出一份带完整图表与解读说明的 Excel 报表。

本仓库包含 **3 个独立的桌面程序**（均为 Tkinter GUI + 单文件打包为 exe）：

| 程序 | 说明 | 推荐度 |
|---|---|---|
| [`GRR_V2.py`](GRR_V2.py) | **主力版本**。功能等同 V1，额外内置 `Help 使用说明` sheet，并在 GUI 中提供数据模板下载 | ⭐ 推荐日常使用 |
| [`GRR_Report.py`](GRR_Report.py) | V1 原始版本。无 Help sheet，GUI 较简单。适合需要干净报表或做版本对照时使用 | 备选 |
| [`临时.py`](临时.py) | **另一套独立工具**（STAR）。多 Sheet 上下限判定 + CPK 计算 + PDF 矢量制图，不做 GRR | 按需使用 |

---

## 目录

- [1. 目录结构](#1-目录结构)
- [2. 环境要求](#2-环境要求)
- [3. 快速开始](#3-快速开始)
- [4. 输入数据格式](#4-输入数据格式)
- [5. 计算原理](#5-计算原理)
- [6. 输出报表结构](#6-输出报表结构)
- [7. 验收判据与颜色编码](#7-验收判据与颜色编码)
- [8. GUI 使用说明](#8-gui-使用说明)
- [9. 打包为 exe](#9-打包为-exe)
- [10. 附加工具：临时.py (STAR)](#10-附加工具临时py-star)
- [11. 常见问题排查](#11-常见问题排查)
- [12. 数据文件清单](#12-数据文件清单)

---

## 1. 目录结构

```
GRR/
├── GRR_V2.py                  # ★ 主力程序：GRR 报表生成器 V2.0
├── GRR_Report.py              # GRR 报表生成器 V1.0（无 Help sheet）
├── 临时.py                     # ★ 附加工具：STAR 统计分析与制图
├── GRR_V2.spec                # PyInstaller 打包配置（V2，含模板文件）
├── GRR_Report.spec            # PyInstaller 打包配置（V1）
├── 模板.xlsx                   # GRR 数据输入模板（3 操作员 × 3 试验 × 10 零件）
├── 模板--整理.xlsx              # STAR 工具的输出占位模板（Reports sheet）
│
├── GRR/                       # 数据工作目录（输入样例 + 历史输出报表）
│   ├── 模板.xlsx               #   模板副本
│   ├── 模板 - 副本.xlsx        #   模板副本
│   ├── SPL.xlsx                #   实测数据样例
│   ├── GRR FT025_L.xlsx        #   实测数据样例（含上下限，200–5000 Hz）
│   ├── GRR数据模板_GRR_Report.xlsx      #   V2 输出的报表样例
│   ├── GRR FT025_L_GRR_Report.xlsx      #   V2 输出的报表样例
│   ├── GRR FT025_L report.xlsx          #   V1 输出的报表样例
│   ├── 模板_GRR_Report.xlsx             #   V2 输出的报表样例
│   └── 模板 - 副本_GRR_Report.xlsx      #   V1 输出的报表样例
│
├── build/                     # PyInstaller 中间产物（已 gitignore）
├── dist/                      # 打包出的 exe（已 gitignore）
└── .gitignore
```

> `build/` 与 `dist/` 为 PyInstaller 生成物，体积约 175 MB，已在 `.gitignore` 中排除。

---

## 2. 环境要求

### 2.1 Python 版本

本仓库在 **Python 3.14.7 (Windows x64)** 下开发与验证通过。
依赖较新，建议使用 **Python 3.11 及以上**。

### 2.2 依赖库

| 库 | 验证版本 | 用途 | 被哪些程序使用 |
|---|---|---|---|
| `pandas` | 3.0.6 | Excel 读写、DataFrame 运算 | 全部 |
| `numpy` | 2.5.3 | 数值计算、`sqrt` | 全部 |
| `scipy` | 1.18.1 | `scipy.stats.f` — F 分布，用于 ANOVA 交互作用显著性检验 | GRR_V2 / GRR_Report |
| `xlsxwriter` | 3.2.9 | Excel 写入与图表生成 | 全部 |
| `openpyxl` | 3.1.5 | 读取 Excel（STAR 工具指定此引擎） | 临时.py |
| `matplotlib` | 3.11.2 | GUI 内嵌图 + PDF 导出 | 临时.py |

`tkinter` 属于 Python 标准库，Windows 官方安装包默认包含，无需额外安装。

### 2.3 安装依赖

```bash
pip install pandas numpy scipy xlsxwriter openpyxl matplotlib
```

验证环境：

```bash
python -c "import pandas, numpy, scipy, xlsxwriter, openpyxl, matplotlib; print('OK')"
```

---

## 3. 快速开始

### 3.1 运行（源码方式）

```bash
python GRR_V2.py
```

启动后弹出 GUI：

1. **选择输入 Excel** —— 点击「浏览」选中宽表数据文件；
   输出路径会自动带出 `<输入文件名>_GRR_Report.xlsx`；
2. 如需模板，点击「下载数据模板」保存 `GRR_数据模板.xlsx` 到本地；
3. 点击 **「生成 GRR 报表」**，进度条滚动后自动打开生成的 Excel。

### 3.2 直接运行打包好的 exe

```bash
dist\GRR_V2.exe
```

无需安装 Python，双击即可运行。

### 3.3 作为库调用（无 GUI）

```python
from GRR_V2 import process_grr_report

process_grr_report("GRR/GRR FT025_L.xlsx", "输出报表.xlsx")
```

`process_grr_report()` 会完成 读取 → 校验 → 计算 → 写报表 全流程，
校验失败时抛出 `ValueError`（错误信息为中文，可直接展示给用户）。

---

## 4. 输入数据格式

### 4.1 Excel 宽表布局

数据为 **宽表**（每行一个频点，每列一个测量项）。表头固定在 **第 3 行**，数据从 **第 4 行** 开始：

| Excel 行 | 内容 | 说明 |
|---|---|---|
| 第 1 行 | `Done` / `Tolerance` / `Trial 1` / `Trial 2` / `Trial 3` | 分组标识，**程序不读取**，仅供人眼识别 |
| 第 2 行 | `Operator A` / `Operator B` / `Operator C` | 操作员分组标识，**程序不读取** |
| **第 3 行** | `Hz` \| `Lower Limit` \| `Upper Limit` \| `A1-1` \| `A1-2` \| … | **真正的表头行**，程序从此行读取 |
| 第 4 行起 | 频点数值与各测量值 | 数据区 |

### 4.2 列定义

| 列 | 必需 | 说明 |
|---|:---:|---|
| `Hz` | ✅ | 频点（X 轴），必须是数值 |
| `Lower Limit` | ⭕ | 下公差限（LSL）。缺失时 `%Tolerance` 相关指标为 `NaN` |
| `Upper Limit` | ⭕ | 上公差限（USL）。缺失同上 |
| 测量列 | ✅ | 命名格式 `{操作员字母}{试验号}-{零件号}`，例如 `A1-1` |

**测量列命名规则**（正则 `^([A-Za-z]+)(\d+)-(\d+)$`）：

```
A1-1  →  Operator = A, Trial = 1, Part = P1
B2-7  →  Operator = B, Trial = 2, Part = P7
C3-10 →  Operator = C, Trial = 3, Part = P10
```

> **要求**：至少 **2 个 Part、2 个 Operator、2 个 Trial**，否则程序报错退出。
> 不满足 `^([A-Za-z]+)(\d+)-(\d+)$` 的列会被自动忽略（可安全放置备注列）。

### 4.3 模板结构

仓库自带 [`模板.xlsx`](模板.xlsx)：**64 行 × 93 列**，包含

- 3 个操作员（A/B/C） × 3 次试验（1/2/3） × 10 个零件（1–10） = **90 个测量列**
- 列布局：`A1-1…C3-10` 占据 D 列至 CO 列，A/B/C 三列分别为 `Hz`、`Lower Limit`、`Upper Limit`
- **61 个数据行**，示例频点从 20 Hz 起

### 4.4 实际数据样例

[`GRR/GRR FT025_L.xlsx`](GRR/GRR%20FT025_L.xlsx) 是一份真实数据：60 行 × 93 列，
频点范围 **200 Hz – 5000 Hz**，上下限 `-61 dB / -55 dB`，10 零件 × 3 操作员 × 3 试验。

```python
# 验证读入结果
df_long, df_display, lower, upper = load_wide_excel_to_long("GRR/GRR FT025_L.xlsx")
# df_long : 5130 行，列 = Part, Operator, Trial, X, Y, Lower Limit, Upper Limit, CurveID
# lower=-61.0, upper=-55.0
```

---

## 5. 计算原理

程序对 **每一个频点** 独立计算 GRR，并同时给出两种算法结果以便交叉验证。

### 5.1 数据重整：宽表 → 长表

宽表被展开为长表，每行表示"某个零件 / 某个操作员 / 某次试验 / 某频点"的一次测量：

| Part | Operator | Trial | X | Y | Lower Limit | Upper Limit | CurveID |
|---|---|---|---|---|---|---|---|
| P1 | A | 1 | 200.0 | -59.218 | -61.0 | -55.0 | A_T1_P1 |
| … | | | | | | | |

`CurveID`（如 `A_T1_P1`）用于标识一条完整曲线，后续"原始曲线图"按此绘制。

### 5.2 ANOVA 法（`grr_anova_for_one_x`）

计算四个平方和：`SS_part`、`SS_operator`、`SS_interaction`、`SS_repeat`，
自由度分别为 `p-1`、`o-1`、`(p-1)(o-1)`、`p·o·(r-1)`。

**第一步：交互作用显著性检验**

```
F = MS_interaction / MS_repeat
p-value = 1 - F.cdf(F, df_interaction, df_repeat)
显著 ⟺ p-value ≤ α  (α = 0.05)
```

**第二步：按显著性分支计算方差分量**

| 分量 | 交互作用**显著** | 交互作用**不显著**（合并交互项与重复项） |
|---|---|---|
| 重复性 `var_repeat` | `MS_repeat` | `(SS_int + SS_rep)/(df_int + df_rep)` |
| 交互作用 `var_interaction` | `max((MS_int − MS_rep)/r, 0)` | `0` |
| 再现性 `var_reprod` | `var_operator + var_interaction` | `var_operator` |
| 操作员 `var_operator` | `max((MS_op − MS_int)/(p·r), 0)` | `max((MS_op − pooled_MS)/(p·r), 0)` |
| 零件 `var_part` | `max((MS_part − MS_int)/(o·r), 0)` | `max((MS_part − pooled_MS)/(o·r), 0)` |

```
var_grr   = var_repeat + var_reprod      # GRR 总方差
var_total = var_grr + var_part            # 总方差
```

所有负差值统一取 `max(..., 0)` 截断为 0。

### 5.3 Xbar-R 法（`grr_xbar_r_for_one_x`）

基于 AIAG 的极差法常量表（`K1_TABLE` / `K2_TABLE` / `K3_TABLE`，支持 2–10）：

| 常量 | 依赖参数 | 取值范围 |
|---|---|---|
| `K1` | 试验次数 `r` | 2 – 10 |
| `K2` | 操作员数 `o` | 2 – 10 |
| `K3` | 零件数 `p` | 2 – 10 |

```
R̄  = 各 (Part, Operator) 单元格极差的平均值
EV = R̄ · K1                                  # 设备变差 Equipment Variation
AV = sqrt( (Xdiff · K2)² − EV²/(p·r) )        # 评定人变差 Appraiser Variation
PV = Rp · K3                                 # 零件变差 Part Variation
GRR   = sqrt(EV² + AV²)
Total = sqrt(GRR² + PV²)
```

> 若 `r`/`o`/`p` 超出 2–10，程序抛出 `ValueError` 提示扩展 `K1_TABLE` 等常量表。

### 5.4 三大核心指标

| 指标 | 公式 | 含义 |
|---|---|---|
| **%Tolerance (GRR)** | `6 × σ_GRR / (USL − LSL) × 100` | GRR 占公差带宽的比例（基于 6σ） |
| **%Study Variation** | `σ_GRR / σ_total × 100` | GRR 占总变差的比例 |
| **NDC** | `floor(1.41 × σ_part / σ_GRR)` | 可区别的零件类别数（`1.41 = 2 × 1.41`） |

其中 `σ = √var`，`σ_total = √(σ_GRR² + σ_part²)`。
缺少上下限时 `%Tolerance` 返回 `NaN`，但 `%Study Variation` 与 `NDC` 仍可正常计算。

### 5.5 结果对照（真实样例 `GRR FT025_L.xlsx`）

57 个频点、10 零件 × 3 操作员 × 3 试验，运算耗时约 1.4 s：

| 频点 | ANOVA %Tol | ANOVA %SV | ANOVA NDC | XR %Tol | XR %SV | XR NDC |
|---:|---:|---:|---:|---:|---:|---:|
| 200 Hz | 11.83 | 27.26 | 4 | 11.55 | 34.81 | 3 |
| 212 Hz | 10.99 | 25.97 | 5 | 10.69 | 32.27 | 4 |
| 236 Hz | 10.18 | 23.93 | 5 | 9.91 | 29.06 | 4 |
| 4250 Hz | 1.06 | 6.39 | 22 | 0.82 | 4.78 | 29 |
| 5000 Hz | 1.02 | 14.37 | 9 | 0.78 | 9.18 | 15 |

其中 **21 / 57** 个频点的 `Operator × Part` 交互作用显著（p ≤ 0.05）。
可见低频段 GRR 偏高（接近 10% 判据红线），高频段表现良好 —— 这正是频域类 GRR 报表
需要"逐频点看曲线"而非"只看一个汇总值"的原因。

---

## 6. 输出报表结构

生成的 Excel 包含 **8 个工作表**（V1 为 7 个，无 `Help 使用说明`）：

### 6.1 `Help 使用说明`（仅 V2）

报表打开时的**默认显示页**（深蓝色标签 `#1F4E79`），共 11 个章节：

1. GRR 与 MSA 概述
2. 方差分量模型 (Variance Components)
3. 两种算法对比：ANOVA 法 vs Xbar-R 法
4. AIAG 验收准则与颜色编码
5. AIAG K1 / K2 / K3 系数表（Xbar-R 法专用）
6. 所有公式汇总
7. Excel 报表结构（逐 sheet 解读）
8. 图表解读
9. 输入数据格式约定
10. 处理流程与异常处理
11. 解读示例（虚拟案例）

页面设置为 A4 横向、1 页宽自适应、隐藏网格线，含页眉页脚。

### 6.2 `Input Data` — 输入信息与宽表回显

- `A1:B23`：本次分析的基础信息（零件/试验/操作员数量、σ 数、α、各判据阈值、X/Y 轴量程与单位等）
- `A25:B31`：两算法的 `%Tolerance` 与 `%Study Variation` 汇总对比
- 其后为完整的原始宽表数据

### 6.3 `Normalized Data` — 标准化长表

宽表展开后的长表（`Part, Operator, Trial, X, Y, Lower Limit, Upper Limit, CurveID`），
样例中为 **5130 行**。便于二次分析与透视。

### 6.4 `GRR Results` — 全量统计结果

**46 列**的完整结果表，每行一个频点，列前缀区分算法：

| 前缀 | 含义 |
|---|---|
| `ANOVA_*` | ANOVA 法结果（22 列） |
| `XR_*` | Xbar-R 法结果（19 列） |
| 无前缀 | 频点 `X`、上下限、`Has Tolerance Info` |

包含 `SD_Repeat`、`SD_Reprod`、`SD_Part`、`SD_GRR`、`SD_Total`、
`%Tolerance_{GRR,EV,AV,PV,OP,OPI}`、`%StudyVariation_{GRR,EV,AV,PV,OP,OPI}`、
`p_interaction`、`interaction_significant`、`K1/K2/K3`、`R̄`、`Xdiff`、`Rp` 等。

### 6.5 `GRR by Frequency` — 精简关键指标

**12 列**，只保留最常看的核心指标，便于快速通读与排序：

```
X | Lower Limit | Upper Limit | Has Tolerance Info
ANOVA_SD_GRR | ANOVA_%Tolerance_GRR | ANOVA_%StudyVariation_GRR | ANOVA_NDC
XR_SD_GRR    | XR_%Tolerance_GRR    | XR_%StudyVariation_GRR    | XR_NDC
```

### 6.6 `Charts` — 三大汇总图

| 插入位置 | 图表 | 说明 |
|---|---|---|
| `B2` | `%Tolerance (GRR) by Frequency` | ANOVA + XR 两条曲线，叠加 10% / 20% 限值线 |
| `B32` | `%Study Variation (GRR) by Frequency` | 同上，含 10% / 20% 限值线 |
| `B62` | `NDC by Frequency` | 含 NDC=5 / NDC=2 限值线 |

X 轴为 **对数刻度**（10 Hz – 20000 Hz，`log_base=10`），符合频域数据的对数分布特性。

### 6.7 `Raw Curves` — 原始曲线图

按 `CurveID` 逐条绘制每个「操作员×试验×零件」的完整频响曲线，
并叠加红色上下限线。图表采用 **双列布局**（左 B 列 / 右 K 列，行距 20 行），
单图尺寸 720×360，X 轴对数刻度 10–20000 Hz。

### 6.8 `MSA 4 Charts` — MSA 四图

AIAG 经典四图布局：

| 位置 | 图表 |
|---|---|
| `A2` | Xbar & Range — Study Variation as a Percentage of Control Limits |
| `H2` | ANOVA — Study Variation as a Percentage of Control Limits |
| `A26` | Xbar & Range — %Study Variation（含限值） |
| `H26` | ANOVA — %Study Variation（含限值） |

Y 轴范围分别为 `0–120`（%Tol）与 `0–100`（%SV）。

---

## 7. 验收判据与颜色编码

### 7.1 阈值（AIAG MSA 4th）

| 指标 | 合格 (Ok) | 临界 (Marginal) | 不合格 |
|---|---|---|---|
| %Tolerance | ≤ 10% | ≤ 20% | > 20% |
| %Study Variation | ≤ 10% | ≤ 20% | > 20% |
| NDC | ≥ 5 | ≥ 2 | < 2 |

其他常量：`NUMBER_OF_SIGMA = 6`、`ANOVA_ALPHA = 0.05`。

### 7.2 颜色约定

**图表限值线：**

| 颜色 | 含义 |
|---|---|
| 🟢 `#00FF00` 深绿 | Ok 限值线（10% / 20% / NDC=5） |
| 🟡 `#FFA500` 橙色 | Marginal 限值线 |
| 🔴 `#FF0000` 红色 | 原始曲线图的上下限线 |

**报表单元格：**

| 底色 | 含义 |
|---|---|
| `#C6EFCE` 绿 | 合格 |
| `#FFEB9C` 黄 | 临界 |
| `#FFC7CE` 红 | 不合格 |

> 判据阈值集中定义在源码顶部常量区，修改后所有图表与判定自动同步。

---

## 8. GUI 使用说明

### 8.1 V2（`GRR_V2.py`）— 850×650，Fluent 卡片式布局

| 组件 | 行为 |
|---|---|
| **输入文件** | 「浏览」选文件；输入框仅显示文件名，下方灰字显示完整目录 |
| **输出文件** | 「另存为」选路径；**留空时自动带出** `<输入名>_GRR_Report.xlsx` |
| **下载数据模板** | 从内置 `GRR/模板.xlsx` 复制一份保存到本地（源码运行与 exe 打包均可用） |
| **生成 GRR 报表** | 执行计算；运行中按钮与输入框禁用，防重复提交 |
| **状态栏** | 彩色状态提示：`●` 待命 / `⏳` 处理中 / `✅` 完成 / `❌` 失败 |

其他特性：Windows **DPI 感知**（高分屏字体不虚）、路径实时校验（区分「完整路径」与「仅文件名」）、
ttk 主题自动降级（vista → winnative → clam → default）、完成后**自动打开报表**。

### 8.2 V1（`GRR_Report.py`）— 980×320，简洁表单

标题 + 输入路径 + 输出路径 + Run 按钮，字体 `KaiTi`。适合脚本化或批处理场景。

---

## 9. 打包为 exe

使用 [PyInstaller](https://pyinstaller.org/) 打包为单文件、无控制台窗口的 exe。

### 9.1 命令

```bash
# V2（推荐，需先把 GRR/模板.xlsx 一起打包）
pyinstaller --onefile --noconsole --name GRR_V2 ^
  --add-data "GRR\模板.xlsx;GRR" GRR_V2.py

# V1
pyinstaller --onefile --noconsole --name GRR_Report GRR_Report.py
```

### 9.2 关于仓库中的 `.spec` 文件

`GRR_V2.spec` 与 `GRR_Report.spec` 已包含完整配置，直接执行 `pyinstaller GRR_V2.spec` 即可。

> ⚠️ **注意**：两个 spec 文件中的路径是**硬编码的绝对路径**
> （`E:\Python_project\Python_projects\PycharmProjects\GRR\...`）。
> 在其他机器上打包前，需将 `Analysis` 中的脚本路径与 `datas` 路径改为当前项目的实际路径。

`GRR_V2.spec` 的关键差异 —— 模板文件通过 `datas` 一并打包：

```python
datas=[('...\\GRR\\模板.xlsx', 'GRR')],
```

运行时由 `get_template_path()` 解析路径，兼容两种模式：

```python
if getattr(sys, "frozen", False):
    base_dir = sys._MEIPASS      # PyInstaller 解压的临时目录
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))
return os.path.join(base_dir, "GRR", "模板.xlsx")
```

### 9.3 打包产物

| 产物 | 体积 |
|---|---|
| `dist/GRR_V2.exe` | ~83.5 MB |
| `dist/GRR_Report.exe` | ~83.5 MB |

体积主要来自 numpy / scipy / pandas / matplotlib 的打包（`build/GRR_*/GRR_*.pkg` 约 83 MB）。
`build/*/warn-*.txt` 中的 351 条 "missing module" 警告均为 Windows 上
POSIX 专属模块（`posix`、`_manylinux`、`multiprocessing.set_start_method` 等）的常规提示，**可忽略**。

---

## 10. 附加工具：临时.py (STAR)

**S**tatistics **T**ransformation & **A**nalysis **R**enderer — 与 GRR 无关的独立工具，
用于 **多 Sheet 的上下限判定 + CPK 计算 + PDF 矢量制图**。

### 10.1 功能

- 读取一个 **多 Sheet** Excel，每个 Sheet 独立处理
- 每个 Sheet 自动计算：**平均值**、**标准差**、**CPK**、**偏差值**（离差平方和）
- 按上/下框线判定 **PASS / FAIL**，汇总出通过率
- 导出 **Excel**（`<原名>--整理.xlsx`）与 **PDF**（`<原名>--制图.pdf`）
- GUI 右侧内嵌 matplotlib 实时预览，可切换 Sheet 查看

### 10.2 GUI 选项

| 选项 | 说明 |
|---|---|
| `Skip Data Without Limits` | 跳过既无上限也无下限的 Sheet |
| `Single-valued test item` | 散点 + 直方图模式（单一测值项） |
| `Sort by deviation` | 按偏差值升序排序，便于快速定位最差样本 |

### 10.3 输入 Sheet 的行列约定

| 位置 | 含义 |
|---|---|
| 第 0 列 | 样本序号 SN |
| 第 1 行 / 第 6 列 | 平均值行 |
| 第 2 行 / 第 0 列 | **上框线**（`B3`=10 Hz、`C3`=20000 Hz 为频率轴量程） |
| 第 3 行 / 第 6 列 | 标准差行 |
| 第 4 行 / 第 0 列 | **下框线**（`B5`=1.33、`C5`=1.33 为 CPK 目标） |
| 第 5 行起 | 数据区（SN 逐行） |
| 第 7 列起 | 各频点数据列 |
| 最后一列 | **偏差值** = 各频点相对平均值的离差平方和 |

CPK 计算：`CPK = min(|mean − upper|, |mean − lower|) / (3 × std)`。

### 10.4 输出 Excel 结构

| Sheet | 内容 |
|---|---|
| `Summary` | 各 Sheet × 各 SN 的 PASS/FAIL 与通过率（置于最前） |
| `Reports` | 各 Sheet 的图表汇总页（含数据量与 MIN_CPK） |
| `<原 Sheet 名>` | 单 Sheet 明细：原始数据 + 折线图 + CPK 图 |

---

## 11. 常见问题排查

| 现象 | 原因 | 处理 |
|---|---|---|
| `Excel格式不正确：至少需要4行、4列` | 文件不是宽表结构 | 使用 `模板.xlsx` 套用格式 |
| `未找到 Hz 列，请确认第3行存在 'Hz'` | 表头不在第 3 行 | 检查表头行位置 |
| `未找到测量列，如 A1-1 / B1-2 / C2-10` | 测量列命名不符正则 | 改为 `{字母}{试验号}-{零件号}` 格式 |
| `Part数量不足，至少需要2个` | 零件数 < 2 | 检查零件号是否重复或缺失 |
| `Operator数量不足，至少需要2个` | 操作员数 < 2 | 同上 |
| `Trial数量不足，至少需要2次试验` | 试验数 < 2 | 同上 |
| `GRR计算结果为空，请检查输入数据格式是否完整` | 所有频点都因数据不足被跳过 | 确认每个频点都有完整测量值 |
| `Trial 数 N 暂不支持，请扩展 K1_TABLE` | 试验数超出 2–10 | 扩展源码顶部 `K1_TABLE` |
| exe 双击无反应 | 被杀毒软件拦截 | 添加信任项；或改用源码运行 |
| 模板下载报"内置模板缺失" | 打包时未包含 `datas` | 按 §9.2 修正 spec 的 `datas` 路径 |
| 报表中 `%Tolerance` 全为 `NaN` | 输入缺少 `Lower Limit` / `Upper Limit` 列 | 补齐上下限列 |
| 中文显示为乱码 | 控制台编码问题 | Windows 终端执行 `chcp 65001` |

---

## 12. 数据文件清单

| 文件 | 类型 | 说明 |
|---|---|---|
| `模板.xlsx` | 输入模板 | 64×93，含示例数据；GRR 工具的数据模板 |
| `模板--整理.xlsx` | 输出占位 | 仅含空的 `Reports` sheet，`临时.py` 运行时作为初始工作簿 |
| `GRR/GRR FT025_L.xlsx` | 实测输入 | 200–5000 Hz，上下限 −61/−55 dB，10×3×3 |
| `GRR/SPL.xlsx` | 实测输入 | 与模板同构（64×93） |
| `GRR/GRR FT025_L report.xlsx` | V1 输出 | 7 sheet，无 Help |
| `GRR/GRR FT025_L_GRR_Report.xlsx` | V2 输出 | 8 sheet，首个为 Help |
| `GRR/GRR数据模板_GRR_Report.xlsx` | V2 输出 | 136 频点的大样例 |
| `GRR/模板_GRR_Report.xlsx` | V2 输出 | 61 频点样例 |
| `GRR/模板 - 副本_GRR_Report.xlsx` | V1 输出 | 精简版样例 |

---

## 附：源码速查

| 位置 | 关键函数 |
|---|---|
| `GRR_V2.py:310` | `get_template_path()` — 模板路径解析（源码 / exe 双模式） |
| `GRR_V2.py:332` | `load_wide_excel_to_long()` — 宽表 → 长表 |
| `GRR_V2.py:454` | `grr_anova_for_one_x()` — ANOVA 法 GRR |
| `GRR_V2.py:582` | `grr_xbar_r_for_one_x()` — Xbar-R 法 GRR |
| `GRR_V2.py:684` | `calculate_grr_all_x()` — 逐频点批量计算 |
| `GRR_V2.py:1039` | `write_help_sheet()` — 生成 Help 说明页 |
| `GRR_V2.py:1712` | `write_report()` — 输出 8 个 sheet 的报表 |
| `GRR_V2.py:2313` | `process_grr_report()` — 端到端入口 |
| `GRR_V2.py:2341` | `class GRRGUI` — V2 图形界面 |
| `临时.py:192` | `class ExcelProcessor` — STAR 数据处理 |
| `临时.py:620` | `class DOVEGUI` — STAR 图形界面与 PDF 导出 |

> 所有源码使用 **UTF-8 编码 + CRLF 换行**，在 Windows 上直接以 IDE 或 VS Code 打开不会乱码。
