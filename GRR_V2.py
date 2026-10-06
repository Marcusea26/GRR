"""
GRR_V2.py - GRR Report Generator (V2)
=====================================
完全复刻 GRR_Report.py 的全部功能，并新增一个 "Help 使用说明" sheet，
用于完整解读 GRR 的计算原理、报告内容与各项指标含义。

V2 与原版的差异：
    1. 新增 write_help_sheet(workbook) 函数：写出 Help sheet。
    2. 在 write_report() 开头调用 write_help_sheet(workbook)，
       并将该 sheet 设为打开时默认显示的 sheet。
    3. 工作簿左侧首个 tab 即为 "Help 使用说明"（深蓝色 tab 标识）。

其余逻辑、阈值、表格格式、GUI 行为与 GRR_Report.py 完全一致。
"""

import os
import re
import math
import random
import sys
import shutil
import warnings
import subprocess
import platform
import numpy as np
import pandas as pd
from scipy.stats import f
import tkinter as tk
from tkinter import filedialog, messagebox, font, ttk

warnings.filterwarnings("ignore", category=UserWarning)

# 版本号
VERSION = "V2.0"

# =========================================================
# 配置参数（主要用于 GRR 分析阈值、坐标轴、标题等）
# =========================================================
NUMBER_OF_SIGMA = 6

PCT_TOL_OK = 10
PCT_TOL_MARGINAL = 20

PCT_STUDY_OK = 10
PCT_STUDY_MARGINAL = 20

NDC_OK = 5
NDC_MARGINAL = 2

ANOVA_ALPHA = 0.05

X_SCALE_NAME = "Frequency"
X_UNIT = "Hz"
X_SCALE_TYPE = "Logarithmic"

Y_SCALE_NAME = "Magnitude"
Y_UNIT = "dB"
Y_SCALE_TYPE = "Linear"

CHART_TITLE_TOL = "%Tolerance (GRR) by Frequency"
CHART_TITLE_STUDY = "%Study Variation (GRR) by Frequency"
CHART_TITLE_NDC = "NDC by Frequency"

# Limit 线颜色
COLOR_DARK_GREEN = "#00FF00"
COLOR_GOLD = "#FFA500"
COLOR_LIMIT_RAW = "#FF0000"   # 原始曲线图上下限线颜色

# 原始曲线图的尺寸与布局
RAW_CHART_WIDTH = 720
RAW_CHART_HEIGHT = 360
RAW_CHART_X_MIN = 10
RAW_CHART_X_MAX = 20000

# 原始曲线图在 Excel 中的插入位置（双列布局）
RAW_CHART_START_ROW = 1   # 从第 2 行开始插入 (Excel 行号 = 1-based)
RAW_CHART_LEFT_COL = 1    # B 列
RAW_CHART_RIGHT_COL = 10  # K 列
RAW_CHART_ROW_STEP = 20   # 图表之间的行间距

# =========================================================
# AIAG 常数表 (K1/K2/K3 用于 Xbar-R 计算)
# =========================================================
K1_TABLE = {
    2: 0.8862,
    3: 0.5908,
    4: 0.4857,
    5: 0.4299,
    6: 0.3946,
    7: 0.3698,
    8: 0.3512,
    9: 0.3367,
    10: 0.3249
}

K2_TABLE = {
    2: 0.7071,
    3: 0.5231,
    4: 0.4467,
    5: 0.4030,
    6: 0.3742,
    7: 0.3534,
    8: 0.3375,
    9: 0.3249,
    10: 0.3146
}

K3_TABLE = {
    2: 0.7071,
    3: 0.5231,
    4: 0.4467,
    5: 0.4030,
    6: 0.3742,
    7: 0.3534,
    8: 0.3375,
    9: 0.3249,
    10: 0.3146
}


# =========================================================
# 工具函数
# =========================================================
def safe_sqrt(x):
    return math.sqrt(x) if x > 0 else 0.0


def get_k1(n):
    return K1_TABLE.get(n, None)


def get_k2(n):
    return K2_TABLE.get(n, None)


def get_k3(n):
    return K3_TABLE.get(n, None)


def random_color():
    return "#{:06x}".format(random.randint(0, 0xFFFFFF))


def parse_measurement_col(col_name):
    """
    将列名解析成 Operator/Trial/Part，例如 'A1-1' => Operator='A', Trial=1, Part='P1'
    """
    if pd.isna(col_name):
        return None

    s = str(col_name).strip()
    m = re.match(r'^([A-Za-z]+)(\d+)-(\d+)$', s)
    if not m:
        return None

    operator = m.group(1)
    trial = int(m.group(2))
    part = f"P{int(m.group(3))}"
    return operator, trial, part


def calc_pct_tol(sd_grr, lower_limit, upper_limit):
    """
    计算 %Tolerance = (6 * sd_grr) / (Upper - Lower) * 100
    若无有效上下限则返回 NaN
    """
    if pd.isna(lower_limit) and pd.isna(upper_limit):
        return np.nan

    if pd.isna(lower_limit) or pd.isna(upper_limit):
        return np.nan

    tolerance = upper_limit - lower_limit
    if tolerance <= 0:
        return np.nan

    return NUMBER_OF_SIGMA * sd_grr / tolerance * 100


def set_chart_style(chart, title, x_axis_name, y_axis_name):
    """
    通用图表格式设置（适用于三大汇总图：%Tol、%Study Variation、NDC）
    """
    chart.set_title({
        'name': title,
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 15, 'bold': True}
    })

    chart.set_x_axis({
        'name': x_axis_name,
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 12, 'bold': True},
        'min': 10,
        'max': 20000,
        'log_base': 10,
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside',
        'label_position': 'low'
    })

    chart.set_y_axis({
        'name': y_axis_name,
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 12, 'bold': True},
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside',
    })

    chart.set_style(1)
    chart.height = 550
    chart.width = 1000
    chart.set_plotarea({
        'border': {'color': '#808080', 'width': 1},
        'fill': {'color': '#F2F2F2'},
        'layout': {
            'x': 0.08,
            'y': 0.1,
            'width': 0.74,
            'height': 0.8,
        }
    })
    chart.set_legend({
        'layout': {
            'x': 0.83,
            'y': 0.1,
            'width': 0.14,
            'height': 0.8,
        }
    })


def set_raw_curve_chart_style(chart, title, y_min=None, y_max=None):
    """
    原始曲线图样式设置，支持 Y 轴自适应或指定范围
    """
    chart.set_title({
        'name': title,
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 13, 'bold': True}
    })

    x_axis_cfg = {
        'name': 'Hz',
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 11, 'bold': True},
        'min': RAW_CHART_X_MIN,
        'max': RAW_CHART_X_MAX,
        'log_base': 10,
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside',
        'label_position': 'low',
        'num_format': '0'
    }

    y_axis_cfg = {
        'name': Y_SCALE_NAME,
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 11, 'bold': True},
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside',
        'num_format': '0'
    }

    if y_min is not None and not pd.isna(y_min):
        y_axis_cfg['min'] = float(y_min)
    if y_max is not None and not pd.isna(y_max):
        y_axis_cfg['max'] = float(y_max)

    chart.set_x_axis(x_axis_cfg)
    chart.set_y_axis(y_axis_cfg)

    chart.set_style(1)
    chart.height = RAW_CHART_HEIGHT
    chart.width = RAW_CHART_WIDTH
    chart.set_plotarea({
        'border': {'color': '#808080', 'width': 1},
        'fill': {'color': '#F2F2F2'},
        'layout': {
            'x': 0.08,
            'y': 0.10,
            'width': 0.70,
            'height': 0.78,
        }
    })
    chart.set_legend({
        'position': 'right'
    })


def open_file_auto(file_path):
    """
    根据系统自动打开生成的文件
    """
    if not os.path.exists(file_path):
        return

    system_name = platform.system()
    try:
        if system_name == "Windows":
            os.startfile(file_path)
        elif system_name == "Darwin":
            subprocess.run(["open", file_path], check=False)
        else:
            subprocess.run(["xdg-open", file_path], check=False)
    except Exception:
        pass


def get_template_path():
    """
    获取数据模板文件的实际路径，兼容两种运行方式：

    1. 源码运行：模板位于源码目录的 GRR/模板.xlsx
    2. PyInstaller --onefile 打包：模板被 --add-data 一同打进 exe，
       PyInstaller 会解压到 sys._MEIPASS 临时目录

    返回模板文件的绝对路径。
    """
    # 优先用 PyInstaller 解压目录
    if getattr(sys, "frozen", False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_dir, "GRR", "模板.xlsx")


# =========================================================
# 读取宽表并转成长表
# =========================================================
def load_wide_excel_to_long(file_path):
    """
    Excel 宽表格式要求：
    - 第 3 行为表头 (含 Hz、Lower Limit、Upper Limit 等)
    - 第 4 行开始为数据行
    - 测量列必须形如 A1-1 / B1-2 / C2-10 等
    返回：
      df_long: 长表数据
      df_display: 原始宽表数据用于显示
      tolerance_lower_global / tolerance_upper_global: 如存在全局上下限则记录
    """
    df_raw = pd.read_excel(file_path, header=None)

    if df_raw.shape[0] < 4 or df_raw.shape[1] < 4:
        raise ValueError("Excel格式不正确：至少需要4行、4列")

    # 第3行做表头
    header_row = df_raw.iloc[2].tolist()
    data_rows = df_raw.iloc[3:].copy()
    data_rows.columns = header_row
    data_rows.reset_index(drop=True, inplace=True)

    # 去除空列、全空列
    data_rows = data_rows.loc[:, ~pd.isna(data_rows.columns)]
    data_rows = data_rows.dropna(axis=1, how='all')

    columns = list(data_rows.columns)

    x_col = None
    lower_col = None
    upper_col = None

    # 寻找 Hz、Lower Limit、Upper Limit 列
    for c in columns:
        cs = str(c).strip().lower()
        if cs == "hz":
            x_col = c
        elif "lower" in cs and "limit" in cs:
            lower_col = c
        elif "upper" in cs and "limit" in cs:
            upper_col = c

    if x_col is None:
        raise ValueError("未找到 Hz 列，请确认第3行存在 'Hz'")

    data_rows[x_col] = pd.to_numeric(data_rows[x_col], errors="coerce")

    # 全局上下限（如有）
    if lower_col is not None:
        data_rows[lower_col] = pd.to_numeric(data_rows[lower_col], errors="coerce")
        lower_notna = data_rows[lower_col].dropna()
        tolerance_lower_global = float(lower_notna.iloc[0]) if not lower_notna.empty else np.nan
    else:
        tolerance_lower_global = np.nan

    if upper_col is not None:
        data_rows[upper_col] = pd.to_numeric(data_rows[upper_col], errors="coerce")
        upper_notna = data_rows[upper_col].dropna()
        tolerance_upper_global = float(upper_notna.iloc[0]) if not upper_notna.empty else np.nan
    else:
        tolerance_upper_global = np.nan

    measurement_cols = []
    parsed_info = {}

    # 找出所有测量列
    for c in columns:
        parsed = parse_measurement_col(c)
        if parsed is not None:
            measurement_cols.append(c)
            parsed_info[c] = parsed

    if not measurement_cols:
        raise ValueError("未找到测量列，如 A1-1 / B1-2 / C2-10")

    long_rows = []
    # 构建长表
    for _, row in data_rows.iterrows():
        x_val = row[x_col]

        if pd.isna(x_val):
            continue

        lower_val = row[lower_col] if lower_col is not None else np.nan
        upper_val = row[upper_col] if upper_col is not None else np.nan

        for col in measurement_cols:
            y_val = row[col]
            if pd.isna(y_val):
                continue

            operator, trial, part = parsed_info[col]

            long_rows.append({
                "Part": part,
                "Operator": operator,
                "Trial": int(trial),
                "X": float(x_val),
                "Y": float(y_val),
                "Lower Limit": lower_val if not pd.isna(lower_val) else np.nan,
                "Upper Limit": upper_val if not pd.isna(upper_val) else np.nan,
                "CurveID": f"{operator}_T{trial}_{part}"
            })

    df_long = pd.DataFrame(long_rows)

    if df_long.empty:
        raise ValueError("没有解析到任何有效测量数据")

    # 仅用于显示宽表的辅助 DataFrame
    df_display = data_rows.copy()
    if lower_col is None:
        df_display["Lower Limit"] = np.nan
    if upper_col is None:
        df_display["Upper Limit"] = np.nan

    return df_long, df_display, tolerance_lower_global, tolerance_upper_global


# =========================================================
# ANOVA 法计算单个频点的 GRR
# =========================================================
def grr_anova_for_one_x(df_x, lower_limit, upper_limit, alpha=0.05):
    """
    对某个特定频点的长表数据进行 ANOVA GRR 计算，返回各类统计结果
    """
    df = df_x.copy()

    parts = sorted(df["Part"].unique())
    operators = sorted(df["Operator"].unique())
    trials = sorted(df["Trial"].unique())

    p = len(parts)
    o = len(operators)
    r = len(trials)

    grand_mean = df["Y"].mean()

    part_means = df.groupby("Part")["Y"].mean()
    op_means = df.groupby("Operator")["Y"].mean()
    cell_means = df.groupby(["Part", "Operator"])["Y"].mean()

    # 计算 SS
    ss_part = o * r * ((part_means - grand_mean) ** 2).sum()
    ss_operator = p * r * ((op_means - grand_mean) ** 2).sum()

    ss_interaction = 0.0
    var_interaction = 0.0
    for part in parts:
        for op in operators:
            if (part, op) in cell_means.index:
                cell_mean = cell_means.loc[(part, op)]
                ss_interaction += (cell_mean - part_means.loc[part] - op_means.loc[op] + grand_mean) ** 2
    ss_interaction *= r

    ss_repeat = 0.0
    for (part, op), grp in df.groupby(["Part", "Operator"]):
        cell_mean = grp["Y"].mean()
        ss_repeat += ((grp["Y"] - cell_mean) ** 2).sum()

    # 自由度
    df_part = p - 1
    df_operator = o - 1
    df_interaction = (p - 1) * (o - 1)
    df_repeat = p * o * (r - 1)

    # MS
    ms_part = ss_part / df_part if df_part > 0 else 0
    ms_operator = ss_operator / df_operator if df_operator > 0 else 0
    ms_interaction = ss_interaction / df_interaction if df_interaction > 0 else 0
    ms_repeat = ss_repeat / df_repeat if df_repeat > 0 else 0

    # 交互作用显著性检验
    p_value_interaction = 1.0
    if ms_repeat > 0 and df_interaction > 0 and df_repeat > 0:
        f_interaction = ms_interaction / ms_repeat
        p_value_interaction = 1 - f.cdf(f_interaction, df_interaction, df_repeat)

    interaction_significant = p_value_interaction <= alpha

    # 方差分量
    if interaction_significant:
        # 当交互作用显著时
        var_repeat = ms_repeat
        var_interaction = max((ms_interaction - ms_repeat) / r, 0)
        var_operator = max((ms_operator - ms_interaction) / (p * r), 0)
        var_part = max((ms_part - ms_interaction) / (o * r), 0)
        var_reprod = var_operator + var_interaction
    else:
        # 交互作用不显著，合并交互项与重复
        pooled_ss = ss_interaction + ss_repeat
        pooled_df = df_interaction + df_repeat
        pooled_ms = pooled_ss / pooled_df if pooled_df > 0 else 0

        var_repeat = pooled_ms
        var_operator = max((ms_operator - pooled_ms) / (p * r), 0)
        var_part = max((ms_part - pooled_ms) / (o * r), 0)
        var_reprod = var_operator
        var_interaction = 0.0

    var_grr = var_repeat + var_reprod
    var_total = var_grr + var_part

    # 标准差
    sd_repeat = safe_sqrt(var_repeat)
    sd_reprod = safe_sqrt(var_reprod)
    sd_part = safe_sqrt(var_part)
    sd_grr = safe_sqrt(var_grr)
    sd_total = safe_sqrt(var_total)
    sd_operator_only = safe_sqrt(var_operator)
    sd_interaction_only = safe_sqrt(var_interaction)

    pct_tol_grr = calc_pct_tol(sd_grr, lower_limit, upper_limit)
    pct_study_grr = sd_grr / sd_total * 100 if sd_total > 0 else np.nan
    ndc = int(np.floor(1.41 * sd_part / sd_grr)) if sd_grr > 0 else 0
    ndc = max(ndc, 1) if sd_part > 0 else 0

    return {
        "method": "ANOVA",
        "parts": p,
        "operators": o,
        "trials": r,
        "sd_repeat": sd_repeat,
        "sd_reprod": sd_reprod,
        "sd_part": sd_part,
        "sd_grr": sd_grr,
        "sd_total": sd_total,
        "sd_operator_only": sd_operator_only,
        "sd_interaction_only": sd_interaction_only,
        "pct_tol_grr": pct_tol_grr,
        "pct_study_grr": pct_study_grr,
        "pct_tol_ev": calc_pct_tol(sd_repeat, lower_limit, upper_limit),
        "pct_tol_av": calc_pct_tol(sd_reprod, lower_limit, upper_limit),
        "pct_tol_pv": calc_pct_tol(sd_part, lower_limit, upper_limit),
        "pct_tol_op": calc_pct_tol(sd_operator_only, lower_limit, upper_limit),
        "pct_tol_opi": calc_pct_tol(sd_interaction_only, lower_limit, upper_limit),
        "pct_study_ev": sd_repeat / sd_total * 100 if sd_total > 0 else np.nan,
        "pct_study_av": sd_reprod / sd_total * 100 if sd_total > 0 else np.nan,
        "pct_study_pv": sd_part / sd_total * 100 if sd_total > 0 else np.nan,
        "pct_study_op": sd_operator_only / sd_total * 100 if sd_total > 0 else np.nan,
        "pct_study_opi": sd_interaction_only / sd_total * 100 if sd_total > 0 else np.nan,
        "p_value_interaction": p_value_interaction,
        "interaction_significant": interaction_significant,
        "ndc": ndc
    }


# =========================================================
# Xbar-R 法（基于 AIAG K1/K2/K3）
# =========================================================
def grr_xbar_r_for_one_x(df_x, lower_limit, upper_limit):
    """
    对某个频点采用 Xbar & Range 方法计算 GRR
    """
    df = df_x.copy()

    parts = sorted(df["Part"].unique())
    operators = sorted(df["Operator"].unique())
    trials = sorted(df["Trial"].unique())

    p = len(parts)
    o = len(operators)
    r = len(trials)

    if p < 2:
        raise ValueError("Xbar-R 需要至少2个 Part")
    if o < 2:
        raise ValueError("Xbar-R 需要至少2个 Operator")
    if r < 2:
        raise ValueError("Xbar-R 需要至少2个 Trial")

    k1 = get_k1(r)
    k2 = get_k2(o)
    k3 = get_k3(p)

    if k1 is None:
        raise ValueError(f"Trial 数 {r} 暂不支持，请扩展 K1_TABLE")
    if k2 is None:
        raise ValueError(f"Operator 数 {o} 暂不支持，请扩展 K2_TABLE")
    if k3 is None:
        raise ValueError(f"Part 数 {p} 暂不支持，请扩展 K3_TABLE")

    # 计算各 (Part, Operator) 的 Range
    cell_ranges = []
    for (part, op), grp in df.groupby(["Part", "Operator"]):
        grp2 = grp.sort_values("Trial")
        y = grp2["Y"].values
        if len(y) >= 2:
            cell_ranges.append(np.max(y) - np.min(y))

    rbar = float(np.mean(cell_ranges)) if len(cell_ranges) > 0 else 0.0
    ev = rbar * k1  # Equipment Variation

    # 操作员均值差
    op_means = df.groupby("Operator")["Y"].mean().reindex(operators)
    if op_means.notna().sum() >= 2:
        xdiff = float(op_means.max() - op_means.min())
        av_term = (xdiff * k2) ** 2 - (ev ** 2) / (p * r)
        av = safe_sqrt(av_term)  # Appraiser Variation
    else:
        xdiff = 0.0
        av = 0.0

    # 零件均值差
    part_means = df.groupby("Part")["Y"].mean().reindex(parts)
    if part_means.notna().sum() >= 2:
        rp = float(part_means.max() - part_means.min())
        pv = rp * k3  # Part Variation
    else:
        rp = 0.0
        pv = 0.0

    grr = math.sqrt(ev ** 2 + av ** 2)     # GRR
    total = math.sqrt(grr ** 2 + pv ** 2)  # Total Variation

    pct_tol_grr = calc_pct_tol(grr, lower_limit, upper_limit)
    pct_study_grr = grr / total * 100 if total > 0 else np.nan

    ndc = int(np.floor(1.41 * pv / grr)) if grr > 0 else 0
    ndc = max(ndc, 1) if pv > 0 else 0

    return {
        "method": "Xbar & Range",
        "parts": p,
        "operators": o,
        "trials": r,
        "k1": k1,
        "k2": k2,
        "k3": k3,
        "rbar": rbar,
        "xdiff": xdiff,
        "rp": rp,
        "sd_repeat": ev,
        "sd_reprod": av,
        "sd_part": pv,
        "sd_grr": grr,
        "sd_total": total,
        "pct_tol_grr": pct_tol_grr,
        "pct_study_grr": pct_study_grr,
        "pct_tol_ev": calc_pct_tol(ev, lower_limit, upper_limit),
        "pct_tol_av": calc_pct_tol(av, lower_limit, upper_limit),
        "pct_tol_pv": calc_pct_tol(pv, lower_limit, upper_limit),
        "pct_study_ev": ev / total * 100 if total > 0 else np.nan,
        "pct_study_av": av / total * 100 if total > 0 else np.nan,
        "pct_study_pv": pv / total * 100 if total > 0 else np.nan,
        "ndc": ndc
    }


# =========================================================
# 计算所有频点的 GRR 结果
# =========================================================
def calculate_grr_all_x(df_long):
    """
    对长表中所有频点进行 ANOVA 和 Xbar-R 的 GRR 计算
    返回包含各频点统计结果的 DataFrame
    """
    x_values = sorted(df_long["X"].dropna().unique())
    rows = []

    for x in x_values:
        df_x = df_long[df_long["X"] == x].copy()
        if df_x.empty:
            continue

        # 至少需要2个 Part/Operator/Trial 才能计算
        if df_x["Part"].nunique() < 2 or df_x["Operator"].nunique() < 2 or df_x["Trial"].nunique() < 2:
            continue

        lower_vals = df_x["Lower Limit"].dropna().unique()
        upper_vals = df_x["Upper Limit"].dropna().unique()

        lower_limit = float(lower_vals[0]) if len(lower_vals) > 0 else np.nan
        upper_limit = float(upper_vals[0]) if len(upper_vals) > 0 else np.nan

        try:
            anova_res = grr_anova_for_one_x(df_x, lower_limit, upper_limit, alpha=ANOVA_ALPHA)
            xr_res = grr_xbar_r_for_one_x(df_x, lower_limit, upper_limit)
        except Exception:
            # 如果该频点数据异常，则跳过
            continue

        rows.append({
            "X": x,
            "Lower Limit": lower_limit,
            "Upper Limit": upper_limit,
            "Has Tolerance Info": not (pd.isna(lower_limit) or pd.isna(upper_limit)),

            "ANOVA_%Tolerance_GRR": anova_res["pct_tol_grr"],
            "ANOVA_%StudyVariation_GRR": anova_res["pct_study_grr"],
            "ANOVA_NDC": anova_res["ndc"],
            "ANOVA_SD_Repeat": anova_res["sd_repeat"],
            "ANOVA_SD_Reprod": anova_res["sd_reprod"],
            "ANOVA_SD_Part": anova_res["sd_part"],
            "ANOVA_SD_GRR": anova_res["sd_grr"],
            "ANOVA_SD_Total": anova_res["sd_total"],
            "ANOVA_SD_Operator_Only": anova_res["sd_operator_only"],
            "ANOVA_SD_Interaction_Only": anova_res["sd_interaction_only"],
            "ANOVA_p_interaction": anova_res["p_value_interaction"],
            "ANOVA_interaction_significant": anova_res["interaction_significant"],
            "ANOVA_%Tolerance_EV": anova_res["pct_tol_ev"],
            "ANOVA_%Tolerance_AV": anova_res["pct_tol_av"],
            "ANOVA_%Tolerance_PV": anova_res["pct_tol_pv"],
            "ANOVA_%Tolerance_OP": anova_res["pct_tol_op"],
            "ANOVA_%Tolerance_OPI": anova_res["pct_tol_opi"],
            "ANOVA_%StudyVariation_EV": anova_res["pct_study_ev"],
            "ANOVA_%StudyVariation_AV": anova_res["pct_study_av"],
            "ANOVA_%StudyVariation_PV": anova_res["pct_study_pv"],
            "ANOVA_%StudyVariation_OP": anova_res["pct_study_op"],
            "ANOVA_%StudyVariation_OPI": anova_res["pct_study_opi"],

            "XR_K1": xr_res["k1"],
            "XR_K2": xr_res["k2"],
            "XR_K3": xr_res["k3"],
            "XR_Rbar": xr_res["rbar"],
            "XR_Xdiff": xr_res["xdiff"],
            "XR_Rp": xr_res["rp"],

            "XR_%Tolerance_GRR": xr_res["pct_tol_grr"],
            "XR_%StudyVariation_GRR": xr_res["pct_study_grr"],
            "XR_NDC": xr_res["ndc"],
            "XR_SD_Repeat": xr_res["sd_repeat"],
            "XR_SD_Reprod": xr_res["sd_reprod"],
            "XR_SD_Part": xr_res["sd_part"],
            "XR_SD_GRR": xr_res["sd_grr"],
            "XR_SD_Total": xr_res["sd_total"],
            "XR_%Tolerance_EV": xr_res["pct_tol_ev"],
            "XR_%Tolerance_AV": xr_res["pct_tol_av"],
            "XR_%Tolerance_PV": xr_res["pct_tol_pv"],
            "XR_%StudyVariation_EV": xr_res["pct_study_ev"],
            "XR_%StudyVariation_AV": xr_res["pct_study_av"],
            "XR_%StudyVariation_PV": xr_res["pct_study_pv"],
        })

    return pd.DataFrame(rows)


# =========================================================
# 生成“频点-GRR”汇总表
# =========================================================
def build_grr_by_frequency_sheet(grr_df):
    """
    将关键结果按频点输出，便于查看各频点的 %Tol、%StudyVariation、NDC 等
    """
    if grr_df.empty:
        return pd.DataFrame(columns=[
            "X",
            "Lower Limit",
            "Upper Limit",
            "Has Tolerance Info",
            "ANOVA_SD_GRR",
            "ANOVA_%Tolerance_GRR",
            "ANOVA_%StudyVariation_GRR",
            "ANOVA_NDC",
            "XR_SD_GRR",
            "XR_%Tolerance_GRR",
            "XR_%StudyVariation_GRR",
            "XR_NDC",
        ])

    df_out = grr_df[[
        "X",
        "Lower Limit",
        "Upper Limit",
        "Has Tolerance Info",
        "ANOVA_SD_GRR",
        "ANOVA_%Tolerance_GRR",
        "ANOVA_%StudyVariation_GRR",
        "ANOVA_NDC",
        "XR_SD_GRR",
        "XR_%Tolerance_GRR",
        "XR_%StudyVariation_GRR",
        "XR_NDC",
    ]].copy()

    return df_out


# =========================================================
# 原始曲线数据准备 (转换为矩阵方便画线)
# =========================================================
def build_raw_curve_matrix(df_long):
    """
    将长表数据透视成：
      X, Lower Limit, Upper Limit, 各曲线列(每个 CurveID 对应一列)
    同时返回曲线元信息 (CurveID / Operator / Trial / Part)
    """
    pivot_df = df_long.pivot_table(
        index="X",
        columns="CurveID",
        values="Y",
        aggfunc="mean"
    ).reset_index()

    pivot_df = pivot_df.sort_values("X").reset_index(drop=True)

    limit_df = (
        df_long.groupby("X")[["Lower Limit", "Upper Limit"]]
        .first()
        .reset_index()
        .sort_values("X")
        .reset_index(drop=True)
    )

    pivot_df = pd.merge(limit_df, pivot_df, on="X", how="left")

    curve_cols = [c for c in pivot_df.columns if c not in ["X", "Lower Limit", "Upper Limit"]]
    pivot_df = pivot_df[["X", "Lower Limit", "Upper Limit"] + curve_cols]

    curve_meta = (
        df_long[["CurveID", "Operator", "Trial", "Part"]]
        .drop_duplicates()
        .sort_values(["Operator", "Trial", "Part"])
        .reset_index(drop=True)
    )

    return pivot_df, curve_meta


def add_raw_curve_chart(chart, sheet_name, data_row_count, x_col_idx, y_col_idx, series_name, color=None, width=1.1):
    """
    为原始曲线图表添加一条线 (X列为日志轴, Y列是曲线值)
    """
    series_cfg = {
        'name': series_name,
        'categories': [sheet_name, 1, x_col_idx, data_row_count, x_col_idx],
        'values': [sheet_name, 1, y_col_idx, data_row_count, y_col_idx],
        'line': {
            'width': width,
            'color': color if color else random_color()
        },
    }
    chart.add_series(series_cfg)


def add_raw_limit_series(chart, sheet_name, data_row_count, x_col_idx, lower_col_idx, upper_col_idx, has_lower, has_upper):
    """
    在原始曲线图中添加上下限线
    """
    if has_lower:
        chart.add_series({
            'name': 'Lower Limit',
            'categories': [sheet_name, 1, x_col_idx, data_row_count, x_col_idx],
            'values': [sheet_name, 1, lower_col_idx, data_row_count, lower_col_idx],
            'line': {'color': COLOR_LIMIT_RAW, 'width': 1.5},
        })
    if has_upper:
        chart.add_series({
            'name': 'Upper Limit',
            'categories': [sheet_name, 1, x_col_idx, data_row_count, x_col_idx],
            'values': [sheet_name, 1, upper_col_idx, data_row_count, upper_col_idx],
            'line': {'color': COLOR_LIMIT_RAW, 'width': 1.5},
        })


def get_chart_position(index):
    """
    将图按 2 列布局，基于 index 返回插入起始行列
    """
    row_block = index // 2
    col_side = index % 2

    row = RAW_CHART_START_ROW + row_block * RAW_CHART_ROW_STEP
    col = RAW_CHART_LEFT_COL if col_side == 0 else RAW_CHART_RIGHT_COL
    return row, col


# =========================================================
# 构建 MSA 4 Charts 数据
# =========================================================
def build_msa_four_chart_df(grr_df):
    """
    生成 MSA 4 Charts 所需的列，包括 X、各类 %Tol/%Study 及限值等
    """
    if grr_df.empty:
        return pd.DataFrame(columns=[
            "X",
            "XR_TVTOL", "XR_PVTOL", "XR_EVTOL", "XR_AVTOL", "XR_GRRTOL", "XR_TOL_OK", "XR_TOL_MARGINAL",
            "ANOVA_TVTOL", "ANOVA_PVTOL", "ANOVA_EVTOL", "ANOVA_AVTOL", "ANOVA_OPTOL", "ANOVA_OPITOL", "ANOVA_GRRTOL", "ANOVA_TOL_OK", "ANOVA_TOL_MARGINAL",
            "XR_PV", "XR_EV", "XR_AV", "XR_GRR", "XR_STUDY_OK", "XR_STUDY_MARGINAL",
            "ANOVA_PV", "ANOVA_EV", "ANOVA_AV", "ANOVA_OP", "ANOVA_OPI", "ANOVA_GRR", "ANOVA_STUDY_OK", "ANOVA_STUDY_MARGINAL"
        ])

    df = pd.DataFrame()
    df["X"] = grr_df["X"]

    # Xbar-R %Tol
    df["XR_TVTOL"] = 100.0
    df["XR_PVTOL"] = grr_df["XR_%Tolerance_PV"]
    df["XR_EVTOL"] = grr_df["XR_%Tolerance_EV"]
    df["XR_AVTOL"] = grr_df["XR_%Tolerance_AV"]
    df["XR_GRRTOL"] = grr_df["XR_%Tolerance_GRR"]
    df["XR_TOL_OK"] = PCT_TOL_OK
    df["XR_TOL_MARGINAL"] = PCT_TOL_MARGINAL

    # ANOVA %Tol
    df["ANOVA_TVTOL"] = 100.0
    df["ANOVA_PVTOL"] = grr_df["ANOVA_%Tolerance_PV"]
    df["ANOVA_EVTOL"] = grr_df["ANOVA_%Tolerance_EV"]
    df["ANOVA_AVTOL"] = grr_df["ANOVA_%Tolerance_AV"]
    df["ANOVA_OPTOL"] = grr_df["ANOVA_%Tolerance_OP"]
    df["ANOVA_OPITOL"] = grr_df["ANOVA_%Tolerance_OPI"]
    df["ANOVA_GRRTOL"] = grr_df["ANOVA_%Tolerance_GRR"]
    df["ANOVA_TOL_OK"] = PCT_TOL_OK
    df["ANOVA_TOL_MARGINAL"] = PCT_TOL_MARGINAL

    # Xbar-R %Study Variation
    df["XR_PV"] = grr_df["XR_%StudyVariation_PV"]
    df["XR_EV"] = grr_df["XR_%StudyVariation_EV"]
    df["XR_AV"] = grr_df["XR_%StudyVariation_AV"]
    df["XR_GRR"] = grr_df["XR_%StudyVariation_GRR"]
    df["XR_STUDY_OK"] = PCT_STUDY_OK
    df["XR_STUDY_MARGINAL"] = PCT_STUDY_MARGINAL

    # ANOVA %Study Variation
    df["ANOVA_PV"] = grr_df["ANOVA_%StudyVariation_PV"]
    df["ANOVA_EV"] = grr_df["ANOVA_%StudyVariation_EV"]
    df["ANOVA_AV"] = grr_df["ANOVA_%StudyVariation_AV"]
    df["ANOVA_OP"] = grr_df["ANOVA_%StudyVariation_OP"]
    df["ANOVA_OPI"] = grr_df["ANOVA_%StudyVariation_OPI"]
    df["ANOVA_GRR"] = grr_df["ANOVA_%StudyVariation_GRR"]
    df["ANOVA_STUDY_OK"] = PCT_STUDY_OK
    df["ANOVA_STUDY_MARGINAL"] = PCT_STUDY_MARGINAL

    return df


def add_series(chart, sheet_name, row_count, x_col, y_col, name, color=None, width=1.5, dash_type=None):
    """
    通用函数：为图表添加系列
    """
    line_cfg = {'width': width, 'color': color if color else random_color()}
    if dash_type:
        line_cfg['dash_type'] = dash_type

    chart.add_series({
        'name': name,
        'categories': [sheet_name, 1, x_col, row_count, x_col],
        'values': [sheet_name, 1, y_col, row_count, y_col],
        'line': line_cfg,
    })


def set_msa_same_style(chart, title, y_axis_name, y_min=None, y_max=None):
    """
    MSA 4 Charts 通用格式设置
    """
    chart.set_title({
        'name': title,
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 15, 'bold': True}
    })

    x_axis_cfg = {
        'name': 'Hz',
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 12, 'bold': True},
        'min': 10,
        'max': 20000,
        'log_base': 10,
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside',
        'label_position': 'low'
    }

    y_axis_cfg = {
        'name': y_axis_name,
        'name_font': {'name': 'Arial', 'color': 'black', 'size': 12, 'bold': True},
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside',
    }

    if y_min is not None:
        y_axis_cfg['min'] = y_min
    if y_max is not None:
        y_axis_cfg['max'] = y_max

    chart.set_x_axis(x_axis_cfg)
    chart.set_y_axis(y_axis_cfg)

    chart.set_style(1)
    chart.height = 450
    chart.width = 800
    chart.set_plotarea({
        'border': {'color': '#808080', 'width': 1},
        'fill': {'color': '#F2F2F2'},
        'layout': {
            'x': 0.08,
            'y': 0.1,
            'width': 0.74,
            'height': 0.8,
        }
    })
    chart.set_legend({
        'layout': {
            'x': 0.83,
            'y': 0.1,
            'width': 0.14,
            'height': 0.8,
        }
    })


# =========================================================
# V2 新增: Help 使用说明 sheet
# =========================================================
def write_help_sheet(workbook):
    """
    在生成的 Excel 报表中新增 "Help 使用说明" sheet，
    完整解读 GRR (Gauge R&R) 计算原理、报告结构、各项指标含义。
    """
    ws = workbook.add_worksheet("Help 使用说明")
    ws.set_tab_color("#1F4E79")
    # 设为打开时的默认显示 sheet
    ws.activate()
    ws.set_first_sheet()

    # ===== 页面设置 =====
    ws.hide_gridlines(2)  # 隐藏屏幕与打印网格线
    ws.set_landscape()
    ws.set_paper(9)       # A4
    ws.fit_to_pages(1, 0)  # 1 页宽，高度不限
    ws.set_margins(0.4, 0.4, 0.5, 0.5)
    ws.set_header('&L&"Microsoft YaHei UI,Bold"&14&K1F4E79GRR 报告解读&R&"Microsoft YaHei UI,Italic"&10V2.0')
    ws.set_footer('&L&"Microsoft YaHei UI,Italic"&9GRR_V2.py&R&"Microsoft YaHei UI,Italic"&9&P / &N')

    # 列宽
    ws.set_column("A:A", 2)   # 左空白
    ws.set_column("B:B", 30)  # 字段名
    ws.set_column("C:C", 74)  # 内容
    ws.set_column("D:D", 28)  # 阈值 / 补充
    ws.set_column("E:E", 2)   # 右空白

    # ============== 格式定义 ==============
    ZH = "Microsoft YaHei UI"
    MONO = "Consolas"
    f_h1 = workbook.add_format({
        "bold": True, "font_size": 22, "font_name": ZH,
        "font_color": "white", "bg_color": "#1F4E79",
        "align": "center", "valign": "vcenter", "border": 2,
        "border_color": "#0F2F4F"
    })
    f_h2 = workbook.add_format({
        "bold": True, "font_size": 14, "font_name": ZH,
        "font_color": "white", "bg_color": "#2E75B6",
        "align": "left", "valign": "vcenter",
        "border": 1, "border_color": "#1F4E79",
        "indent": 1
    })
    f_h3 = workbook.add_format({
        "bold": True, "font_size": 12, "font_name": ZH,
        "font_color": "white", "bg_color": "#5B9BD5",
        "align": "left", "valign": "vcenter",
        "border": 1, "border_color": "#2E75B6",
        "indent": 1
    })
    f_label = workbook.add_format({
        "bold": True, "font_name": ZH, "font_size": 10,
        "bg_color": "#D9EAF7", "border": 1, "border_color": "#B4C7E7",
        "valign": "top", "text_wrap": True, "indent": 1
    })
    f_label_alt = workbook.add_format({
        "bold": True, "font_name": ZH, "font_size": 10,
        "bg_color": "#EAF1F9", "border": 1, "border_color": "#B4C7E7",
        "valign": "top", "text_wrap": True, "indent": 1
    })
    f_value = workbook.add_format({
        "font_name": ZH, "font_size": 10,
        "bg_color": "#FFFFFF", "border": 1, "border_color": "#D9D9D9",
        "valign": "top", "text_wrap": True, "indent": 1
    })
    f_value_alt = workbook.add_format({
        "font_name": ZH, "font_size": 10,
        "bg_color": "#F7F9FC", "border": 1, "border_color": "#D9D9D9",
        "valign": "top", "text_wrap": True, "indent": 1
    })
    f_code = workbook.add_format({
        "font_name": MONO, "font_size": 10,
        "bg_color": "#F5F5F5", "border": 1, "border_color": "#D9D9D9",
        "valign": "top", "text_wrap": True, "indent": 1, "font_color": "#1F4E79"
    })
    f_code_alt = workbook.add_format({
        "font_name": MONO, "font_size": 10,
        "bg_color": "#FAFAFA", "border": 1, "border_color": "#D9D9D9",
        "valign": "top", "text_wrap": True, "indent": 1, "font_color": "#1F4E79"
    })
    f_th = workbook.add_format({
        "bold": True, "font_name": ZH, "font_size": 10,
        "bg_color": "#4F81BD", "font_color": "white",
        "border": 1, "border_color": "#2E75B6",
        "align": "center", "valign": "vcenter", "text_wrap": True
    })
    f_good = workbook.add_format({
        "bold": True, "font_name": ZH, "font_size": 10,
        "bg_color": "#C6EFCE", "font_color": "#006100",
        "border": 1, "border_color": "#A9D18E",
        "align": "center", "valign": "vcenter"
    })
    f_mid = workbook.add_format({
        "bold": True, "font_name": ZH, "font_size": 10,
        "bg_color": "#FFEB9C", "font_color": "#9C6500",
        "border": 1, "border_color": "#FFD966",
        "align": "center", "valign": "vcenter"
    })
    f_bad = workbook.add_format({
        "bold": True, "font_name": ZH, "font_size": 10,
        "bg_color": "#FFC7CE", "font_color": "#9C0006",
        "border": 1, "border_color": "#F4B084",
        "align": "center", "valign": "vcenter"
    })
    f_note = workbook.add_format({
        "italic": True, "font_name": ZH, "font_size": 9,
        "font_color": "#7F7F7F", "border": 1, "border_color": "#E0E0E0",
        "valign": "top", "text_wrap": True, "bg_color": "#FAFAFA", "indent": 1
    })
    f_tip = workbook.add_format({
        "italic": True, "font_name": ZH, "font_size": 10,
        "font_color": "#385723", "border": 1, "border_color": "#A9D18E",
        "valign": "top", "text_wrap": True, "bg_color": "#E2EFDA", "indent": 1
    })
    f_subtitle = workbook.add_format({
        "bold": True, "italic": True, "font_size": 12,
        "font_name": ZH,
        "font_color": "#1F4E79", "bg_color": "#DEEBF7",
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": "#9DC3E6"
    })
    # 原生 Excel 表的列格式
    f_t_int = workbook.add_format({
        "font_name": MONO, "font_size": 10,
        "align": "center", "valign": "vcenter", "border": 1
    })
    f_t_num = workbook.add_format({
        "font_name": MONO, "font_size": 10,
        "align": "center", "valign": "vcenter", "border": 1,
        "num_format": "0.0000", "font_color": "#1F4E79"
    })
    f_t_use = workbook.add_format({
        "font_name": ZH, "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1, "indent": 1
    })

    # 斑马纹辅助
    def L(i): return f_label_alt if i % 2 == 1 else f_label
    def V(i): return f_value_alt if i % 2 == 1 else f_value
    def C(i): return f_code_alt if i % 2 == 1 else f_code

    r = 1  # 行游标 (0-indexed)

    # ============== 标题区 ==============
    ws.set_row(r, 50)
    ws.merge_range(r, 1, r, 3, "  GRR  (Gauge R&R)  报  表  完  整  解  读", f_h1)
    r += 1
    ws.set_row(r, 26)
    ws.merge_range(r, 1, r, 3,
                   "  基于 AIAG MSA 4th · 完全复刻 GRR_Report.py + 新增本 Help sheet",
                   f_subtitle)
    r += 2

    # ============== 目录 ==============
    ws.set_row(r, 30)
    ws.merge_range(r, 1, r, 3, "  📖  目  录  (Table of Contents)", f_h2)
    r += 1
    toc = [
        "1.  GRR 与 MSA 概述",
        "2.  方差分量模型 (Variance Components)",
        "3.  两种算法对比：ANOVA 法 vs Xbar-R 法",
        "4.  AIAG 验收准则与颜色编码",
        "5.  AIAG K1 / K2 / K3 系数表",
        "6.  所有公式汇总",
        "7.  Excel 报表结构（逐 sheet 解读）",
        "8.  图表解读",
        "9.  输入数据格式约定",
        "10. 处理流程与异常处理",
        "11. 解读示例",
    ]
    for i, item in enumerate(toc):
        ws.set_row(r, 22)
        ws.write(r, 1, f"  {i+1:>2}.", L(i))
        ws.merge_range(r, 2, r, 3, item, V(i))
        r += 1
    r += 1

    # ============== Section 1: GRR 与 MSA 概述 ==============
    ws.set_row(r, 30)
    ws.merge_range(r, 1, r, 3, "  1.  GRR 与 MSA 概述", f_h2)
    r += 1
    sec1 = [
        ("GRR 全称",
         "Gauge Repeatability & Reproducibility（量具的重复性与再现性）。\n"
         "是 MSA (Measurement System Analysis, 测量系统分析) 中的核心子模块。"),
        ("Repeatability (EV)",
         "同一操作员、同一量具、同一零件在相同条件下多次测量的变差。\n"
         "反映设备自身的波动，即 Equipment Variation。"),
        ("Reproducibility (AV)",
         "不同操作员使用同一量具测量同一零件的变差。\n"
         "反映操作员之间的差异，即 Appraiser Variation。"),
        ("适用场景",
         "本工具针对频响类 (Frequency Response) 测量数据：\n"
         "  · X 轴：频率 Hz（10 ~ 20000 Hz，对数刻度）\n"
         "  · Y 轴：幅值 dB（线性刻度）\n"
         "  · 典型应用：扬声器、麦克风、滤波器、振动传感器等"),
        ("判定原则",
         "若 GRR 占总变差比例过大（红区），说明测量系统本身的波动 > 零件本身的波动，\n"
         "应先改进量具 / 培训操作员，再用此系统做产品判定；否则测量结果不可信。"),
    ]
    for i, (k, v) in enumerate(sec1):
        ws.set_row(r, 50)
        ws.write(r, 1, k, L(i))
        ws.merge_range(r, 2, r, 3, v, V(i))
        r += 1
    r += 1

    # ============== Section 2: 方差分量模型 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  2. 方差分量模型 (Variance Components)", f_h2)
    r += 1
    ws.set_row(r, 24)
    ws.write(r, 1, "分量 (Component)", f_th)
    ws.write(r, 2, "含义 / 物理意义", f_th)
    ws.write(r, 3, "计算来源", f_th)
    r += 1
    sec2 = [
        ("EV (Equipment Variation)",
         "设备变差 / 重复性 (Repeatability)。\n"
         "同一零件在相同条件下多次测量值的标准差。",
         "ANOVA: √MS_Repeat\nXR: R̄ · K1"),
        ("AV (Appraiser Variation)",
         "操作员变差 / 再现性 (Reproducibility)。\n"
         "不同操作员测量均值差异导致的标准差。",
         "ANOVA: √(Var_Operator + Var_Interaction)\n"
         "XR: √[(xdiff · K2)² − EV² / (p · r)]"),
        ("OP (Operator only)",
         "[ANOVA 专属] 剥离交互作用后的纯操作员效应。",
         "ANOVA: √max(MS_Operator − MS_pooled, 0)"),
        ("OPI (Operator × Part)",
         "[ANOVA 专属] 操作员 × 零件的交互作用。\n"
         "不同操作员对不同零件的测量行为差异。",
         "ANOVA: √max(MS_Interaction − MS_Repeat, 0) / √r"),
        ("PV (Part Variation)",
         "零件变差 (Part Variation)。\n"
         "不同零件之间测量均值差异的标准差。",
         "ANOVA: √Var_Part\nXR: Rp · K3"),
        ("GRR (Gauge R&R)",
         "测量系统综合变差：GRR = √(EV² + AV²)。",
         "σ_GRR"),
        ("Total Variation",
         "总变差（测量系统 + 零件）：Total = √(GRR² + PV²)。",
         "σ_Total"),
    ]
    for i, (code, desc, calc) in enumerate(sec2):
        ws.set_row(r, 42)
        ws.write(r, 1, code, L(i))
        ws.write(r, 2, desc, V(i))
        ws.write(r, 3, calc, C(i))
        r += 1
    r += 1

    # ============== Section 3: 两种算法对比 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  3. 两种算法对比：ANOVA 法 vs Xbar-R 法", f_h2)
    r += 1
    ws.set_row(r, 24)
    ws.write(r, 1, "对比项", f_th)
    ws.write(r, 2, "ANOVA 法", f_th)
    ws.write(r, 3, "Xbar-R 法 (AIAG MSA 4th)", f_th)
    r += 1
    sec3 = [
        ("核心思路",
         "将总平方和 SS 分解到各因子 (Part / Operator / OP×Part / Repeat)，\n"
         "用期望均方估计方差分量。",
         "基于每个 (Part, Operator) 单元的极差 R 与各因子均值差。"),
        ("能否区分 OP 与 OPI",
         "✅ 能分离出纯操作员效应 (OP) 与交互作用 (OPI)",
         "❌ 不能，AV = OP + OPI 合并估计"),
        ("Trial 数范围",
         "理论上无上限（本工具支持 2–10）",
         "需查 K1 表，目前支持 2–10"),
        ("Operator 数范围",
         "≥ 2（本工具支持 2–10）",
         "≥ 2，需查 K2 表"),
        ("Part 数范围",
         "≥ 2（本工具支持 2–10）",
         "≥ 2，需查 K3 表"),
        ("显著性检验",
         "对 Operator × Part 做 F 检验 (α=0.05)，\n"
         "不显著时合并交互与重复项 (pooled)。",
         "无 F 检验，直接估计。"),
        ("方差估计保守度",
         "略激进（误差更小但可能受异常点影响）",
         "更保守（误差更大但更稳定）"),
        ("结果在报告中的位置",
         "GRR Results sheet 中 ANOVA_* 前缀的所有列",
         "GRR Results sheet 中 XR_* 前缀的所有列"),
    ]
    for i, (k, a, x) in enumerate(sec3):
        ws.set_row(r, 40)
        ws.write(r, 1, k, L(i))
        ws.write(r, 2, a, V(i))
        ws.write(r, 3, x, V(i))
        r += 1
    r += 1

    # ============== Section 4: 验收准则 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  4. AIAG 验收准则与颜色编码", f_h2)
    r += 1
    ws.set_row(r, 24)
    ws.write(r, 1, "指标 (Indicator)", f_th)
    ws.write(r, 2, "AIAG 阈值 (Ok / Marginal / 不通过)", f_th)
    ws.write(r, 3, "判定与建议", f_th)
    r += 1
    sec4 = [
        ("%Tolerance (GRR)\n占公差带宽的比例",
         f"🟢 OK: ≤ {PCT_TOL_OK}%\n"
         f"🟠 Marginal: {PCT_TOL_OK}% ~ {PCT_TOL_MARGINAL}%\n"
         f"🔴 不通过: > {PCT_TOL_MARGINAL}%",
         "🟢 合格 → 系统可用于过程监控\n"
         "🟠 边缘 → 可有条件接受，建议改进\n"
         "🔴 不合格 → 必须改进测量系统"),
        ("%Study Variation (GRR)\n占总变差的比例",
         f"🟢 OK: ≤ {PCT_STUDY_OK}%\n"
         f"🟠 Marginal: {PCT_STUDY_OK}% ~ {PCT_STUDY_MARGINAL}%\n"
         f"🔴 不通过: > {PCT_STUDY_MARGINAL}%",
         "🟢 合格 → 系统适合 SPC 控制图\n"
         "🟠 边缘 → 评估能否用于 SPC\n"
         "🔴 不合格 → 系统不适用于 SPC"),
        ("NDC\n(Number of Distinct Categories)",
         f"🟢 OK: ≥ {NDC_OK}\n"
         f"🟠 Marginal: {NDC_MARGINAL} ~ {NDC_OK - 1}\n"
         f"🔴 不通过: < {NDC_MARGINAL}",
         "🟢 ≥5 → 系统分辨力充足\n"
         "🟠 2~4 → 边缘，建议增加\n"
         "🔴 <2 → 系统无法区分零件差异"),
        ("%Tolerance (PV)\n零件变差占公差比例",
         "参考: ≥ 80% 通常较好\n"
         "(无强制 OK/Marginal 标准)",
         "若 PV 太小，说明零件本身差异不足，\n"
         "GRR 评估将失去意义。"),
    ]
    for i, (ind, th, act) in enumerate(sec4):
        ws.set_row(r, 60)
        ws.write(r, 1, ind, L(i))
        ws.write(r, 2, th, V(i))
        ws.write(r, 3, act, V(i))
        r += 1

    # 颜色样例行
    ws.set_row(r, 22)
    ws.write(r, 1, "颜色样例", f_label)
    ws.write(r, 2, "🟢 OK / 合格", f_good)
    ws.write(r, 3, "🟠 Marginal / 边缘", f_mid)
    r += 1
    ws.set_row(r, 22)
    ws.write(r, 1, "", f_label)
    ws.write(r, 2, "🔴 不通过 / 不合格", f_bad)
    ws.write(r, 3, "", f_value)
    r += 2

    # ============== Section 5: K 系数表 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  5. AIAG K1 / K2 / K3 系数表 (Xbar-R 法专用)", f_h2)
    r += 1
    ws.set_row(r, 36)
    ws.merge_range(r, 1, r, 3,
                   "K1 / K2 / K3 是 AIAG MSA 4th 手册中给定的常数，"
                   "用于 Xbar-R 法把 Range / 均值差换算成 σ。",
                   f_tip)
    r += 1

    # K1 表 (原生 Excel 表)
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  K1 表 — 基于 Trial (重复次数) n", f_h3)
    r += 1
    k1_data = [[n, K1_TABLE[n], "EV = R̄ · K1"] for n in range(2, 11)]
    ws.add_table(r, 1, r + 9, 3, {
        "name": "K1Table",
        "header_row": True,
        "style": "Table Style Medium 2",
        "columns": [
            {"header": "Trial (n)", "format": f_t_int},
            {"header": "K1 (d2)",   "format": f_t_num},
            {"header": "用途",       "format": f_t_use},
        ],
        "data": k1_data,
    })
    r += 10
    r += 1

    # K2 表 (原生 Excel 表)
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  K2 表 — 基于 Operator (操作员) 数 o", f_h3)
    r += 1
    k2_data = [[o, K2_TABLE[o], "AV 项 = (xdiff · K2)² − EV² / (p · r)"] for o in range(2, 11)]
    ws.add_table(r, 1, r + 9, 3, {
        "name": "K2Table",
        "header_row": True,
        "style": "Table Style Medium 2",
        "columns": [
            {"header": "Operator (o)", "format": f_t_int},
            {"header": "K2",            "format": f_t_num},
            {"header": "用途",          "format": f_t_use},
        ],
        "data": k2_data,
    })
    r += 10
    r += 1

    # K3 表 (原生 Excel 表)
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  K3 表 — 基于 Part (零件) 数 p", f_h3)
    r += 1
    k3_data = [[p, K3_TABLE[p], "PV = Rp · K3"] for p in range(2, 11)]
    ws.add_table(r, 1, r + 9, 3, {
        "name": "K3Table",
        "header_row": True,
        "style": "Table Style Medium 2",
        "columns": [
            {"header": "Part (p)", "format": f_t_int},
            {"header": "K3 (d2*)", "format": f_t_num},
            {"header": "用途",      "format": f_t_use},
        ],
        "data": k3_data,
    })
    r += 10
    r += 1

    # ============== Section 6: 公式汇总 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  6. 所有公式汇总", f_h2)
    r += 1
    sec6 = [
        ("%Tolerance (X)",
         "(6 · σ_X) / (USL − LSL) × 100%",
         "6σ 覆盖 99.73% 正态分布；除以 Tolerance 得到占公差的比例。"),
        ("%Study Variation (X)",
         "σ_X / σ_Total × 100%",
         "占测量系统总变差的比例，用于评估 GRR 的相对大小。"),
        ("NDC",
         "floor(1.41 · σ_PV / σ_GRR)，下限 1",
         "1.41 ≈ √2，AIAG 经验因子；表示系统能可靠区分的零件类别数。"),
        ("ANOVA F 检验 (交互作用)",
         "F = MS_Interaction / MS_Repeat\np = 1 − F.cdf(F, df_interaction, df_repeat)",
         "p ≤ 0.05 时认为交互作用显著，需分别估计 OP 与 OPI。"),
        ("方差分量 (ANOVA, 交互显著)",
         "Var_EV = MS_Repeat\n"
         "Var_OPI = max(MS_Interaction − MS_Repeat, 0) / r\n"
         "Var_OP  = max(MS_Operator − MS_Interaction, 0) / (p · r)\n"
         "Var_PV  = max(MS_Part − MS_Interaction, 0) / (o · r)",
         "对负方差取 max(·, 0) 是 AIAG 推荐做法。"),
        ("方差分量 (ANOVA, 交互不显著)",
         "pooled_MS = (SS_Interaction + SS_Repeat) / (df_Interaction + df_Repeat)\n"
         "Var_OP = max(MS_Operator − pooled_MS, 0) / (p · r)\n"
         "Var_PV = max(MS_Part − pooled_MS, 0) / (o · r)",
         "合并后只输出 1 个交互自由度，统计功效更稳健。"),
        ("Xbar-R 标准差估计",
         "EV = R̄ · K1\n"
         "AV = √max((xdiff · K2)² − EV² / (p · r), 0)\n"
         "PV = Rp · K3",
         "R̄ = 各单元极差均值；xdiff = 操作员均值极差；Rp = 零件均值极差。"),
        ("综合与汇总",
         "GRR   = √(EV² + AV²)\n"
         "Total = √(GRR² + PV²)",
         "最终所有指标都基于这两个综合值计算。"),
    ]
    for i, (k, formula, note) in enumerate(sec6):
        ws.set_row(r, 52)
        ws.write(r, 1, k, L(i))
        ws.write(r, 2, formula, C(i))
        ws.write(r, 3, note, V(i))
        r += 1
    r += 1

    # ============== Section 7: 报表结构 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  7. Excel 报表结构（逐 sheet 解读）", f_h2)
    r += 1
    ws.set_row(r, 24)
    ws.write(r, 1, "Sheet 名", f_th)
    ws.write(r, 2, "作用 / 内容", f_th)
    ws.write(r, 3, "主要列 / 关键内容", f_th)
    r += 1
    sec7 = [
        ("Help 使用说明",
         "本 sheet：完整解读 GRR 报告与方法学。",
         "11 个章节，详见上方目录。"),
        ("Input Data",
         "基础信息、阈值、原始宽表 + 摘要表（带颜色判定）。",
         "Part/Trial/Operator 数、Sigma、阈值、X/Y Scale 设置；\n"
         "摘要表三块：%Tol (GRR)、%Study Var (GRR)、NDC，\n"
         "分别用绿/橙/红条件格式标识 Xbar-R 与 ANOVA 的最大/最差值。"),
        ("Normalized Data",
         "解析后的长表（Long format）。",
         "列：Part, Operator, Trial, X, Y, Lower Limit, Upper Limit, CurveID\n"
         "CurveID 形如 'A_T1_P1'，代表一次测量曲线。"),
        ("GRR Results",
         "每个频点的完整 ANOVA + Xbar-R 计算结果（宽表）。",
         "≈ 40 列：X、上下限、各方差分量 (SD_*)、\n"
         "%Tolerance (EV/AV/PV/OP/OPI/GRR)、%Study Variation、\n"
         "NDC、交互作用 p_value、interaction_significant、K1/K2/K3/R̄ 等。"),
        ("GRR by Frequency",
         "GRR Results 的精简版：每个频点的关键 12 列。",
         "X / 上下限 / Has Tolerance Info / ANOVA 三件套 / XR 三件套。"),
        ("Charts",
         "三大指标 vs Hz 的汇总折线图（散点-平滑）。",
         "三张图：%Tolerance、%Study Variation、NDC，\n"
         "分别叠加 10%/20%（或 NDC=5/2）绿色/金色限值线。"),
        ("Raw Curves",
         "原始测量曲线矩阵 + 全部 / 按 Operator / 按 Part 三组散点图。",
         "X, Lower Limit, Upper Limit + 各 CurveID 一列；\n"
         "双列布局的图表按 (idx // 2, idx % 2) 自动排版。"),
        ("MSA 4 Charts",
         "AIAG MSA 4 Charts 的频响扩展版。",
         "四张图（2×2 排布）：\n"
         "  · Xbar-R Study Variation % of Tolerance\n"
         "  · ANOVA Study Variation % of Tolerance (含 OP/OPI)\n"
         "  · Xbar-R %Study Variation (含限值线)\n"
         "  · ANOVA %Study Variation (含限值线)"),
    ]
    for i, (sn, role, detail) in enumerate(sec7):
        ws.set_row(r, 56)
        ws.write(r, 1, sn, L(i))
        ws.write(r, 2, role, V(i))
        ws.write(r, 3, detail, V(i))
        r += 1
    r += 1

    # ============== Section 8: 图表解读 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  8. 图表解读", f_h2)
    r += 1
    ws.set_row(r, 24)
    ws.write(r, 1, "图表", f_th)
    ws.write(r, 2, "看什么", f_th)
    ws.write(r, 3, "如何判读", f_th)
    r += 1
    sec8 = [
        ("%Tolerance (GRR) by Frequency",
         "ANOVA 与 XR 两条曲线，及 10% / 20% 限值线。",
         "曲线整体在 10% 以下 → 系统对公差而言良好；\n"
         "10~20% → 边缘，部分频段需关注；\n"
         ">20% → 系统本身波动过大。"),
        ("%Study Variation (GRR) by Frequency",
         "ANOVA 与 XR 两条曲线，及 10% / 20% 限值线。",
         "数值越小代表 GRR 占比越小，系统越适合 SPC；\n"
         "建议所有频段 ≤ 30%，绝大部分 ≤ 10%。"),
        ("NDC by Frequency",
         "ANOVA 与 XR 两条曲线，及 NDC=5 / NDC=2 限值线。",
         "≥5 → 系统分辨力充足；\n"
         "2~4 → 边缘，常提示测量噪声偏大；\n"
         "<2 → 系统无法区分零件。"),
        ("Raw Curves — All / By Operator / By Part",
         "所有原始测量曲线 + 上下限红线。",
         "观察各曲线重合度（曲线越集中 → 系统越好）；\n"
         "操作员间一致性、按零件的区分能力都可直接读出。"),
        ("MSA 4 Charts — Study Variation % of Tolerance",
         "TVTOL(基准100%) + PV/EV/AV/GRR / OP/OPI 6~7 条曲线。",
         "PV 是否主导 (≥80%)；GRR 占比；\n"
         "OP 与 OPI 的大小关系 — OPI 偏大说明操作员对不同零件反应不一致。"),
        ("MSA 4 Charts — %Study Variation",
         "PV / EV / AV / OP / OPI / GRR 各分量随频率的占比。",
         "EV 主导 → 设备问题（重复性差）；\n"
         "AV/OP 主导 → 操作员差异大；\n"
         "PV 主导 → 零件差异，测量系统合格。"),
    ]
    for i, (nm, look, judge) in enumerate(sec8):
        ws.set_row(r, 50)
        ws.write(r, 1, nm, L(i))
        ws.write(r, 2, look, V(i))
        ws.write(r, 3, judge, V(i))
        r += 1
    r += 1

    # ============== Section 9: 输入数据格式 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  9. 输入数据格式约定", f_h2)
    r += 1
    sec9 = [
        ("Excel 宽表布局",
         "· 第 1、2 行：可放任意描述性文字（程序不读）\n"
         "· 第 3 行：表头 (header_row = df.iloc[2])\n"
         "· 第 4 行起：数据行",
         "本工具强制 row 3 为表头，row 4 起为数据。"),
        ("必需列",
         "Hz 列（大小写不敏感，必须存在）",
         "该列对应频率轴 X，会被转成数值；缺失时直接报错。"),
        ("可选列",
         "Lower Limit / Upper Limit (列名含 lower/upper + limit)\n"
         "若不存在则 %Tolerance 与 GRR 同列的 %Tol_* 全部为 NaN。",
         "建议提供 — 没有公差就只好看 %Study Variation 与 NDC。"),
        ("测量列命名规则",
         "正则：^([A-Za-z]+)(\\d+)-(\\d+)$",
         "例：A1-1 → Operator='A', Trial=1, Part='P1'；\n"
         "    B2-10 → Operator='B', Trial=2, Part='P10'；\n"
         "不匹配的列会被忽略。"),
        ("数量约束",
         "Part ≥ 2, Operator ≥ 2, Trial ≥ 2",
         "Xbar-R 法额外要求 p/o/r ≤ 10 (K 表范围)；\n"
         "若不满足会抛 ValueError 或跳过该频点。"),
        ("X 轴范围",
         "推荐 10 Hz ~ 20000 Hz（log 轴）",
         "本工具图表 X 轴固定 10~20000，Hz 范围不一致时仅绘图受限。"),
        ("空值处理",
         "NaN 单元格会被跳过，不参与 GRR 计算",
         "若某频点 Part/Operator/Trial < 2 也会被跳过（不报错）。"),
    ]
    for i, (k, v, note) in enumerate(sec9):
        ws.set_row(r, 50)
        ws.write(r, 1, k, L(i))
        ws.write(r, 2, v, V(i))
        ws.write(r, 3, note, f_note)
        r += 1
    r += 1

    # ============== Section 10: 处理流程与异常 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  10. 处理流程与异常处理", f_h2)
    r += 1
    sec10 = [
        ("完整流程",
         "读取 Excel → 解析列 → 转长表 → 逐频点计算 ANOVA+XR → "
         "汇总 → 写 7 + 1 个 sheet → 自动打开。",
         "process_grr_report() 是单点入口；GUI 仅负责收件/出件。"),
        ("频点跳过条件",
         "· 该 Hz 下 Part/Operator/Trial 任一 < 2\n"
         "· ANOVA 或 Xbar-R 计算抛异常 (except: continue)",
         "跳过是静默的 — 不会在报告中列出被跳过的频点。"),
        ("整体校验",
         "process_grr_report() 入口处要求 Part/Operator/Trial ≥ 2，\n"
         "否则抛 ValueError 并在 GUI 中以红色状态显示。",
         "若 GRR Results 为空也会报错 — 通常意味着输入格式有误。"),
        ("NDC 下限处理",
         "PV > 0 时 NDC = max(floor(1.41·PV/GRR), 1)",
         "避免 PV 极小时 NDC=0；偏离 AIAG 原版定义。"),
        ("未实现的扩展",
         "· K1/K2/K3 表目前仅支持 2~10，更大样本会抛错\n"
         "· ANOVA 固定 α=0.05，未在 GUI 中可调",
         "可在源码顶部常量处自行扩展 K 表或改 ANOVA_ALPHA。"),
    ]
    for i, (k, v, note) in enumerate(sec10):
        ws.set_row(r, 60)
        ws.write(r, 1, k, L(i))
        ws.write(r, 2, v, V(i))
        ws.write(r, 3, note, f_note)
        r += 1
    r += 1

    # ============== Section 11: 解读示例 ==============
    ws.set_row(r, 24)
    ws.merge_range(r, 1, r, 3, "  11. 解读示例 (虚拟案例)", f_h2)
    r += 1
    ws.set_row(r, 70)
    ws.merge_range(r, 1, r, 3,
                   "假设某频点的摘要显示：%Tol = 8.5%, %Study Var = 28%, NDC = 3。\n"
                   "  · %Tol ≤ 10% → 🟢 系统对公差而言合格。\n"
                   "  · %Study Var 在 20%~30% → 🟠 偏大，GRR 占比接近 1/3，需要关注。\n"
                   "  · NDC = 3 < 5 → 🟠 系统分辨力处于边缘。\n"
                   "结论：系统基本可用但有改进空间；建议先看 Raw Curves 找到哪些频段 "
                   "GRR 突然飙升，再判断是 EV 还是 AV 主导，针对性优化。",
                   f_tip)
    r += 2

    # ============== Footer ==============
    ws.set_row(r, 22)
    ws.merge_range(r, 1, r, 3,
                   f"  GRR_V2 · {VERSION} · "
                   "基于 AIAG MSA 4th 与 scipy.stats 实现",
                   f_subtitle)
    r += 1
    ws.set_row(r, 18)
    ws.merge_range(r, 1, r, 3,
                   "  本 sheet 旨在帮助阅读报告，所有阈值与公式源自源码顶部常量，"
                   "如需修改请直接编辑 GRR_V2.py 顶部配置区。",
                   f_note)


# =========================================================
# 写报表到 Excel
# =========================================================
def write_report(output_file, df_display, df_long, grr_df, tolerance_lower, tolerance_upper):
    """
    将计算结果写入 Excel，生成多个工作表并绘制相关图表
    """
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        # === V2: 先写入 Help 使用说明 sheet ===
        write_help_sheet(workbook)

        # 一些通用单元格格式
        fmt_title = workbook.add_format({
            "bold": True, "font_size": 12, "align": "center",
            "valign": "vcenter", "border": 1, "bg_color": "#9BBB59"
        })
        fmt_header = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "border": 1, "bg_color": "#4F81BD", "font_color": "white"
        })
        fmt_label = workbook.add_format({
            "border": 1, "bg_color": "#D9EAF7"
        })
        fmt_value = workbook.add_format({
            "border": 1, "bg_color": "#B8CCE4"
        })
        fmt_num = workbook.add_format({
            "border": 1, "num_format": "0.0000"
        })
        fmt_pct = workbook.add_format({
            "border": 1, "num_format": "0.0%"
        })
        fmt_center = workbook.add_format({
            "border": 1, "align": "center", "valign": "vcenter"
        })
        fmt_good = workbook.add_format({
            "border": 1, "bg_color": "#C6EFCE", "font_color": "#006100"
        })
        fmt_mid = workbook.add_format({
            "border": 1, "bg_color": "#FFEB9C", "font_color": "#9C6500"
        })
        fmt_bad = workbook.add_format({
            "border": 1, "bg_color": "#FFC7CE", "font_color": "#9C0006"
        })

        # =====================================================
        # Sheet 1: Input Data (输入信息与宽表数据)
        # =====================================================
        ws1 = workbook.add_worksheet("Input Data")
        writer.sheets["Input Data"] = ws1

        parts = df_long["Part"].nunique()
        trials = df_long["Trial"].nunique()
        operators = df_long["Operator"].nunique()

        # 填写基本信息
        ws1.write("A1", "Number of parts (2-10) =", fmt_label)
        ws1.write("B1", parts, fmt_value)
        ws1.write("A2", "Number of trials (2-10) =", fmt_label)
        ws1.write("B2", trials, fmt_value)
        ws1.write("A3", "Number of operators (2-10) =", fmt_label)
        ws1.write("B3", operators, fmt_value)
        ws1.write("A4", "Number of Sigma", fmt_label)
        ws1.write("B4", NUMBER_OF_SIGMA, fmt_value)
        ws1.write("A5", "ANOVA α Limit (Significant: p-value ≤ α)", fmt_label)
        ws1.write("B5", ANOVA_ALPHA, fmt_value)

        ws1.write("A6", "%Tolerance Ok Limit", fmt_label)
        ws1.write("B6", f"{PCT_TOL_OK}%", fmt_value)
        ws1.write("A7", "%Tolerance Marginal Limit", fmt_label)
        ws1.write("B7", f"{PCT_TOL_MARGINAL}%", fmt_value)
        ws1.write("A8", "%Study Variation Ok Limit", fmt_label)
        ws1.write("B8", f"{PCT_STUDY_OK}%", fmt_value)
        ws1.write("A9", "%Study Variation Marginal Limit", fmt_label)
        ws1.write("B9", f"{PCT_STUDY_MARGINAL}%", fmt_value)
        ws1.write("A10", "NDC Ok Limit", fmt_label)
        ws1.write("B10", NDC_OK, fmt_value)
        ws1.write("A11", "NDC Marginal Limit", fmt_label)
        ws1.write("B11", NDC_MARGINAL, fmt_value)

        ws1.write("A14", "Lower X Scale", fmt_label)
        ws1.write("B14", float(df_long["X"].min()), fmt_value)
        ws1.write("A15", "Upper X Scale", fmt_label)
        ws1.write("B15", float(df_long["X"].max()), fmt_value)
        ws1.write("A16", "X Unit", fmt_label)
        ws1.write("B16", X_UNIT, fmt_value)
        ws1.write("A17", "X Scale Name", fmt_label)
        ws1.write("B17", X_SCALE_NAME, fmt_value)
        ws1.write("A18", "X Scale Type", fmt_label)
        ws1.write("B18", X_SCALE_TYPE, fmt_value)
        ws1.write("A19", "Lower Y Scale", fmt_label)
        if pd.notna(tolerance_lower):
            ws1.write("B19", tolerance_lower, fmt_value)
        else:
            ws1.write_blank("B19", None, fmt_value)
        ws1.write("A20", "Upper Y Scale", fmt_label)
        if pd.notna(tolerance_upper):
            ws1.write("B20", tolerance_upper, fmt_value)
        else:
            ws1.write_blank("B20", None, fmt_value)
        ws1.write("A21", "Y Unit", fmt_label)
        ws1.write("B21", Y_UNIT, fmt_value)
        ws1.write("A22", "Y Scale Name", fmt_label)
        ws1.write("B22", Y_SCALE_NAME, fmt_value)
        ws1.write("A23", "Y Scale Type", fmt_label)
        ws1.write("B23", Y_SCALE_TYPE, fmt_value)

        # 计算最大 %Tol、%Study、最小 NDC (用于摘要)
        xr_tol = grr_df["XR_%Tolerance_GRR"].max() / 100 if not grr_df.empty and grr_df["XR_%Tolerance_GRR"].notna().any() else np.nan
        anova_tol = grr_df["ANOVA_%Tolerance_GRR"].max() / 100 if not grr_df.empty and grr_df["ANOVA_%Tolerance_GRR"].notna().any() else np.nan

        xr_sv = grr_df["XR_%StudyVariation_GRR"].max() / 100 if not grr_df.empty and grr_df["XR_%StudyVariation_GRR"].notna().any() else np.nan
        anova_sv = grr_df["ANOVA_%StudyVariation_GRR"].max() / 100 if not grr_df.empty and grr_df["ANOVA_%StudyVariation_GRR"].notna().any() else np.nan

        xr_ndc = grr_df["XR_NDC"].min() if not grr_df.empty and grr_df["XR_NDC"].notna().any() else np.nan
        anova_ndc = grr_df["ANOVA_NDC"].min() if not grr_df.empty and grr_df["ANOVA_NDC"].notna().any() else np.nan

        # 摘要表
        ws1.write("A25", "...", fmt_header)
        ws1.write("B25", "%Tolerance (GRR)", fmt_header)
        ws1.write("A26", "Xbar & Range", fmt_label)
        ws1.write("A27", "ANOVA", fmt_label)
        if pd.notna(xr_tol):
            ws1.write_number("B26", xr_tol, fmt_pct)
        else:
            ws1.write_blank("B26", None, fmt_center)
        if pd.notna(anova_tol):
            ws1.write_number("B27", anova_tol, fmt_pct)
        else:
            ws1.write_blank("B27", None, fmt_center)

        ws1.write("A29", "...", fmt_header)
        ws1.write("B29", "%Study Variation (GRR)", fmt_header)
        ws1.write("A30", "Xbar & Range", fmt_label)
        ws1.write("A31", "ANOVA", fmt_label)
        if pd.notna(xr_sv):
            ws1.write_number("B30", xr_sv, fmt_pct)
        else:
            ws1.write_blank("B30", None, fmt_center)
        if pd.notna(anova_sv):
            ws1.write_number("B31", anova_sv, fmt_pct)
        else:
            ws1.write_blank("B31", None, fmt_center)

        ws1.write("A33", "...", fmt_header)
        ws1.write("B33", "NDC", fmt_header)
        ws1.write("A34", "Xbar & Range", fmt_label)
        ws1.write("A35", "ANOVA", fmt_label)
        if pd.notna(xr_ndc):
            ws1.write_number("B34", float(xr_ndc), fmt_center)
        else:
            ws1.write_blank("B34", None, fmt_center)
        if pd.notna(anova_ndc):
            ws1.write_number("B35", float(anova_ndc), fmt_center)
        else:
            ws1.write_blank("B35", None, fmt_center)

        # 条件格式 (%Tol)
        ws1.conditional_format("B26:B27", {
            "type": "cell", "criteria": "<=", "value": PCT_TOL_OK / 100,
            "format": fmt_good
        })
        ws1.conditional_format("B26:B27", {
            "type": "cell", "criteria": "between",
            "minimum": PCT_TOL_OK / 100, "maximum": PCT_TOL_MARGINAL / 100,
            "format": fmt_mid
        })
        ws1.conditional_format("B26:B27", {
            "type": "cell", "criteria": ">", "value": PCT_TOL_MARGINAL / 100,
            "format": fmt_bad
        })

        # 条件格式 (%Study Variation)
        ws1.conditional_format("B30:B31", {
            "type": "cell", "criteria": "<=", "value": PCT_STUDY_OK / 100,
            "format": fmt_good
        })
        ws1.conditional_format("B30:B31", {
            "type": "cell", "criteria": "between",
            "minimum": PCT_STUDY_OK / 100, "maximum": PCT_STUDY_MARGINAL / 100,
            "format": fmt_mid
        })
        ws1.conditional_format("B30:B31", {
            "type": "cell", "criteria": ">", "value": PCT_STUDY_MARGINAL / 100,
            "format": fmt_bad
        })

        # 条件格式 (NDC)
        ws1.conditional_format("B34:B35", {
            "type": "cell", "criteria": ">=", "value": NDC_OK,
            "format": fmt_good
        })
        ws1.conditional_format("B34:B35", {
            "type": "cell", "criteria": "between",
            "minimum": NDC_MARGINAL, "maximum": NDC_OK - 0.0001,
            "format": fmt_mid
        })
        ws1.conditional_format("B34:B35", {
            "type": "cell", "criteria": "<", "value": NDC_MARGINAL,
            "format": fmt_bad
        })

        # 把宽表数据写入 Input Data
        start_row = 0
        start_col = 2
        for c_idx, col_name in enumerate(df_display.columns):
            ws1.write(start_row, start_col + c_idx, str(col_name), fmt_header)

        for r_idx in range(len(df_display)):
            for c_idx, col_name in enumerate(df_display.columns):
                val = df_display.iloc[r_idx, c_idx]
                if pd.isna(val):
                    ws1.write_blank(start_row + 1 + r_idx, start_col + c_idx, None, fmt_center)
                elif c_idx == 0:
                    ws1.write_number(start_row + 1 + r_idx, start_col + c_idx, float(val), fmt_num)
                else:
                    try:
                        ws1.write_number(start_row + 1 + r_idx, start_col + c_idx, float(val), fmt_num)
                    except Exception:
                        ws1.write(start_row + 1 + r_idx, start_col + c_idx, str(val), fmt_center)

        ws1.set_column("A:A", 32)
        ws1.set_column("B:B", 18)
        ws1.set_column(start_col, start_col + len(df_display.columns), 10)

        # =====================================================
        # Sheet 2: Normalized Data (长表数据)
        # =====================================================
        df_long.to_excel(writer, sheet_name="Normalized Data", index=False)
        ws2 = writer.sheets["Normalized Data"]
        ws2.set_column(0, len(df_long.columns) - 1, 14)

        # =====================================================
        # Sheet 3: GRR Results (完整计算结果)
        # =====================================================
        grr_df.to_excel(writer, sheet_name="GRR Results", index=False)
        ws3 = writer.sheets["GRR Results"]
        ws3.set_column(0, len(grr_df.columns) - 1, 18)

        # =====================================================
        # Sheet 4: GRR by Frequency (每个频点的关键结果)
        # =====================================================
        grr_by_freq_df = build_grr_by_frequency_sheet(grr_df)
        grr_by_freq_df.to_excel(writer, sheet_name="GRR by Frequency", index=False)
        ws4 = writer.sheets["GRR by Frequency"]
        ws4.set_column(0, len(grr_by_freq_df.columns) - 1, 20)

        # =====================================================
        # Sheet 5: Charts (三大指标折线图)
        # =====================================================
        ws5 = workbook.add_worksheet("Charts")
        writer.sheets["Charts"] = ws5

        chart_cols = [
            "X",
            "ANOVA_%Tolerance_GRR",
            "XR_%Tolerance_GRR",
            "ANOVA_%StudyVariation_GRR",
            "XR_%StudyVariation_GRR",
            "ANOVA_NDC",
            "XR_NDC",
            "Tol_10_Line",
            "Tol_20_Line",
            "Study_10_Line",
            "Study_20_Line",
            "NDC_5_Line",
            "NDC_2_Line"
        ]

        for c, name in enumerate(chart_cols):
            ws5.write(0, c, name, fmt_header)

        has_any_tol_data = (
            (not grr_df.empty) and (
                grr_df["ANOVA_%Tolerance_GRR"].notna().any() or
                grr_df["XR_%Tolerance_GRR"].notna().any()
            )
        )

        # 写入图表数据列
        for r in range(len(grr_df)):
            ws5.write(r + 1, 0, grr_df.iloc[r]["X"], fmt_center)

            for c, name in enumerate(chart_cols[1:7], start=1):
                val = grr_df.iloc[r][name]
                if pd.isna(val):
                    ws5.write_blank(r + 1, c, None, fmt_center)
                else:
                    ws5.write(r + 1, c, val, fmt_center)

            if has_any_tol_data:
                ws5.write(r + 1, 7, PCT_TOL_OK, fmt_center)
                ws5.write(r + 1, 8, PCT_TOL_MARGINAL, fmt_center)
            else:
                ws5.write_blank(r + 1, 7, None, fmt_center)
                ws5.write_blank(r + 1, 8, None, fmt_center)

            ws5.write(r + 1, 9, PCT_STUDY_OK, fmt_center)
            ws5.write(r + 1, 10, PCT_STUDY_MARGINAL, fmt_center)
            ws5.write(r + 1, 11, NDC_OK, fmt_center)
            ws5.write(r + 1, 12, NDC_MARGINAL, fmt_center)

        # 图表 1: %Tolerance GRR
        chart1 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        chart1.add_series({
            'name': 'ANOVA %Tolerance GRR',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 1, len(grr_df), 1],
            'line': {'color': random_color(), 'width': 1.5},
        })
        chart1.add_series({
            'name': 'XR %Tolerance GRR',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 2, len(grr_df), 2],
            'line': {'color': random_color(), 'width': 1.5},
        })
        if has_any_tol_data:
            chart1.add_series({
                'name': '10% Limit',
                'categories': ['Charts', 1, 0, len(grr_df), 0],
                'values': ['Charts', 1, 7, len(grr_df), 7],
                'line': {'color': COLOR_DARK_GREEN, 'width': 1.5},
            })
            chart1.add_series({
                'name': '20% Limit',
                'categories': ['Charts', 1, 0, len(grr_df), 0],
                'values': ['Charts', 1, 8, len(grr_df), 8],
                'line': {'color': COLOR_GOLD, 'width': 1.5},
            })
        set_chart_style(chart1, CHART_TITLE_TOL, "Hz", "%Tolerance")
        ws5.insert_chart('B2', chart1)

        # 图表 2: %Study Variation GRR
        chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        chart2.add_series({
            'name': 'ANOVA %Study Variation GRR',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 3, len(grr_df), 3],
            'line': {'color': random_color(), 'width': 1.5},
        })
        chart2.add_series({
            'name': 'XR %Study Variation GRR',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 4, len(grr_df), 4],
            'line': {'color': random_color(), 'width': 1.5},
        })
        chart2.add_series({
            'name': '10% Limit',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 9, len(grr_df), 9],
            'line': {'color': COLOR_DARK_GREEN, 'width': 1.5},
        })
        chart2.add_series({
            'name': '20% Limit',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 10, len(grr_df), 10],
            'line': {'color': COLOR_GOLD, 'width': 1.5},
        })
        set_chart_style(chart2, CHART_TITLE_STUDY, "Hz", "%Study Variation")
        ws5.insert_chart('B32', chart2)

        # 图表 3: NDC
        chart3 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        chart3.add_series({
            'name': 'ANOVA NDC',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 5, len(grr_df), 5],
            'line': {'color': random_color(), 'width': 1.5},
        })
        chart3.add_series({
            'name': 'XR NDC',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 6, len(grr_df), 6],
            'line': {'color': random_color(), 'width': 1.5},
        })
        chart3.add_series({
            'name': 'NDC=5 Limit',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 11, len(grr_df), 11],
            'line': {'color': COLOR_DARK_GREEN, 'width': 1.5},
        })
        chart3.add_series({
            'name': 'NDC=2 Limit',
            'categories': ['Charts', 1, 0, len(grr_df), 0],
            'values': ['Charts', 1, 12, len(grr_df), 12],
            'line': {'color': COLOR_GOLD, 'width': 1.5},
        })
        set_chart_style(chart3, CHART_TITLE_NDC, "Hz", "NDC")
        ws5.insert_chart('B62', chart3)

        ws5.set_column(0, 12, 18)

        # =====================================================
        # Sheet 6: Raw Curves (原始曲线图)
        # =====================================================
        ws6 = workbook.add_worksheet("Raw Curves")
        writer.sheets["Raw Curves"] = ws6

        raw_matrix_df, curve_meta_df = build_raw_curve_matrix(df_long)

        # 写入矩阵数据
        for c_idx, col_name in enumerate(raw_matrix_df.columns):
            ws6.write(0, c_idx, str(col_name), fmt_header)

        for r_idx in range(len(raw_matrix_df)):
            for c_idx, col_name in enumerate(raw_matrix_df.columns):
                val = raw_matrix_df.iloc[r_idx, c_idx]
                if pd.isna(val):
                    ws6.write_blank(r_idx + 1, c_idx, None, fmt_center)
                else:
                    ws6.write_number(r_idx + 1, c_idx, float(val), fmt_num)

        ws6.set_column(0, len(raw_matrix_df.columns) - 1, 14)

        # 计算 y 轴范围
        y_candidates = pd.concat([
            df_long["Y"],
            df_long["Lower Limit"],
            df_long["Upper Limit"]
        ], axis=0).dropna()

        y_min = y_candidates.min() if not y_candidates.empty else None
        y_max = y_candidates.max() if not y_candidates.empty else None
        y_pad = (y_max - y_min) * 0.05 if y_min is not None and y_max is not None and y_max > y_min else 1
        y_min_plot = y_min - y_pad if y_min is not None else None
        y_max_plot = y_max + y_pad if y_max is not None else None

        curve_columns = list(raw_matrix_df.columns)
        curve_to_colidx = {name: idx for idx, name in enumerate(curve_columns)}

        has_lower_limit_curve = raw_matrix_df["Lower Limit"].notna().any()
        has_upper_limit_curve = raw_matrix_df["Upper Limit"].notna().any()

        chart_index = 0

        # 总的 Raw Curves 图
        chart_all = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        add_raw_limit_series(
            chart_all, 'Raw Curves', len(raw_matrix_df), 0, 1, 2,
            has_lower_limit_curve, has_upper_limit_curve
        )
        for curve_id in curve_columns[3:]:
            add_raw_curve_chart(
                chart=chart_all,
                sheet_name='Raw Curves',
                data_row_count=len(raw_matrix_df),
                x_col_idx=0,
                y_col_idx=curve_to_colidx[curve_id],
                series_name=curve_id,
                color=random_color()
            )
        set_raw_curve_chart_style(chart_all, "All Raw Curves", y_min_plot, y_max_plot)
        row_pos, col_pos = get_chart_position(chart_index)
        ws6.insert_chart(row_pos, col_pos, chart_all)
        chart_index += 1

        # 按操作员绘图
        operators_sorted = sorted(curve_meta_df["Operator"].dropna().unique())
        for op in operators_sorted:
            op_curves = curve_meta_df[curve_meta_df["Operator"] == op]["CurveID"].tolist()
            if not op_curves:
                continue

            chart_op = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
            add_raw_limit_series(
                chart_op, 'Raw Curves', len(raw_matrix_df), 0, 1, 2,
                has_lower_limit_curve, has_upper_limit_curve
            )

            for curve_id in op_curves:
                add_raw_curve_chart(
                    chart=chart_op,
                    sheet_name='Raw Curves',
                    data_row_count=len(raw_matrix_df),
                    x_col_idx=0,
                    y_col_idx=curve_to_colidx[curve_id],
                    series_name=curve_id,
                    color=random_color()
                )
            set_raw_curve_chart_style(chart_op, f"Raw Curves by Operator - {op}", y_min_plot, y_max_plot)
            row_pos, col_pos = get_chart_position(chart_index)
            ws6.insert_chart(row_pos, col_pos, chart_op)
            chart_index += 1

        # 按零件绘图
        parts_sorted = sorted(curve_meta_df["Part"].dropna().unique())
        for part in parts_sorted:
            part_curves = curve_meta_df[curve_meta_df["Part"] == part]["CurveID"].tolist()
            if not part_curves:
                continue

            chart_part = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
            add_raw_limit_series(
                chart_part, 'Raw Curves', len(raw_matrix_df), 0, 1, 2,
                has_lower_limit_curve, has_upper_limit_curve
            )

            for curve_id in part_curves:
                add_raw_curve_chart(
                    chart=chart_part,
                    sheet_name='Raw Curves',
                    data_row_count=len(raw_matrix_df),
                    x_col_idx=0,
                    y_col_idx=curve_to_colidx[curve_id],
                    series_name=curve_id,
                    color=random_color()
                )
            set_raw_curve_chart_style(chart_part, f"Raw Curves by Part - {part}", y_min_plot, y_max_plot)
            row_pos, col_pos = get_chart_position(chart_index)
            ws6.insert_chart(row_pos, col_pos, chart_part)
            chart_index += 1

        # 曲线元信息 (CurveID, Operator, Trial, Part) 展示
        meta_start_col = max(len(raw_matrix_df.columns) + 5, 25)
        for c_idx, col_name in enumerate(curve_meta_df.columns):
            ws6.write(0, meta_start_col + c_idx, str(col_name), fmt_header)

        for r_idx in range(len(curve_meta_df)):
            for c_idx, col_name in enumerate(curve_meta_df.columns):
                val = curve_meta_df.iloc[r_idx, c_idx]
                ws6.write(r_idx + 1, meta_start_col + c_idx, str(val), fmt_center)

        ws6.set_column(meta_start_col, meta_start_col + len(curve_meta_df.columns), 16)

        # =====================================================
        # Sheet 7: MSA 4 Charts（四象限图）
        # =====================================================
        msa4_df = build_msa_four_chart_df(grr_df)
        ws7 = workbook.add_worksheet("MSA 4 Charts")
        writer.sheets["MSA 4 Charts"] = ws7

        # 写表头与数据
        for c_idx, col_name in enumerate(msa4_df.columns):
            ws7.write(0, c_idx, str(col_name), fmt_header)

        for r_idx in range(len(msa4_df)):
            for c_idx, col_name in enumerate(msa4_df.columns):
                val = msa4_df.iloc[r_idx, c_idx]
                if pd.isna(val):
                    ws7.write_blank(r_idx + 1, c_idx, None, fmt_center)
                else:
                    ws7.write_number(r_idx + 1, c_idx, float(val), fmt_center)

        ws7.set_column(0, len(msa4_df.columns) - 1, 16)

        row_count = len(msa4_df)
        cidx = {name: idx for idx, name in enumerate(msa4_df.columns)}

        # 图 1: Xbar & Range - Study Variation as % of Tolerance
        chart_msa_1 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        add_series(chart_msa_1, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_TVTOL"], "TVTOL%", random_color(), 1.5, 'dash')
        add_series(chart_msa_1, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_PVTOL"], "PVTOL%", random_color(), 1.5)
        add_series(chart_msa_1, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_EVTOL"], "EVTOL%", random_color(), 1.5)
        add_series(chart_msa_1, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_AVTOL"], "AVTOL%", random_color(), 1.5)
        add_series(chart_msa_1, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_GRRTOL"], "GRRTOL%", random_color(), 1.5, 'dash')
        add_series(chart_msa_1, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_TOL_OK"], "Ok Limit", COLOR_DARK_GREEN, 1.5)
        add_series(chart_msa_1, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_TOL_MARGINAL"], "Marginal Limit", COLOR_GOLD, 1.5)
        set_msa_same_style(chart_msa_1, "Xbar & Range - Study Variation as a Percentage of Control Limits", "Study Variation % of Tolerance", 0, 120)
        ws7.insert_chart("A2", chart_msa_1)

        # 图 2: ANOVA - Study Variation as % of Tolerance
        chart_msa_2 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_TVTOL"], "TVTOL%", random_color(), 1.5, 'dash')
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_PVTOL"], "PVTOL%", random_color(), 1.5)
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_EVTOL"], "EVTOL%", random_color(), 1.5)
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_AVTOL"], "AVTOL%", random_color(), 1.5)
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_OPTOL"], "OPTOL%", random_color(), 1.5)
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_OPITOL"], "OPITOL%", random_color(), 1.5)
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_GRRTOL"], "GRRTOL%", random_color(), 1.5, 'dash')
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_TOL_OK"], "Ok Limit", COLOR_DARK_GREEN, 1.5)
        add_series(chart_msa_2, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_TOL_MARGINAL"], "Marginal Limit", COLOR_GOLD, 1.5)
        set_msa_same_style(chart_msa_2, "ANOVA - Study Variation as a Percentage of Control Limits", "Study Variation % of Tolerance", 0, 120)
        ws7.insert_chart("H2", chart_msa_2)

        # 图 3: Xbar & Range - %Study Variation (含限值)
        chart_msa_3 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        add_series(chart_msa_3, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_PV"], "PV%", random_color(), 1.5)
        add_series(chart_msa_3, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_EV"], "EV%", random_color(), 1.5)
        add_series(chart_msa_3, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_AV"], "AV%", random_color(), 1.5)
        add_series(chart_msa_3, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_GRR"], "GRR%", random_color(), 1.5, 'dash')
        add_series(chart_msa_3, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_STUDY_OK"], "Ok Limit", COLOR_DARK_GREEN, 1.5)
        add_series(chart_msa_3, "MSA 4 Charts", row_count, cidx["X"], cidx["XR_STUDY_MARGINAL"], "Marginal Limit", COLOR_GOLD, 1.5)
        set_msa_same_style(chart_msa_3, "Xbar & Range - Study Variation & Number of Distinct Categories", "%Study Variation", 0, 100)
        ws7.insert_chart("A26", chart_msa_3)

        # 图 4: ANOVA - %Study Variation (含限值)
        chart_msa_4 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_PV"], "PV%", random_color(), 1.5)
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_EV"], "EV%", random_color(), 1.5)
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_AV"], "AV%", random_color(), 1.5)
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_OP"], "OP%", random_color(), 1.5)
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_OPI"], "OPI%", random_color(), 1.5)
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_GRR"], "GRR%", random_color(), 1.5, 'dash')
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_STUDY_OK"], "Ok Limit", COLOR_DARK_GREEN, 1.5)
        add_series(chart_msa_4, "MSA 4 Charts", row_count, cidx["X"], cidx["ANOVA_STUDY_MARGINAL"], "Marginal Limit", COLOR_GOLD, 1.5)
        set_msa_same_style(chart_msa_4, "ANOVA - Study Variation & Number of Distinct Categories", "%Study Variation", 0, 100)
        ws7.insert_chart("H26", chart_msa_4)


# =========================================================
# 核心处理函数：读取 -> 计算 -> 写报告
# =========================================================
def process_grr_report(input_file, output_file):
    # 读取宽表并转成长表
    df_long, df_display, tolerance_lower, tolerance_upper = load_wide_excel_to_long(input_file)

    parts = df_long["Part"].nunique()
    operators = df_long["Operator"].nunique()
    trials = df_long["Trial"].nunique()

    if parts < 2:
        raise ValueError("Part数量不足，至少需要2个")
    if operators < 2:
        raise ValueError("Operator数量不足，至少需要2个")
    if trials < 2:
        raise ValueError("Trial数量不足，至少需要2次试验")

    # 计算所有频点 GRR
    grr_df = calculate_grr_all_x(df_long)

    if grr_df.empty:
        raise ValueError("GRR计算结果为空，请检查输入数据格式是否完整")

    # 输出 Excel 报表
    write_report(output_file, df_display, df_long, grr_df, tolerance_lower, tolerance_upper)


# =========================================================
# GUI 界面 (V2 美化版)
# =========================================================
class GRRGUI:
    """
    美化后的 GUI：
      · Microsoft YaHei UI 字体 + Fluent 配色
      · 文件路径只显示文件名，完整目录显示在下方小字
      · 卡片式分组 (输入/输出/状态)，点击按钮有 hover 反馈
      · 运行期间禁用按钮 + 进度条 + 状态文字
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("GRR Report Generator · V2.0")
        self.root.geometry("850x650")
        self.root.minsize(750, 650)
        self.root.resizable(True, True)
        self.root.configure(bg="#FFFFFF")

        # Windows DPI 感知，避免高 DPI 屏幕字体发虚
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                from ctypes import windll
                windll.user32.SetProcessDPIAware()
            except Exception:
                pass

        # 字体
        self.font_title = font.Font(family="Microsoft YaHei UI", size=20, weight="bold")
        self.font_subtitle = font.Font(family="Microsoft YaHei UI", size=10)
        self.font_section = font.Font(family="Microsoft YaHei UI", size=12, weight="bold")
        self.font_field = font.Font(family="Microsoft YaHei UI", size=11)
        self.font_btn = font.Font(family="Microsoft YaHei UI", size=11, weight="bold")
        self.font_run = font.Font(family="Microsoft YaHei UI", size=14, weight="bold")
        self.font_status = font.Font(family="Microsoft YaHei UI", size=10)
        self.font_path = font.Font(family="Consolas", size=9)

        # Fluent 配色
        self.brand = "#1F4E79"
        self.brand_hover = "#16466B"
        self.success = "#107C10"
        self.success_hover = "#0B5A0B"
        self.danger = "#D83B01"
        self.muted = "#605E5C"
        self.border = "#E1DFDD"
        self.card_bg = "#F3F2F1"

        # 状态
        self.input_path = None
        self.output_path = None
        self._busy = False
        # trace 抑制标志：浏览按钮设置 StringVar 时避免触发回调
        self._suppress_input_trace = False
        self._suppress_output_trace = False

        # ttk 主题
        self.style = ttk.Style()
        for theme in ("vista", "winnative", "clam", "default"):
            try:
                self.style.theme_use(theme)
                break
            except Exception:
                continue
        try:
            self.style.configure("Big.Horizontal.TProgressbar", thickness=10)
        except Exception:
            pass

        self.build_ui()

    # ------------------------------------------------------------------
    # UI 构建
    # ------------------------------------------------------------------
    def build_ui(self):
        main = tk.Frame(self.root, bg="#FFFFFF", padx=28, pady=22)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)

        # 标题栏
        header = tk.Frame(main, bg="#FFFFFF")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        tk.Label(
            header, text="GRR Report Generator",
            font=self.font_title, bg="#FFFFFF", fg=self.brand
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header, text="V2.0 · 基于 AIAG MSA 4th",
            font=self.font_subtitle, bg="#FFFFFF", fg=self.muted
        ).grid(row=0, column=1, sticky="e")

        # 蓝色分隔线
        sep = tk.Frame(main, height=2, bg=self.brand)
        sep.grid(row=1, column=0, sticky="ew", pady=(16, 18))

        # 输入文件卡片
        self._build_input_card(main, row=2)
        # 输出文件卡片
        self._build_output_card(main, row=3)
        # 数据模板下载卡片
        self._build_template_card(main, row=4)
        # 主操作按钮
        self._build_run_button(main, row=5)
        # 进度与状态
        self._build_status(main, row=6)

    def _build_input_card(self, parent, row):
        card = tk.Frame(
            parent, bg=self.card_bg, padx=18, pady=14,
            highlightthickness=1, highlightbackground=self.border
        )
        card.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        card.columnconfigure(0, weight=1)

        tk.Label(
            card, text="📂  输入文件",
            font=self.font_section, bg=self.card_bg, fg=self.brand, anchor="w"
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        row_frame = tk.Frame(card, bg=self.card_bg)
        row_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        row_frame.columnconfigure(0, weight=1)

        self.input_var = tk.StringVar()
        self.input_entry = tk.Entry(
            row_frame, textvariable=self.input_var,
            font=self.font_field, relief="solid", bd=1,
            bg="#FFFFFF", fg="#323130",
            disabledbackground="#F5F5F5", disabledforeground="#A19F9D",
            highlightthickness=1,
            highlightcolor=self.brand, highlightbackground=self.border
        )
        self.input_entry.grid(row=0, column=0, sticky="ew", ipady=6, padx=(0, 8))
        # 监听输入变化：实时同步下方路径标签、联动输出路径、进行合法性提示
        self.input_var.trace_add("write", self._on_input_var_changed)

        self.input_browse_btn = tk.Button(
            row_frame, text="  浏览...  ",
            font=self.font_btn,
            bg=self.brand, fg="#FFFFFF",
            activebackground=self.brand_hover, activeforeground="#FFFFFF",
            relief="flat", bd=0, padx=18, pady=6,
            cursor="hand2", disabledforeground="#FFFFFF",
            command=self.select_input_file
        )
        self.input_browse_btn.grid(row=0, column=1)

        # 下方显示完整目录路径
        self.input_path_label = tk.Label(
            card, text="尚未选择文件，可直接键入完整路径或点击浏览…", font=self.font_path,
            bg=self.card_bg, fg=self.muted, anchor="w"
        )
        self.input_path_label.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    def _build_output_card(self, parent, row):
        card = tk.Frame(
            parent, bg=self.card_bg, padx=18, pady=14,
            highlightthickness=1, highlightbackground=self.border
        )
        card.grid(row=row, column=0, sticky="ew", pady=(0, 18))
        card.columnconfigure(0, weight=1)

        tk.Label(
            card, text="💾  输出文件",
            font=self.font_section, bg=self.card_bg, fg=self.brand, anchor="w"
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        row_frame = tk.Frame(card, bg=self.card_bg)
        row_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        row_frame.columnconfigure(0, weight=1)

        self.output_var = tk.StringVar()
        self.output_entry = tk.Entry(
            row_frame, textvariable=self.output_var,
            font=self.font_field, relief="solid", bd=1,
            bg="#FFFFFF", fg="#323130",
            disabledbackground="#F5F5F5", disabledforeground="#A19F9D",
            highlightthickness=1,
            highlightcolor=self.brand, highlightbackground=self.border
        )
        self.output_entry.grid(row=0, column=0, sticky="ew", ipady=6, padx=(0, 8))
        # 监听输出变化：实时同步完整路径标签，并在缺少扩展名时自动补 .xlsx
        self.output_var.trace_add("write", self._on_output_var_changed)

        self.output_browse_btn = tk.Button(
            row_frame, text="  另存为...  ",
            font=self.font_btn,
            bg=self.brand, fg="#FFFFFF",
            activebackground=self.brand_hover, activeforeground="#FFFFFF",
            relief="flat", bd=0, padx=18, pady=6,
            cursor="hand2", disabledforeground="#FFFFFF",
            command=self.select_output_file
        )
        self.output_browse_btn.grid(row=0, column=1)

        self.output_path_label = tk.Label(
            card, text="可手动修改路径，默认与输入文件同目录", font=self.font_path,
            bg=self.card_bg, fg=self.muted, anchor="w"
        )
        self.output_path_label.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    def _build_template_card(self, parent, row):
        """
        数据模板下载卡片：展示说明 + 一个 "下载数据模板" 按钮。
        模板文件来源：
          - 源码运行：脚本同目录的 GRR/模板.xlsx
          - PyInstaller --onefile：被打包到 sys._MEIPASS 临时目录
        """
        card = tk.Frame(
            parent, bg=self.card_bg, padx=18, pady=14,
            highlightthickness=1, highlightbackground=self.border
        )
        card.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=0)

        tk.Label(
            card, text="📋  数据模板",
            font=self.font_section, bg=self.card_bg, fg=self.brand, anchor="w"
        ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))

        tk.Label(
            card,
            text="还没有数据模板？点击右侧按钮下载到本地：",
            font=self.font_field, bg=self.card_bg, fg="#323130", anchor="w"
        ).grid(row=1, column=0, sticky="ew", padx=(0, 12))

        self.template_btn = tk.Button(
            card, text="  📋 下载数据模板  ",
            font=self.font_btn,
            bg=self.brand, fg="#FFFFFF",
            activebackground=self.brand_hover, activeforeground="#FFFFFF",
            relief="flat", bd=0, padx=18, pady=6,
            cursor="hand2", disabledforeground="#FFFFFF",
            command=self.download_template
        )
        self.template_btn.grid(row=1, column=1)

    def _build_run_button(self, parent, row):
        btn_frame = tk.Frame(parent, bg="#FFFFFF")
        btn_frame.grid(row=row, column=0, pady=(0, 6))

        self.run_button = tk.Button(
            btn_frame, text="  ▶   生成 GRR 报告   ",
            font=self.font_run,
            bg=self.success, fg="#FFFFFF",
            activebackground=self.success_hover, activeforeground="#FFFFFF",
            relief="flat", bd=0, padx=28, pady=12,
            cursor="hand2", disabledforeground="#FFFFFF",
            command=self.run_program
        )
        self.run_button.pack()

    def _build_status(self, parent, row):
        status_frame = tk.Frame(parent, bg="#FFFFFF")
        status_frame.grid(row=row, column=0, sticky="ew", pady=(14, 0))
        status_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(status_frame, mode="indeterminate")
        self.progress.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        self.status_label = tk.Label(
            status_frame, text="●  等待选择输入文件…",
            font=self.font_status, bg="#FFFFFF", fg=self.muted, anchor="w"
        )
        self.status_label.grid(row=1, column=0, sticky="ew")

    # ------------------------------------------------------------------
    # 文件选择
    # ------------------------------------------------------------------
    def select_input_file(self):
        if self._busy:
            return
        file_path = filedialog.askopenfilename(
            title="选择原始宽表 Excel 文件",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        # 使用 suppress 标志避免 set() 触发的 trace 回调干扰浏览结果
        self._suppress_input_trace = True
        try:
            self.input_path = file_path
            # 只在 entry 中显示文件名，完整目录交给下方 label
            self.input_var.set(os.path.basename(file_path))
        finally:
            self._suppress_input_trace = False
        self.input_path_label.config(
            text=f"📁  {os.path.dirname(file_path)}", fg=self.muted
        )
        # 自动联动输出路径：仅在输出框为空时同步默认输出
        if not self.output_var.get().strip():
            base, _ = os.path.splitext(file_path)
            out = base + "_GRR_Report.xlsx"
            self._suppress_output_trace = True
            try:
                self.output_path = out
                self.output_var.set(os.path.basename(out))
            finally:
                self._suppress_output_trace = False
            self.output_path_label.config(
                text=f"📁  {os.path.dirname(out)}", fg=self.muted
            )
        self._set_status("●  已选择输入文件，可修改输出路径后点击生成", self.brand)

    def select_output_file(self):
        if self._busy:
            return
        initial_dir = os.path.dirname(self.output_path) if self.output_path else None
        file_path = filedialog.asksaveasfilename(
            title="选择输出报表文件",
            initialdir=initial_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file_path:
            return
        self._suppress_output_trace = True
        try:
            self.output_path = file_path
            self.output_var.set(os.path.basename(file_path))
        finally:
            self._suppress_output_trace = False
        self.output_path_label.config(
            text=f"📁  {os.path.dirname(file_path)}", fg=self.muted
        )

    # ------------------------------------------------------------------
    # 数据模板下载
    # ------------------------------------------------------------------
    def download_template(self):
        """
        弹"另存为"对话框，把内置模板 GRR/模板.xlsx 复制到用户指定的位置。
        """
        if self._busy:
            return

        template_src = get_template_path()
        if not os.path.exists(template_src):
            messagebox.showerror(
                "模板缺失",
                f"内置模板不存在：\n{template_src}\n\n"
                "· 源码运行：请把 GRR/模板.xlsx 放回源码目录\n"
                "· exe 打包：请在 PyInstaller spec 的 datas 中加入模板路径"
            )
            self._set_status("❌  内置模板文件缺失", self.danger)
            return

        # 默认目录：优先用输出文件所在目录，否则桌面，最后退回到 home
        if self.output_path:
            initial_dir = os.path.dirname(self.output_path)
        else:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            initial_dir = desktop if os.path.isdir(desktop) else os.path.expanduser("~")

        target = filedialog.asksaveasfilename(
            title="保存数据模板",
            initialdir=initial_dir,
            initialfile="GRR_数据模板.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not target:
            return

        try:
            # copy2 保留原文件的修改时间，使模板看起来"来自程序"
            shutil.copy2(template_src, target)
            self._set_status(f"✅  模板已保存到：{target}", self.success)
            messagebox.showinfo("完成", f"数据模板已下载：\n{target}")
        except Exception as e:
            self._set_status(f"❌  模板下载失败：{e}", self.danger)
            messagebox.showerror("下载失败", f"复制模板时出错：\n{e}")

    # ------------------------------------------------------------------
    # 输入框 / 输出框实时联动与校验
    # ------------------------------------------------------------------
    def _is_full_path(self, text):
        """判断字符串是否为完整路径（包含盘符或路径分隔符）"""
        if not text:
            return False
        if re.match(r'^[A-Za-z]:[\\/]', text):
            return True
        if text.startswith('\\\\'):
            return True
        if os.sep in text or (os.altsep and os.altsep in text):
            return True
        return False

    def _on_input_var_changed(self, *_args):
        """输入框内容变化时回调：同步路径、校验文件、联动输出"""
        if getattr(self, "_suppress_input_trace", False):
            return
        text = self.input_var.get().strip()

        if not text:
            # 用户清空输入
            self.input_path = None
            self.input_path_label.config(
                text="尚未选择文件，可直接键入完整路径或点击浏览…", fg=self.muted
            )
            self._set_status("●  等待选择输入文件…", self.muted)
            return

        # 用户键入了完整路径（含盘符或路径分隔符）
        if self._is_full_path(text):
            full_path = os.path.abspath(text)
            self.input_path = full_path
            self.input_path_label.config(
                text=f"📁  {os.path.dirname(full_path)}", fg=self.muted
            )
        else:
            # 用户键入的是裸文件名或相对路径，按当前目录解析
            full_path = os.path.abspath(text)
            self.input_path = full_path
            self.input_path_label.config(
                text=f"📁  {os.path.dirname(full_path)}", fg=self.muted
            )

        # 校验 1：文件是否存在
        if not os.path.exists(full_path):
            self.input_path_label.config(
                text=f"⚠  文件不存在：{full_path}", fg=self.danger
            )
            self._set_status("⚠  输入文件不存在，请检查路径", self.danger)
            return

        # 校验 2：是否为 Excel 文件
        if not full_path.lower().endswith(('.xlsx', '.xls')):
            self.input_path_label.config(
                text=f"⚠  请选择 Excel 文件 (*.xlsx / *.xls)：{full_path}",
                fg=self.danger,
            )
            self._set_status("⚠  输入文件必须是 Excel 文件", self.danger)
            return

        # 仅在输出框为空时，自动联动出默认输出路径
        if not self.output_var.get().strip():
            base, _ = os.path.splitext(full_path)
            new_out = base + "_GRR_Report.xlsx"
            self._suppress_output_trace = True
            try:
                self.output_path = new_out
                self.output_var.set(os.path.basename(new_out))
            finally:
                self._suppress_output_trace = False
            self.output_path_label.config(
                text=f"📁  {os.path.dirname(new_out)}", fg=self.muted
            )

        self._set_status(
            "●  已选择输入文件，可修改输出路径后点击生成", self.brand
        )

    def _on_output_var_changed(self, *_args):
        """输出框内容变化时回调：自动补扩展名、校验目录可写性"""
        if getattr(self, "_suppress_output_trace", False):
            return
        text = self.output_var.get().strip()

        if not text:
            self.output_path = None
            self.output_path_label.config(
                text="可手动修改路径，默认与输入文件同目录", fg=self.muted
            )
            return

        # 用户键入的可能仅是文件名（含完整路径则解析为绝对路径）
        if self._is_full_path(text):
            full_path = os.path.abspath(text)
        else:
            # 视为相对当前工作目录的文件名
            full_path = os.path.abspath(text)

        # 自动补齐 .xlsx 扩展名
        if not full_path.lower().endswith(('.xlsx', '.xls')):
            full_path += '.xlsx'

        self.output_path = full_path

        directory = os.path.dirname(full_path)
        # 校验目录是否存在；若不存在但在当前目录下可创建则不视为错误
        if directory and not os.path.isdir(directory):
            self.output_path_label.config(
                text=f"⚠  输出目录不存在：{directory}", fg=self.danger
            )
            self._set_status("⚠  输出目录不存在", self.danger)
            return

        self.output_path_label.config(
            text=f"📁  {directory}", fg=self.muted
        )

    # ------------------------------------------------------------------
    # 主流程
    # ------------------------------------------------------------------
    def run_program(self):
        if self._busy:
            return
        input_file = self.input_path
        output_file = self.output_path

        if not input_file:
            messagebox.showerror("提示", "请先选择输入文件")
            return
        if not os.path.exists(input_file):
            messagebox.showerror("提示", "输入文件不存在")
            return
        if not output_file:
            messagebox.showerror("提示", "请指定输出文件")
            return

        self._set_busy(True)
        try:
            self.progress.start(10)
            self.root.update_idletasks()
            process_grr_report(input_file, output_file)
            self.progress.stop()
            self._set_status("✅  处理完成，已自动打开报表", self.success)
            open_file_auto(output_file)
            messagebox.showinfo("完成", f"GRR 报表已生成：\n{output_file}")
        except Exception as e:
            self.progress.stop()
            self._set_status(f"❌  处理失败：{e}", self.danger)
            messagebox.showerror("错误", str(e))
        finally:
            self._set_busy(False)

    # ------------------------------------------------------------------
    # 辅助
    # ------------------------------------------------------------------
    def _set_busy(self, busy):
        self._busy = busy
        if busy:
            self.run_button.config(state="disabled", bg="#A0A0A0", cursor="arrow")
            self.input_browse_btn.config(state="disabled")
            self.output_browse_btn.config(state="disabled")
            self.input_entry.config(state="disabled")
            self.output_entry.config(state="disabled")
            self.template_btn.config(state="disabled")
            self._set_status("⏳  正在处理，请稍候…", self.brand)
        else:
            self.run_button.config(state="normal", bg=self.success, cursor="hand2")
            self.input_browse_btn.config(state="normal")
            self.output_browse_btn.config(state="normal")
            self.input_entry.config(state="normal")
            self.output_entry.config(state="normal")
            self.template_btn.config(state="normal")

    def _set_status(self, text, color):
        self.status_label.config(text=text, fg=color)

    def run(self):
        self.root.mainloop()


# =========================================================
# 主程序入口
# =========================================================
if __name__ == "__main__":
    app = GRRGUI()
    app.run()
