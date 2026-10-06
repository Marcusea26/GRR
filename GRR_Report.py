import os
import re
import math
import random
import warnings
import subprocess
import platform
import numpy as np
import pandas as pd
from scipy.stats import f
import tkinter as tk
from tkinter import filedialog, messagebox, font, ttk

warnings.filterwarnings("ignore", category=UserWarning)

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
# 写报表到 Excel
# =========================================================
def write_report(output_file, df_display, df_long, grr_df, tolerance_lower, tolerance_upper):
    """
    将计算结果写入 Excel，生成多个工作表并绘制相关图表
    """
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

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
# GUI 界面
# =========================================================
class GRRGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("GRR Report Generator")
        self.root.geometry("980x320")
        self.root.minsize(860, 280)
        self.root.resizable(True, True)

        self.custom_font = font.Font(family="KaiTi", size=14)
        self.title_font = font.Font(family="KaiTi", size=18, weight="bold")
        self.input_path = None

        self.build_ui()

    def build_ui(self):
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.grid(row=0, column=0, sticky="nsew")

        main_frame.grid_columnconfigure(1, weight=1)

        title_label = tk.Label(
            main_frame,
            text="GRR Report Generator",
            font=self.title_font,
            fg="#1F4E79"
        )
        title_label.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 20))

        tk.Label(main_frame, text="输入 Excel 文件：", font=self.custom_font).grid(row=1, column=0, sticky="w", pady=10)
        self.input_entry = tk.Entry(main_frame, font=self.custom_font, fg="red")
        self.input_entry.grid(row=1, column=1, padx=10, sticky="ew")
        tk.Button(main_frame, text="浏览", font=self.custom_font, command=self.select_input_file, width=10).grid(row=1, column=2, padx=(5, 0))

        tk.Label(main_frame, text="输出 Excel 文件：", font=self.custom_font).grid(row=2, column=0, sticky="w", pady=10)
        self.output_entry = tk.Entry(main_frame, font=self.custom_font, fg="blue")
        self.output_entry.grid(row=2, column=1, padx=10, sticky="ew")
        tk.Button(main_frame, text="另存为", font=self.custom_font, command=self.select_output_file, width=10).grid(row=2, column=2, padx=(5, 0))

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=25)

        self.run_button = tk.Button(
            btn_frame,
            text="运行",
            font=self.custom_font,
            width=12,
            height=2,
            bg="#D9EAD3"
        )
        self.run_button.config(command=self.run_program)
        self.run_button.pack()

        self.progress = ttk.Progressbar(main_frame, mode="indeterminate")
        self.progress.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        self.status_label = tk.Label(
            main_frame,
            text="请选择输入文件",
            font=self.custom_font,
            fg="green",
            anchor="w"
        )
        self.status_label.grid(row=5, column=0, columnspan=3, sticky="ew", pady=5)

    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择原始宽表Excel文件",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_path = file_path
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, file_path)

            base, ext = os.path.splitext(file_path)
            output_path = base + "_GRR_Report.xlsx"
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_path)

            self.status_label.config(text="已选择输入文件", fg="green")

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择输出报表文件",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)

    def run_program(self):
        input_file = self.input_entry.get().strip()
        output_file = self.output_entry.get().strip()

        if not input_file:
            messagebox.showerror("错误", "请先选择输入文件！")
            return

        if not os.path.exists(input_file):
            messagebox.showerror("错误", "输入文件不存在！")
            return

        if not output_file:
            messagebox.showerror("错误", "请指定输出文件！")
            return

        try:
            self.run_button.config(state="disabled")
            self.progress.start(10)
            self.status_label.config(text="正在处理，请稍候...", fg="blue")
            self.root.update()

            process_grr_report(input_file, output_file)

            self.progress.stop()
            self.run_button.config(state="normal")
            self.status_label.config(text="处理完成，已自动打开报表。", fg="green")

            open_file_auto(output_file)

            messagebox.showinfo("完成", f"GRR报表已生成：\n{output_file}")

        except Exception as e:
            self.progress.stop()
            self.run_button.config(state="normal")
            self.status_label.config(text="处理失败", fg="red")
            messagebox.showerror("错误", str(e))

    def run(self):
        self.root.mainloop()


# =========================================================
# 主程序入口
# =========================================================
if __name__ == "__main__":
    app = GRRGUI()
    app.run()
