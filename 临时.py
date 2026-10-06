import pandas as pd
import random
import tkinter as tk
from tkinter import filedialog, messagebox, font, ttk
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


plt.rcParams['font.sans-serif'] = ['SimHei']  # 支持中文字体
plt.rcParams['axes.unicode_minus'] = False    # 显示负号

# =====================================
# 常量定义
# =====================================
CPK_TARGET = 1.33
FREQ_MIN = 10
FREQ_MAX = 20000


def random_color():
    return "#{:06x}".format(random.randint(0, 0xFFFFFF))


def set_chart_style(chart, title_name, title_size, x_name, y_name,
                    min_x=None, max_x=None, log_base=None,
                    min_y=None, max_y=None, legend=True):
    chart.set_title({
        'name': title_name,
        'name_font': {
            'name': 'Arial',
            'color': 'black',
            'size': title_size,
            'bold': True
        }
    })
    chart.set_x_axis({
        'name': x_name,
        'name_font': {
            'name': 'Arial',
            'color': 'black',
            'size': title_size - 3,
            'bold': True
        },
        'min': min_x,
        'max': max_x,
        'log_base': log_base,
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside',
        'label_position': 'low'
    })
    chart.set_y_axis({
        'name': y_name,
        'name_font': {
            'name': 'Arial',
            'color': 'black',
            'size': title_size - 3,
            'bold': True
        },
        'min': min_y,
        'max': max_y,
        'major_gridlines': {'visible': True, 'line': {'color': 'black', 'width': 0.25}},
        'minor_tick_mark': 'inside',
        'major_tick_mark': 'inside'
    })
    chart.set_style(1)
    chart.set_plotarea({
        'border': {'color': '#808080', 'width': 1},
        'fill': {'color': '#F2F2F2'},
        'layout': {'x': 0.08, 'y': 0.1, 'width': 0.74, 'height': 0.8}
    })
    if legend:
        chart.set_legend({'layout': {'x': 0.83, 'y': 0.1, 'width': 0.14, 'height': 0.8}})


def create_line_chart(wb, sheet_name, df, row, col, main=True, title_size=10):
    chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth'})

    for j in range(2, 5):
        if j == 3:
            continue
        chart.add_series({
            'name': [sheet_name, j, 0],
            'categories': [sheet_name, 0, 7, 0, col - 1],
            'values': [sheet_name, j, 7, j, col - 1],
            'line': {'color': 'red', 'width': 1.5}
        })

    for j in range(5, row):
        chart.add_series({
            'name': [sheet_name, j, 0],
            'categories': [sheet_name, 0, 7, 0, col - 1],
            'values': [sheet_name, j, 7, j, col - 1],
            'line': {'color': random_color(), 'width': 1.5}
        })

    set_chart_style(chart, sheet_name, title_size, 'Frequency[Hz]', [sheet_name, 5, 5],
                    FREQ_MIN, FREQ_MAX, log_base=10, legend=True)
    chart.height = 500 if main else 400
    chart.width = 700 if main else 550
    return chart


def create_cpk_chart(wb, sheet_name, df, row, col, main=True, title_size=10):
    chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    chart.add_series({
        'name': [sheet_name, row, 6],
        'categories': [sheet_name, 0, 7, 0, col - 1],
        'values': [sheet_name, row, 7, row, col - 1],
        'line': {'color': 'green', 'width': 1.5}
    })
    chart.add_series({
        'name': str(CPK_TARGET),
        'categories': [sheet_name, 2, 1, 2, 2],
        'values': [sheet_name, 4, 1, 4, 2],
        'line': {'color': 'red', 'width': 1.5}
    })
    set_chart_style(chart, sheet_name + '--CPK', title_size, 'Frequency[Hz]', 'CPK',
                    FREQ_MIN, FREQ_MAX, log_base=10, min_y=0, max_y=5, legend=True)
    chart.height = 500 if main else 400
    chart.width = 600 if main else 550
    return chart


def create_scatter_chart(wb, sheet_name, df, row, col, main=True, title_size=10):
    chart = wb.add_chart({'type': 'scatter'})
    chart.add_series({
        'name': sheet_name,
        'categories': [sheet_name, 5, 0, row - 1, 0],
        'values': [sheet_name, 5, 7, row - 1, 7],
        'marker': {'type': 'circle', 'size': 4, 'border': {'color': 'blue'}, 'fill': {'color': 'blue'}},
        'line': None
    })
    chart.add_series({
        'name': '上框线',
        'categories': [sheet_name, 5, 0, row - 1, 0],
        'values': [sheet_name, 5, 9, row - 1, 9],
        'line': {'color': 'red', 'width': 1.25},
        'marker': {'type': 'none'}
    })
    chart.add_series({
        'name': '下框线',
        'categories': [sheet_name, 5, 0, row - 1, 0],
        'values': [sheet_name, 5, 10, row - 1, 10],
        'line': {'color': 'red', 'width': 1.25},
        'marker': {'type': 'none'}
    })
    set_chart_style(chart, sheet_name, 10, ' ', [sheet_name, 5, 5], legend=True)
    chart.height = 500 if main else 400
    chart.width = 700 if main else 550
    return chart


def create_bar_chart(wb, sheet_name, df, row, col, main=True, title_size=10):
    chart = wb.add_chart({'type': 'column'})
    chart.add_series({
        'name': sheet_name + '_Histogram',
        'categories': [sheet_name, 5, 17, 11, 17],
        'values': [sheet_name, 5, 16, 11, 16],
        'fill': {'color': '#98FB98'},
        'border': {'color': 'black'},
        'data_labels': {
            'value': True,
            'font': {
                'size': 15,
                'bold': True,
                'color': 'red'
            }
        }
    })

    set_chart_style(
        chart,
        sheet_name + '_Histogram',
        title_size,
        x_name='μ±3σ',
        y_name='Count',
        legend=True
    )

    chart.height = 500 if main else 400
    chart.width = 600 if main else 550
    return chart


# =====================================
# Excel 处理类
# =====================================
class ExcelProcessor:
    def __init__(self, input_path, output_path, skip_limits, scatter_mode, sort_by_dev, gui_callback=None):
        self.input_path = input_path
        self.output_path = output_path
        self.skip_limits = skip_limits
        self.scatter_mode = scatter_mode
        self.sort_by_dev = sort_by_dev
        self.gui_callback = gui_callback
        self._processed_data_for_ui = {}

    def process(self):
        try:
            sheet_names = pd.ExcelFile(self.input_path).sheet_names
            processed_sheets_data = {}
            sn_list = []

            with pd.ExcelWriter(self.output_path, engine='xlsxwriter') as writer:
                self._create_reports_sheet(writer)
                skipped_count = 0

                for sheet_number in range(0, len(sheet_names)):
                    df = pd.read_excel(
                        self.input_path,
                        sheet_name=sheet_names[sheet_number],
                        header=None,
                        engine='openpyxl'
                    )

                    if self.skip_limits:
                        upper_row_values = pd.to_numeric(df.iloc[2, 7:], errors='coerce')
                        lower_row_values = pd.to_numeric(df.iloc[4, 7:], errors='coerce')
                        if upper_row_values.dropna().empty and lower_row_values.dropna().empty:
                            skipped_count += 1
                            continue

                    sorted_df = self._process_single_sheet(
                        df, sheet_names[sheet_number],
                        writer, sheet_number, skipped_count
                    )
                    processed_sheets_data[sheet_names[sheet_number]] = sorted_df

                    if sheet_number == 0:
                        sn_list = sorted_df.iloc[5:, 0].tolist()
                        sn_list = [sn for sn in sn_list if pd.notna(sn) and str(sn).strip() != '']

                self._processed_data_for_ui = processed_sheets_data
                self._create_summary_sheet(writer, sn_list, sheet_names, processed_sheets_data)

                sheets_order = ['Summary'] + [s for s in writer.sheets if s != 'Summary']
                writer.book.worksheets_objs = [writer.sheets[s] for s in sheets_order]

            messagebox.showinfo("完成", "程序已成功运行并保存文件！")

            if self.gui_callback:
                self.gui_callback(list(self._processed_data_for_ui.keys()), self._processed_data_for_ui)

        except Exception as e:
            messagebox.showerror("错误", f"运行程序时出错: {str(e)}")

    def _create_reports_sheet(self, writer):
        df_reports = pd.DataFrame([], index=range(400), columns=range(50))
        df_reports.to_excel(writer, sheet_name='Reports', index=False, header=False)

    def _create_summary_sheet(self, writer, sn_list, sheet_names, processed_data):
        header = [''] + sheet_names + ['Result']

        sheet_limit_cols_map = {}
        for sheet_name in sheet_names:
            if sheet_name not in processed_data:
                continue
            df_sheet = processed_data[sheet_name]
            limit_cols = []
            for fc in range(7, df_sheet.shape[1] - 1):
                upper_val = pd.to_numeric([df_sheet.iloc[2, fc]], errors='coerce')[0]
                lower_val = pd.to_numeric([df_sheet.iloc[4, fc]], errors='coerce')[0]
                if not pd.isna(upper_val) or not pd.isna(lower_val):
                    limit_cols.append(fc)
            sheet_limit_cols_map[sheet_name] = limit_cols

        ws_reports = writer.sheets['Reports']

        for col_index, sheet_name in enumerate(sheet_names, start=1):
            pass_count_total = 0
            fail_count_total = 0
            limit_cols = sheet_limit_cols_map.get(sheet_name, [])

            for sn in sn_list:
                if sheet_name not in processed_data:
                    continue
                df_sheet = processed_data[sheet_name]
                sn_rows = df_sheet[df_sheet.iloc[:, 0] == sn]
                if sn_rows.empty:
                    continue
                sn_row = sn_rows.iloc[0]

                pass_fail_list = []
                for freq_col in limit_cols:
                    val = pd.to_numeric([sn_row[freq_col]], errors='coerce')[0]
                    upper = pd.to_numeric([df_sheet.iloc[2, freq_col]], errors='coerce')[0]
                    lower = pd.to_numeric([df_sheet.iloc[4, freq_col]], errors='coerce')[0]

                    if pd.isna(val):
                        continue
                    if not pd.isna(upper) and val > upper:
                        pass_fail_list.append('FAIL')
                    elif not pd.isna(lower) and val < lower:
                        pass_fail_list.append('FAIL')
                    else:
                        pass_fail_list.append('PASS')

                if not pass_fail_list:
                    continue
                if 'FAIL' in pass_fail_list:
                    fail_count_total += 1
                else:
                    pass_count_total += 1

            total_sn = pass_count_total + fail_count_total
            pass_rate = f"{(pass_count_total / total_sn) * 100:.2f}%" if total_sn > 0 else ''
            fail_rate = f"{(fail_count_total / total_sn) * 100:.2f}%" if total_sn > 0 else ''

            ws_reports.write((col_index - 1) * 22 + 3, 19, pass_rate)
            ws_reports.write((col_index - 1) * 22 + 4, 19, fail_rate)
            ws_reports.write((col_index - 1) * 22 + 3, 18, 'Pass_rate')
            ws_reports.write((col_index - 1) * 22 + 4, 18, 'Fail_rate')

        # 关键修复：dtype=object，避免 FutureWarning
        df_summary = pd.DataFrame(
            '',
            index=range(len(sn_list)),
            columns=range(len(sheet_names) + 2),
            dtype=object
        )

        for row_idx, sn in enumerate(sn_list):
            df_summary.iloc[row_idx, 0] = sn
            results_for_sn = []

            for col_index, sheet_name in enumerate(sheet_names, start=1):
                if sheet_name not in processed_data:
                    df_summary.iloc[row_idx, col_index] = 'NA'
                    continue

                df_sheet = processed_data[sheet_name]
                sn_rows = df_sheet[df_sheet.iloc[:, 0] == sn]
                if sn_rows.empty:
                    df_summary.iloc[row_idx, col_index] = 'NA'
                    continue

                result_val = sn_rows.iloc[0, 3]
                df_summary.iloc[row_idx, col_index] = result_val
                results_for_sn.append(result_val)

            valid_overall = [r for r in results_for_sn if r != 'NA']
            if not valid_overall:
                df_summary.iloc[row_idx, len(sheet_names) + 1] = 'NA'
            elif 'FAIL' in valid_overall:
                df_summary.iloc[row_idx, len(sheet_names) + 1] = 'FAIL'
            else:
                df_summary.iloc[row_idx, len(sheet_names) + 1] = 'PASS'

        df_summary = pd.concat([pd.DataFrame([header], dtype=object), df_summary], ignore_index=True)
        df_summary = df_summary.astype(object)

        df_summary = df_summary[
            ~df_summary.apply(lambda r: all(str(x).strip() in ['NA', ''] for x in r), axis=1)
        ]

        result_col = df_summary.iloc[1:, -1]
        valid_results = [r for r in result_col if r != 'NA']
        total_valid = len(valid_results)
        pass_count = valid_results.count('PASS')
        fail_count = valid_results.count('FAIL')

        pass_rate_summary = f"{(pass_count / total_valid) * 100:.2f}%" if total_valid > 0 else ''
        fail_rate_summary = f"{(fail_count / total_valid) * 100:.2f}%" if total_valid > 0 else ''

        df_summary.loc[0, len(sheet_names) + 4] = 'Data count'
        df_summary.loc[0, len(sheet_names) + 5] = total_valid
        df_summary.loc[1, len(sheet_names) + 4] = 'Result_Pass_rate'
        df_summary.loc[2, len(sheet_names) + 4] = 'Result_Fail_rate'
        df_summary.loc[1, len(sheet_names) + 5] = pass_rate_summary
        df_summary.loc[2, len(sheet_names) + 5] = fail_rate_summary

        df_summary.to_excel(writer, sheet_name='Summary', index=False, header=False)
        ws_summary = writer.sheets['Summary']

        for col_index, sheet_name in enumerate(sheet_names, start=1):
            ws_summary.write_url(0, col_index, f"internal:'{sheet_name}'!A1", string=sheet_name)

        fail_format = writer.book.add_format({'bg_color': '#FFC7CE'})
        pass_format = writer.book.add_format({'bg_color': '#C6EFCE'})

        start_row = 1
        end_row = len(df_summary) - 1
        start_col = 1
        end_col = len(header) - 1

        ws_summary.conditional_format(start_row, start_col, end_row, end_col, {
            'type': 'text',
            'criteria': 'containing',
            'value': 'FAIL',
            'format': fail_format
        })
        ws_summary.conditional_format(start_row, start_col, end_row, end_col, {
            'type': 'text',
            'criteria': 'containing',
            'value': 'PASS',
            'format': pass_format
        })

    def _process_single_sheet(self, df, sheet_name, writer, sheet_number, skipped_count):
        row, col = df.shape

        df.iloc[1, 6] = '平均值'
        df.iloc[3, 6] = '标准差'
        df.iloc[2, 0] = '上框线'
        df.iloc[4, 0] = '下框线'
        df.loc[4, col] = '偏差值'
        df.iloc[2, 1] = FREQ_MIN
        df.iloc[2, 2] = FREQ_MAX
        df.iloc[4, 1] = CPK_TARGET
        df.iloc[4, 2] = CPK_TARGET

        if self.scatter_mode and col < 15:
            upper_value = df.iloc[2, 7]
            df.loc[4, 9] = '上框线'
            for r in range(5, row):
                df.iloc[r, 9] = upper_value

            lower_value = df.iloc[4, 7]
            df.loc[4, 10] = '下框线'
            for r in range(5, row):
                df.iloc[r, 10] = lower_value

        for fre in range(7, col):
            df.iloc[1, fre] = pd.to_numeric(df.iloc[5:row, fre], errors='coerce').mean()
            df.iloc[3, fre] = pd.to_numeric(df.iloc[5:row, fre], errors='coerce').std()

        for freq in range(7, col):
            std_val = df.iloc[3, freq]
            if pd.isna(std_val) or std_val == 0:
                cpk = None
            else:
                cpk = min(
                    abs(df.iloc[1, freq] - df.iloc[2, freq]),
                    abs(df.iloc[1, freq] - df.iloc[4, freq])
                ) / (3 * std_val)
            df.loc[row, freq] = cpk

        df.iloc[row, 6] = 'CPK'

        row_data = df.iloc[2, 7:]
        non_empty_columns = row_data[row_data.notna()].index.tolist()
        if non_empty_columns:
            for count_row in range(5, row):
                gold = sum(
                    (df.iloc[count_row, f] - df.iloc[1, f]) ** 2
                    for f in range(non_empty_columns[0], non_empty_columns[-1] + 1)
                )
                df.loc[count_row, col] = gold

        header_df = df.iloc[:5]
        data_df = df.iloc[5:].copy()
        data_df[col] = pd.to_numeric(data_df[col], errors='coerce')

        if self.sort_by_dev:
            data_df = data_df.sort_values(by=col, ascending=True, na_position='last')

        df = pd.concat([header_df, data_df], ignore_index=True)

        df.iloc[4, 4] = '偏差值'
        df.iloc[5:, 4] = df.iloc[5:, col]

        limit_cols = []
        for fc in range(7, col):
            upper_val = pd.to_numeric([df.iloc[2, fc]], errors='coerce')[0]
            lower_val = pd.to_numeric([df.iloc[4, fc]], errors='coerce')[0]
            if not pd.isna(upper_val) or not pd.isna(lower_val):
                limit_cols.append(fc)

        for r in range(5, len(df) - 1):
            pass_fail_list = []
            for freq_col in limit_cols:
                val = pd.to_numeric([df.iloc[r, freq_col]], errors='coerce')[0]
                upper = pd.to_numeric([df.iloc[2, freq_col]], errors='coerce')[0]
                lower = pd.to_numeric([df.iloc[4, freq_col]], errors='coerce')[0]

                if pd.isna(val):
                    continue
                if not pd.isna(upper) and val > upper:
                    pass_fail_list.append('FAIL')
                elif not pd.isna(lower) and val < lower:
                    pass_fail_list.append('FAIL')
                else:
                    pass_fail_list.append('PASS')

            if not pass_fail_list:
                overall = 'NA'
            elif 'FAIL' in pass_fail_list:
                overall = 'FAIL'
            else:
                overall = 'PASS'
            df.iloc[r, 3] = overall

        if self.scatter_mode and col < 15:
            for idx in range(11, 19):
                df.loc[4, idx] = None

            if 13 not in df.columns:
                df[13] = None
            df[13] = df[13].astype(object)

            df.iloc[5, 13] = 'Mean'
            df.iloc[6, 13] = 'Count'
            df.iloc[7, 13] = 'Stdev'
            df.iloc[8, 13] = 'Valid_Lower'
            df.iloc[9, 13] = 'Valid_Upper'
            df.iloc[10, 13] = 'Mean_est'
            df.iloc[11, 13] = 'Count_est'
            df.iloc[12, 13] = 'Stdev_est'
            df.iloc[13, 13] = 'Exception_est'
            df.iloc[14, 13] = 'CPK'

            df.iloc[5, 14] = pd.to_numeric(df.iloc[5:row, 7], errors='coerce').mean()
            df.iloc[6, 14] = row - 5
            df.iloc[7, 14] = pd.to_numeric(df.iloc[5:row, 7], errors='coerce').std()
            df.iloc[8, 14] = df.iloc[5, 14] - 3.5 * df.iloc[7, 14]
            df.iloc[9, 14] = df.iloc[5, 14] + 3.5 * df.iloc[7, 14]

            def mean_in_range(series, lower, upper):
                series_numeric = pd.to_numeric(series, errors='coerce')
                filtered = series_numeric[(series_numeric >= lower) & (series_numeric <= upper)]
                mean_val = filtered.mean()
                std_val = filtered.std()
                count_val = len(filtered)
                return mean_val, count_val, std_val

            mean_val, count_val, std_val = mean_in_range(df.iloc[5:row, 7], df.iloc[8, 14], df.iloc[9, 14])

            df.iloc[10, 14] = mean_val
            df.iloc[11, 14] = count_val
            df.iloc[12, 14] = std_val
            df.iloc[13, 14] = df.iloc[6, 14] - df.iloc[11, 14]
            df.iloc[14, 14] = df.iloc[row, 7]

            df.iloc[5, 15] = df.iloc[10, 14] - df.iloc[12, 14] * 3.5
            df.iloc[6, 15] = df.iloc[10, 14] - df.iloc[12, 14] * 2.5
            df.iloc[7, 15] = df.iloc[10, 14] - df.iloc[12, 14] * 1.5
            df.iloc[8, 15] = df.iloc[10, 14] - df.iloc[12, 14] * 0.5
            df.iloc[9, 15] = df.iloc[10, 14] + df.iloc[12, 14] * 0.5
            df.iloc[10, 15] = df.iloc[10, 14] + df.iloc[12, 14] * 1.5
            df.iloc[11, 15] = df.iloc[10, 14] + df.iloc[12, 14] * 2.5
            df.iloc[12, 15] = df.iloc[10, 14] + df.iloc[12, 14] * 3.5

            src = pd.to_numeric(df.iloc[5:row, 7], errors='coerce')
            df.iloc[5, 16] = ((src > df.iloc[5, 15]) & (src < df.iloc[6, 15])).sum()
            df.iloc[6, 16] = ((src > df.iloc[6, 15]) & (src < df.iloc[7, 15])).sum()
            df.iloc[7, 16] = ((src > df.iloc[7, 15]) & (src < df.iloc[8, 15])).sum()
            df.iloc[8, 16] = ((src > df.iloc[8, 15]) & (src < df.iloc[9, 15])).sum()
            df.iloc[9, 16] = ((src > df.iloc[9, 15]) & (src < df.iloc[10, 15])).sum()
            df.iloc[10, 16] = ((src > df.iloc[10, 15]) & (src < df.iloc[11, 15])).sum()
            df.iloc[11, 16] = ((src > df.iloc[11, 15]) & (src < df.iloc[12, 15])).sum()

            df.iloc[5, 17] = round(df.iloc[10, 14] - df.iloc[12, 14] * 3, 2)
            df.iloc[6, 17] = round(df.iloc[10, 14] - df.iloc[12, 14] * 2, 2)
            df.iloc[7, 17] = round(df.iloc[10, 14] - df.iloc[12, 14] * 1, 2)
            df.iloc[8, 17] = round(df.iloc[10, 14] - df.iloc[12, 14] * 0, 2)
            df.iloc[9, 17] = round(df.iloc[10, 14] + df.iloc[12, 14] * 1, 2)
            df.iloc[10, 17] = round(df.iloc[10, 14] + df.iloc[12, 14] * 2, 2)
            df.iloc[11, 17] = round(df.iloc[10, 14] + df.iloc[12, 14] * 3, 2)

        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        ws = writer.sheets[sheet_name]
        fail_format = writer.book.add_format({'bg_color': '#FFC7CE'})
        pass_format = writer.book.add_format({'bg_color': '#C6EFCE'})

        ws.conditional_format(5, 3, len(df) - 1, 3, {
            'type': 'text',
            'criteria': 'containing',
            'value': 'FAIL',
            'format': fail_format
        })
        ws.conditional_format(5, 3, len(df) - 1, 3, {
            'type': 'text',
            'criteria': 'containing',
            'value': 'PASS',
            'format': pass_format
        })

        self._add_charts(writer, df, sheet_name, row, col, sheet_number, skipped_count)
        return df

    def _add_charts(self, writer, df, sheet_name, row, col, sheet_number, skipped_count):
        wb = writer.book
        ws = writer.sheets[sheet_name]
        ws2 = writer.sheets['Reports']
        base_index = (sheet_number - skipped_count) * 22 + 1

        if self.scatter_mode and col < 15:
            ws.insert_chart('A6', create_scatter_chart(wb, sheet_name, df, row, col, main=True))
            ws.insert_chart('L6', create_bar_chart(wb, sheet_name, df, row, col, main=True))
            ws2.insert_chart(f"A{base_index}", create_scatter_chart(wb, sheet_name, df, row, col, main=False))
            ws2.insert_chart(f"J{base_index}", create_bar_chart(wb, sheet_name, df, row, col, main=False))
        else:
            ws.insert_chart('A6', create_line_chart(wb, sheet_name, df, row, col, main=True))
            ws2.insert_chart(f"A{base_index}", create_line_chart(wb, sheet_name, df, row, col, main=False))
            ws.insert_chart('L6', create_cpk_chart(wb, sheet_name, df, row, col, main=True))
            ws2.insert_chart(f"J{base_index}", create_cpk_chart(wb, sheet_name, df, row, col, main=False))

        ws2.write(base_index, 18, 'Data count')
        ws2.write(base_index + 1, 18, 'MIN_CPK')
        ws2.write_url(f"S{base_index}", f"internal:'{sheet_name}'!A1", string=sheet_name)
        ws.write_url("B2", f"internal:'Reports'!A{base_index}", string="Reports")
        ws.write_url("A2", f"internal:'Summary'!A1", string="Summary")

        data_row_count = row - 5
        ws2.write(base_index, 19, data_row_count)

        cpk_values = df.loc[row, 7:col - 1]
        min_cpk = cpk_values.min(skipna=True)
        ws2.write(base_index + 1, 19, '' if pd.isna(min_cpk) else min_cpk)


# =====================================
# GUI 部分
# =====================================
class DOVEGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.state('zoomed')
        self.root.title("STAR(Statistics Transformation & Analysis Renderer)")

        self.custom_font = font.Font(family="KaiTi", size=20)
        self.full_input_path = None
        self.sheet_data_map = {}

        self.current_fig = None
        self.current_canvas = None
        self.is_closing = False

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self._build_ui()

    def _build_ui(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        left_frame = tk.Frame(main_frame, width=400)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        tk.Label(left_frame, text="Select Excel File:", font=self.custom_font).pack(pady=20)

        self.input_entry = tk.Entry(left_frame, width=40, font=self.custom_font, justify='center', fg="red")
        self.input_entry.pack(pady=10)

        tk.Button(left_frame, text="Browse", font=self.custom_font, command=self.select_input_file).pack(pady=10)

        tk.Label(left_frame, text="Output Excel File:", font=self.custom_font).pack(pady=10)
        self.output_entry = tk.Entry(left_frame, width=40, font=self.custom_font, justify='center', fg="red")
        self.output_entry.pack(pady=10)

        self.include_var = tk.BooleanVar()
        tk.Checkbutton(left_frame, text="Skip Data Without Limits", font=self.custom_font,
                       variable=self.include_var).pack(pady=10)

        self.scatter_var = tk.BooleanVar()
        tk.Checkbutton(left_frame, text="Single-valued test item", font=self.custom_font,
                       variable=self.scatter_var).pack(pady=10)

        self.sort_var = tk.BooleanVar()
        tk.Checkbutton(left_frame, text="Sort by deviation", font=self.custom_font,
                       variable=self.sort_var).pack(pady=10)

        tk.Button(left_frame, text="Run", font=self.custom_font, command=self.run_program).pack(pady=20)

        tk.Label(left_frame, text="PDF file:", font=self.custom_font).pack(pady=5)
        self.pdf_name_entry = tk.Entry(left_frame, width=40, font=self.custom_font, justify='center', fg="red")
        self.pdf_name_entry.pack(pady=5)

        self.export_all_btn = tk.Button(left_frame, text="Export all figure", font=self.custom_font,
                                        command=self.export_all_pdf)
        self.export_all_btn.pack(pady=10)

        tk.Label(left_frame, text="Marcus.Li", font=self.custom_font, fg="white").pack(
            side='bottom', anchor='w', padx=10, pady=10
        )

        sheet_top = tk.Frame(right_frame)
        sheet_top.pack(side='top', fill='x', pady=10)

        tk.Label(sheet_top, text="Sheet:", font=self.custom_font).pack(side='left', padx=5)

        self.sheet_selector = ttk.Combobox(sheet_top, font=self.custom_font, state="readonly", width=30)
        self.sheet_selector.pack(side='left', padx=5)
        self.sheet_selector.bind("<<ComboboxSelected>>", self.display_sheet_charts)

        self.chart_canvas_frame = tk.Frame(right_frame)
        self.chart_canvas_frame.pack(side='top', fill=tk.BOTH, expand=True, padx=10, pady=10)

    def on_close(self):
        if self.is_closing:
            return
        self.is_closing = True

        try:
            if self.current_canvas is not None:
                try:
                    self.current_canvas.get_tk_widget().destroy()
                except Exception:
                    pass
                self.current_canvas = None

            if self.current_fig is not None:
                try:
                    plt.close(self.current_fig)
                except Exception:
                    pass
                self.current_fig = None

            plt.close('all')

            try:
                self.root.quit()
            except Exception:
                pass

            try:
                self.root.destroy()
            except Exception:
                pass

        except Exception as e:
            print("关闭窗口时出错：", e)

    def select_input_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.full_input_path = file_path
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, os.path.basename(file_path))

            output_path = file_path.replace('.xlsx', '--整理.xlsx')
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, os.path.basename(output_path))

            pdf_name = os.path.splitext(os.path.basename(file_path))[0] + "--制图.pdf"
            self.pdf_name_entry.delete(0, tk.END)
            self.pdf_name_entry.insert(0, pdf_name)

    def export_all_pdf(self):
        if not self.sheet_data_map:
            messagebox.showerror("错误", "请先运行生成数据后再导出！")
            return

        if not self.full_input_path:
            messagebox.showerror("错误", "请先选择输入文件！")
            return

        output_dir = os.path.dirname(self.full_input_path)
        default_name = os.path.splitext(os.path.basename(self.full_input_path))[0] + "--制图.pdf"
        pdf_name = self.pdf_name_entry.get().strip() or default_name
        pdf_path = os.path.join(output_dir, pdf_name)

        try:
            with PdfPages(pdf_path) as pdf:
                for sheet_name, df in self.sheet_data_map.items():
                    rows, cols = df.shape
                    is_single = self.scatter_var.get() and cols > 13 and df.iloc[5, 13] == 'Mean'

                    fig, axes = plt.subplots(2, 1, figsize=(10, 7))
                    fig.subplots_adjust(hspace=0.3)
                    fig.suptitle(sheet_name, fontsize=14)

                    if is_single:
                        data_end_row = rows - 1
                        scatter_y = pd.to_numeric(df.iloc[5:data_end_row, 7], errors='coerce')
                        scatter_x = pd.to_numeric(df.iloc[5:data_end_row, 0], errors='coerce')
                        if scatter_x.isna().all():
                            scatter_x = range(5, data_end_row)

                        axes[0].scatter(scatter_x, scatter_y, s=30, color='blue')

                        upper_line = pd.to_numeric(df.iloc[5:data_end_row, 9], errors='coerce')
                        lower_line = pd.to_numeric(df.iloc[5:data_end_row, 10], errors='coerce')

                        if not upper_line.isna().all():
                            axes[0].plot(scatter_x, upper_line, color='red', linewidth=1)
                        if not lower_line.isna().all():
                            axes[0].plot(scatter_x, lower_line, color='red', linewidth=1)

                        axes[0].set_title(f"{sheet_name}", fontsize=12)
                        axes[0].grid(True, linestyle='--', linewidth=0.5)
                        for spine in axes[0].spines.values():
                            spine.set_visible(True)

                        hist_y = pd.to_numeric(df.iloc[5:12, 16], errors='coerce')
                        hist_x = pd.to_numeric(df.iloc[5:12, 17], errors='coerce')
                        if not hist_y.isna().all() and not hist_x.isna().all():
                            width = (max(hist_x) - min(hist_x)) / (len(hist_x) * 2) if len(hist_x) > 1 else 0.1
                            axes[1].bar(
                                hist_x, hist_y, width=width, align='center',
                                color='#98FB98', edgecolor='black', linewidth=1.2
                            )
                            for x, y in zip(hist_x, hist_y):
                                axes[1].text(
                                    x, y + 0.5, f"{int(y)}",
                                    ha='center', va='bottom',
                                    color='red', fontsize=10, fontweight='bold'
                                )

                        axes[1].set_xlabel('μ±3σ')
                        axes[1].set_ylabel('Count')
                        axes[1].set_title(f"{sheet_name}_Histogram", fontsize=12)
                        axes[1].grid(True, linestyle='--', linewidth=0.5, color='gray')
                        for spine in axes[1].spines.values():
                            spine.set_visible(True)

                    else:
                        freq = pd.to_numeric(df.iloc[0, 7:], errors='coerce')
                        if freq.isna().all():
                            freq = pd.to_numeric(df.iloc[2, 7:], errors='coerce')

                        for r in range(5, rows - 1):
                            vals = pd.to_numeric(df.iloc[r, 7:], errors='coerce')
                            if not vals.isna().all():
                                axes[0].plot(freq, vals, linewidth=1)

                        upper_vals = pd.to_numeric(df.iloc[2, 7:], errors='coerce')
                        lower_vals = pd.to_numeric(df.iloc[4, 7:], errors='coerce')

                        if not upper_vals.isna().all():
                            axes[0].plot(freq, upper_vals, color='red', linestyle='--', linewidth=1)
                        if not lower_vals.isna().all():
                            axes[0].plot(freq, lower_vals, color='red', linestyle='--', linewidth=1)

                        axes[0].set_xscale('log')
                        axes[0].set_title(f"{sheet_name}", fontsize=12)
                        axes[0].grid(True, linestyle='--', linewidth=0.5)
                        for spine in axes[0].spines.values():
                            spine.set_visible(True)

                        if 'CPK' in df.iloc[:, 6].values:
                            cpk_row_index = df[df.iloc[:, 6] == 'CPK'].index[0]
                            cpk_vals = pd.to_numeric(df.iloc[cpk_row_index, 7:], errors='coerce')
                            if not cpk_vals.isna().all():
                                axes[1].plot(freq, cpk_vals, color='green', linewidth=1)

                            axes[1].axhline(y=1.33, color='red', linestyle='--', linewidth=1)
                            axes[1].set_ylim(0, 5)
                            axes[1].set_xscale('log')

                        axes[1].set_title(f"{sheet_name}--CPK", fontsize=12)
                        axes[1].grid(True, linestyle='--', linewidth=0.5)
                        for spine in axes[1].spines.values():
                            spine.set_visible(True)

                    pdf.savefig(fig)
                    plt.close(fig)

            plt.close('all')
            messagebox.showinfo("成功", f"所有图表已导出到 PDF：\n{pdf_path}")

        except Exception as e:
            messagebox.showerror("错误", f"导出PDF时出错: {str(e)}")

    def run_program(self):
        if not self.full_input_path:
            messagebox.showerror("错误", "请先选择输入文件！")
            return

        output_path = os.path.join(os.path.dirname(self.full_input_path), self.output_entry.get())

        processor = ExcelProcessor(
            self.full_input_path,
            output_path,
            self.include_var.get(),
            self.scatter_var.get(),
            self.sort_var.get(),
            gui_callback=self.update_sheet_selector
        )
        processor.process()

    def update_sheet_selector(self, sheet_names, sheet_data_map):
        self.sheet_data_map = sheet_data_map
        self.sheet_selector['values'] = sheet_names
        if sheet_names:
            self.sheet_selector.current(0)
            self.display_sheet_charts()

    def display_sheet_charts(self, event=None):
        for widget in self.chart_canvas_frame.winfo_children():
            try:
                widget.destroy()
            except Exception:
                pass

        if self.current_fig is not None:
            try:
                plt.close(self.current_fig)
            except Exception:
                pass
            self.current_fig = None

        self.current_canvas = None

        current_sheet = self.sheet_selector.get()
        if not current_sheet or current_sheet not in self.sheet_data_map:
            return

        df = self.sheet_data_map[current_sheet]
        rows, cols = df.shape

        is_single = self.scatter_var.get() and cols > 13 and df.iloc[5, 13] == 'Mean'

        fig, axes = plt.subplots(2, 1, figsize=(12, 5))

        if is_single:
            data_end_row = rows - 1
            scatter_x = range(5, data_end_row)
            scatter_y = pd.to_numeric(df.iloc[5:data_end_row, 7], errors='coerce')

            axes[0].scatter(scatter_x, scatter_y, s=30, color='blue', label='Data')

            upper_line = pd.to_numeric(df.iloc[5:data_end_row, 9], errors='coerce')
            if not upper_line.isna().all():
                axes[0].plot(scatter_x, upper_line, color='red', linewidth=1)

            lower_line = pd.to_numeric(df.iloc[5:data_end_row, 10], errors='coerce')
            if not lower_line.isna().all():
                axes[0].plot(scatter_x, lower_line, color='red', linewidth=1)

            axes[0].set_title(f"{current_sheet}", fontsize=12)
            axes[0].grid(True, linestyle='--', linewidth=0.5)
            for spine in axes[0].spines.values():
                spine.set_visible(True)
            if axes[0].get_legend():
                axes[0].legend().remove()

            hist_y = pd.to_numeric(df.iloc[5:12, 16], errors='coerce')
            hist_x = pd.to_numeric(df.iloc[5:12, 17], errors='coerce')
            if not hist_y.isna().all() and not hist_x.isna().all():
                width = (max(hist_x) - min(hist_x)) / (len(hist_x) * 2) if len(hist_x) > 1 else 0.1
                axes[1].bar(
                    hist_x, hist_y, width=width, align='center',
                    color='#98FB98', edgecolor='black', linewidth=1.2
                )

                for x, y in zip(hist_x, hist_y):
                    axes[1].text(
                        x, y + 0.5, f"{int(y)}",
                        ha='center', va='bottom',
                        color='red', fontsize=10, fontweight='bold'
                    )

            axes[1].set_xlabel('μ±3σ', fontsize=10)
            axes[1].set_ylabel('Count', fontsize=10)
            axes[1].set_title(f"{current_sheet}_Histogram", fontsize=12)
            if len(hist_x.dropna()) > 0:
                axes[1].set_xticks([round(v, 2) for v in hist_x.dropna()])
            axes[1].grid(True, linestyle='--', linewidth=0.5, color='gray')
            for spine in axes[1].spines.values():
                spine.set_visible(True)
            if axes[1].get_legend():
                axes[1].legend().remove()

        else:
            freq = pd.to_numeric(df.iloc[0, 7:], errors='coerce')
            if freq.isna().all():
                freq = pd.to_numeric(df.iloc[2, 7:], errors='coerce')

            for r in range(5, rows - 1):
                vals = pd.to_numeric(df.iloc[r, 7:], errors='coerce')
                if not vals.isna().all():
                    axes[0].plot(freq, vals, linewidth=1)

            upper_vals = pd.to_numeric(df.iloc[2, 7:], errors='coerce')
            if not upper_vals.isna().all():
                axes[0].plot(freq, upper_vals, color='red', linestyle='--', linewidth=1)

            lower_vals = pd.to_numeric(df.iloc[4, 7:], errors='coerce')
            if not lower_vals.isna().all():
                axes[0].plot(freq, lower_vals, color='red', linestyle='--', linewidth=1)

            axes[0].set_xscale('log')
            axes[0].set_title(f"{current_sheet}", fontsize=12)
            axes[0].grid(True, linestyle='--', linewidth=0.5)
            for spine in axes[0].spines.values():
                spine.set_visible(True)
            if axes[0].get_legend():
                axes[0].legend().remove()

            if 'CPK' in df.iloc[:, 6].values:
                cpk_row_index = df[df.iloc[:, 6] == 'CPK'].index[0]
                cpk_vals = pd.to_numeric(df.iloc[cpk_row_index, 7:], errors='coerce')
                if not cpk_vals.isna().all():
                    axes[1].plot(freq, cpk_vals, color='green', linewidth=1)
                axes[1].axhline(y=1.33, color='red', linestyle='--', linewidth=1)
                axes[1].set_ylim(0, 5)
                axes[1].set_xscale('log')

            axes[1].set_title(f"{current_sheet}--CPK", fontsize=12)
            axes[1].grid(True, linestyle='--', linewidth=0.5)
            for spine in axes[1].spines.values():
                spine.set_visible(True)
            if axes[1].get_legend():
                axes[1].legend().remove()

        self.current_fig = fig
        self.current_canvas = FigureCanvasTkAgg(fig, master=self.chart_canvas_frame)
        self.current_canvas.draw()
        self.current_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def run(self):
        try:
            self.root.mainloop()
        finally:
            try:
                plt.close('all')
            except Exception:
                pass


if __name__ == "__main__":
    app = DOVEGUI()
    app.run()
