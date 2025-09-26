import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import pandas as pd
import yaml


def get_column_name(col_num: int) -> str:
    """列号转Excel列名"""
    col_name = ''
    while col_num >= 0:
        col_name = chr(col_num % 26 + 65) + col_name
        col_num = col_num // 26 - 1
        if col_num < 0:
            break
    return col_name

def load_config(config_path: str = "config.yaml"):
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    return config

def generate_weekly_report(csv_path: str, output_dir: str, start_date: str = "2020-01-09", config_path: str = "config.yaml"):
    try:
        # 读取配置
        config = load_config(config_path)
        report_cfg = config['report']
        style_cfg = config['excel_style']
        filter_countrylist = config['filter_countrylist']
        target_value = config['target_value']

        # 输入数据
        data = pd.read_csv(csv_path, header=0)
        data['注册日期'] = pd.to_datetime(data['注册日期']).dt.date
        start_date = datetime.strptime(report_cfg['start_date'], "%Y-%m-%d").date()

        # 输出路径
        current_time = datetime.now().strftime(report_cfg['output_time_format'])
        outpath = Path(output_dir) / f"{report_cfg['file_prefix']}_{current_time}.xlsx"

        # 买量情况
        df_cost = data[data.OS.isnull()][['注册日期', '国家', '新增', '花费', 'eCPI']].reset_index()
        df_cost_1 = data[data.OS.isnull() & data.国家.isin(filter_countrylist)][['注册日期', '国家', '新增', '花费', 'eCPI']].reset_index()
        pivot_df1 = pd.pivot_table(df_cost_1, values='新增', index='注册日期', columns='国家', aggfunc='sum', fill_value=0)
        pivot_df2 = pd.pivot_table(df_cost_1, values='花费', index='注册日期', columns='国家', aggfunc='sum', fill_value=0)
        pivot_df3 = pd.pivot_table(df_cost_1, values='eCPI', index='注册日期', columns='国家', aggfunc='sum', fill_value=0)

        for df, col in zip([pivot_df1, pivot_df2, pivot_df3], ['新增', '花费', 'eCPI']):
            df.replace(0, "", inplace=True)
            df['合计'] = df_cost[df_cost.国家.isnull()][['注册日期', col]].set_index('注册日期')

        df_cost_dir = {
            "新增-分国家": pivot_df1,
            "花费-分国家": pivot_df2,
            "eCPI-分国家": pivot_df3,
        }
        #   整体
        df_1 = data[data.OS.isnull() & data.国家.isnull()].drop(columns=['OS', '国家']).T
        #   分系统
        df_2 = data[(data.OS == 'Android') & data.国家.isnull()].drop(columns=['OS', '国家']).T
        df_3 = data[(data.OS == 'iOS') & data.国家.isnull()].drop(columns=['OS', '国家']).T
        #   分国家
        df_4 = data[data.OS.isnull() & (data.国家 == '美国')].drop(columns=['OS', '国家']).T
        df_5 = data[data.OS.isnull() & (data.国家 == '港澳台')].drop(columns=['OS', '国家']).T
        df_6 = data[data.OS.isnull() & (data.国家 == '德法英')].drop(columns=['OS', '国家']).T
        df_7 = data[data.OS.isnull() & (data.国家 == '欧盟加澳')].drop(columns=['OS', '国家']).T
        df_8 = data[data.OS.isnull() & (data.国家 == '韩国')].drop(columns=['OS', '国家']).T
        df_9 = data[data.OS.isnull() & (data.国家 == '日本')].drop(columns=['OS', '国家']).T

        df_dir = {
            "KPI-整体": df_1,
            "KPI-Android": df_2,
            "KPI-iOS": df_3,
            "KPI-美国": df_4,
            "KPI-港澳台": df_5,
            "KPI-德法英": df_6,
            "KPI-欧盟加澳": df_7,
            "KPI-韩国": df_8,
            "KPI-日本": df_9,
        }

        # 输出 Excel
        with pd.ExcelWriter(outpath, engine='xlsxwriter') as writer:
            workbook = writer.book
            sheet1 = workbook.add_worksheet('买量')

            # 格式定义
            percent_format = workbook.add_format(style_cfg['formats']['percent_format'])
            decimal_format = workbook.add_format(style_cfg['formats']['decimal_format'])
            money_format = workbook.add_format(style_cfg['formats']['money_format'])
            int_format = workbook.add_format(style_cfg['formats']['int_format'])
            header_format = workbook.add_format(style_cfg['formats']['header_format'])

            #   添加上框线
            top_percent_format = workbook.add_format(style_cfg['formats']['top_percent_format'])
            top_decimal_format = workbook.add_format(style_cfg['formats']['top_decimal_format'])
            top_money_format = workbook.add_format(style_cfg['formats']['top_money_format'])
            top_int_format = workbook.add_format(style_cfg['formats']['top_int_format'])
            # 设置字体，居左，粗体
            bold_left_format = workbook.add_format(style_cfg['formats']['bold_left_format'])
            # 设置字体，居中，粗体
            bold_centered_format = workbook.add_format(style_cfg['formats']['bold_centered_format'])
            # 设置居中
            centered_format = workbook.add_format(style_cfg['formats']['centered_format'])

            csr_1 = style_cfg['color_scales']['csr1']
            csr_11 = style_cfg['color_scales']['csr11']
            csr_3 = style_cfg['color_scales']['csr3']
            #   数据条格式
            #   蓝色
            dbf_3 = style_cfg['date_bars']['dbf_3']

            top_border_format = workbook.add_format()
            top_border_format.set_top(2)  # 数值越大，边框越粗
            bottom_border_format = workbook.add_format()
            bottom_border_format.set_bottom(2)  # 数值越大，边框越粗
            left_border_format = workbook.add_format()
            left_border_format.set_left(2)  # 数值越大，边框越粗
            right_border_format = workbook.add_format()
            right_border_format.set_right(2)  # 数值越大，边框越粗
            #   设置指定行格式

            #   以3个为一组，标记条件格式区域：新增、eCPI......
            #   注意逗号分隔
            df_reset = df_1.reset_index()
            #   提取行号-对应界定指标
            matching_rows = df_reset.index[df_reset.iloc[:, 0].isin(target_value)].tolist()
            new_match = list(map(lambda x: x + 1, matching_rows))
            # 设置对应顺序的格式
            num_format = [percent_format, percent_format, decimal_format, percent_format, percent_format,
                          percent_format, percent_format, percent_format, percent_format, decimal_format,
                          percent_format,
                          decimal_format, decimal_format, decimal_format]
            top_num_format = [top_percent_format, top_percent_format, top_decimal_format, top_percent_format,
                              top_percent_format, top_percent_format, top_percent_format, top_percent_format,
                              top_percent_format,
                              top_decimal_format, top_percent_format,
                              top_decimal_format, top_decimal_format, top_decimal_format]

            # 买量 sheet
            start_row, row_list = 0, []
            for name, df in df_cost_dir.items():
                sheet = writer.sheets['买量']
                sheet.merge_range(f'A{start_row + 1}:B{start_row + 1}', name, header_format)
                df.to_excel(writer, sheet_name='买量', startrow=start_row + 1, index=True, header=True)
                start_row += len(df) + 3
                row_list.append(start_row)

            for i in range(row_list[0]):
                sheet1.set_row(i, 17, int_format)
            for i in range(row_list[0], row_list[2]):
                sheet1.set_row(i, 17, money_format)
            sheet1.set_column("A:A", 20)
            sheet1.set_column("B:AA", 13)
            sheet1.conditional_format("A2:AA2", {'type': 'no_blanks', 'format': header_format})
            sheet1.conditional_format(f"B1:AA{row_list[0]}", csr_1)
            sheet1.conditional_format(f"A{row_list[0]}:AA{row_list[0] + 2}", {'type': 'no_blanks', 'format': header_format})
            sheet1.conditional_format(f"B{row_list[0]}:AA{row_list[1]}", csr_1)
            sheet1.conditional_format(f"A{row_list[1]}:AA{row_list[1] + 2}", {'type': 'no_blanks', 'format': header_format})
            sheet1.conditional_format(f"B{row_list[1]}:AA{row_list[2]}", csr_11)

            # KPI sheet
            df_reset = df_dir["KPI-整体"].reset_index()
            row1 = df_reset.index[df_reset.iloc[:, 0] == '新增'].tolist()
            row2 = df_reset.index[df_reset.iloc[:, 0] == 'eCPI'].tolist()
            row3 = df_reset.index[df_reset.iloc[:, 0] == '花费'].tolist()
            row4 = df_reset.index[df_reset.iloc[:, 0] == 'D7_payer'].tolist()
            row5 = df_reset.index[df_reset.iloc[:, 0] == 'D7_CPP'].tolist()

            for name, df in df_dir.items():
                df.replace(0, '', inplace=True)
                sheet = writer.book.add_worksheet(name)
                df.to_excel(writer, sheet_name=name, startrow=0, index=True, header=False, freeze_panes=(1, 1))

                sheet.set_column("A:A", 20, header_format)
                sheet.set_column("B:BB", 17, decimal_format)
                sheet.set_row(0, 20, header_format)
                sheet.conditional_format(f"A1:{get_column_name(len(df.iloc[0, :]))}1", {'type': 'no_blanks', 'format': header_format})

                sheet.conditional_format(f"A{row1[0] + 1}:BB{row1[0] + 1}", csr_1)
                sheet.conditional_format(f"A{row2[0] + 1}:BB{row2[0] + 1}", csr_3)
                sheet.conditional_format(f"A{row5[0] + 1}:BB{row5[0] + 1}", csr_3)
                if row1: sheet.set_row(row1[0], 17, int_format)
                if row2: sheet.set_row(row2[0], 17, money_format)
                if row3: sheet.set_row(row3[0], 17, money_format)
                if row4: sheet.set_row(row4[0], 17, int_format)
                if row5: sheet.set_row(row5[0], 17, money_format)

                num_loc = 0
                i = 0
                while i < len(new_match) - 2:
                    sheet.conditional_format(f"A{new_match[i]}:BB{new_match[i + 2]}", csr_1)
                    sheet.conditional_format(f"A{new_match[i + 1]}:BB{new_match[i + 1]}", dbf_3)
                    for j in range(new_match[i] - 1, new_match[i + 2]):
                        sheet.set_row(j, 17, num_format[num_loc])
                        sheet.set_row(new_match[i] - 1, 17, top_num_format[num_loc])
                    i += 3
                    num_loc += 1

        messagebox.showinfo("✅ 完成", f"周报生成成功：{outpath}")
        return outpath

    except Exception as e:
        messagebox.showerror("❌ 错误", str(e))
        return None

