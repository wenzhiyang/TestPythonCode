import datetime
import calendar
from openpyxl.styles import PatternFill
import openpyxl
from openpyxl import load_workbook
import numpy as np
import pandas as pd
from SNFunctionList import *
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.styles import numbers
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.dimensions import ColumnDimension, RowDimension
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment, numbers
from openpyxl.utils import range_boundaries
from openpyxl.formula.translate import Translator
from pycel import ExcelCompiler
import os
import pandas as pd
import mysql.connector

from datetime import datetime, timedelta
import re
from openpyxl import Workbook
import calendar
from dateutil.relativedelta import relativedelta
from collections import defaultdict
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import datetime
# 连接到 MySQL 数据库

import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import datetime
import re



common_columns = [
         '2#', '转S1-1#', '2-1#', '2-1T#', '2A-1TB2#', '2A-2#', '2A',
        'R6#', '石S1-1#', '石S1-1T#', 'GHBA-4#', '针S1-1#', '针S1-1T#',
        'S1-2#', 'S1-3T#', '高S2-2#', 'S2-2T#', 'S2-1#', '低S2-2#','S2-5#',
        'S4-1#','S5-1#'
    ]
customer_dict = {
    'A001': '安德丰',
    'A002': '安普瑞斯',
    'A003': '埃克森',
    'B001': '保力新',
    'B002': '比亚迪',
    'G001': '国轩',
    'J001': '力神聚元',
    'L001': '力神',
    'L003': '朗泰通',
    'L004': '朗泰沣',
    'N002': '宁德时代',
    'N003': '能优',
    'S003': '索理德',
    'S004':'三一',
    'T002': '拓邦',
    'T004': '天鹏',
    'T005': '天弋',
    'W001':'沃能',
    'W002': '五行',
    'X001':'鑫辉',
    'X003': '星恒'
}
def parse_rateorder_file(file_path, plan_file_path):
    """读取配比单文件并处理数据，增加计划部文件路径参数"""
    try:
        # 初始化分配详情列表
        allocation_details = []

        # 读取计划部文件获取计划数
        plan_df = pd.read_excel(
            plan_file_path,
            sheet_name='2025年6月成品排产跟踪',
            skiprows=2,  # 跳过前两行（标题和空行）
            usecols='A:B,D',  # 客户名称(A), 产品型号(B), 计划数(D)
            names=['客户名称', '产品型号', '计划数']
        )
        plan_df['计划数'] = pd.to_numeric(plan_df['计划数'], errors='coerce').fillna(0)

        # 修改：增加配方编号列(A列)
        df = pd.read_excel(
            file_path,
            sheet_name='每日配比单进度跟踪',
            skiprows=1,
            usecols='B,D:H',  # 增加A列(配方编号)
            header=None,
            names=['配方编号', '产品型号', '客户代码', '配比单生产总量(公斤)', '已生产数量', '备注']
        )

        # 筛选备注不包含"生产完"的行
        df = df[df['备注'].fillna('').str.contains('生产完') == False]

        # 计算剩余生产量
        df['剩余生产量'] = (df['配比单生产总量(公斤)'] - df['已生产数量']) / 1000
        df.reset_index(drop=True, inplace=True)

        # 拆分客户代码列
        df_expanded_list = []

        for idx, row in df.iterrows():
            formula_code = row['配方编号']  # 获取配方编号
            customer_codes = str(row['客户代码']).replace(',', ' ').split()
            pattern = r'^[A-Za-z]\d{3}$'
            valid_codes = [code for code in customer_codes if pd.Series(code).str.match(pattern)[0]]

            if not valid_codes:
                continue

            # 单个客户代码处理
            if len(valid_codes) == 1:
                new_row = row.copy()
                new_row['客户代码'] = valid_codes[0]
                new_row['客户名称'] = customer_dict.get(valid_codes[0], '')
                df_expanded_list.append(new_row)
                continue

            # 多个客户代码需要分配剩余生产量
            total_remaining = row['剩余生产量']
            customer_data = []

            # 获取每个客户的计划数
            for code in valid_codes:
                customer_name = customer_dict.get(code, '')
                plan_value = 0
                if customer_name and row['产品型号']:
                    match = plan_df[
                        (plan_df['客户名称'] == customer_name) &
                        (plan_df['产品型号'] == row['产品型号'])
                        ]
                    if not match.empty:
                        plan_value = match['计划数'].iloc[0]
                customer_data.append({
                    '客户代码': code,
                    '客户名称': customer_name,
                    '计划数': plan_value
                })

            # 新分配逻辑：按顺序分配给计划数>0的客户，最后未分配完的给最后一个客户
            allocations = [0] * len(customer_data)  # 初始化分配列表
            allocated = 0  # 已分配量

            # 按顺序遍历客户
            for i, cust in enumerate(customer_data):
                # 如果还有剩余量且当前客户有计划数
                if allocated < total_remaining and cust['计划数'] > 0:
                    # 可分配量 = min(客户计划数, 剩余总量 - 已分配量)
                    alloc_amount = min(cust['计划数'], total_remaining - allocated)
                    allocations[i] = alloc_amount
                    allocated += alloc_amount

            # 如果还有剩余未分配，分配给最后一个客户
            if allocated < total_remaining:
                # 计算剩余量
                remaining = total_remaining - allocated
                # 分配给最后一个客户
                allocations[-1] += remaining
                allocated += remaining

            # 记录分配详情（仅多客户情况）
            for i, cust in enumerate(customer_data):
                allocation_details.append({
                    '配方编号': formula_code,
                    '产品型号': row['产品型号'],
                    '客户代码': cust['客户代码'],
                    '客户名称': cust['客户名称'],
                    '分配数量': allocations[i]
                })

            # 创建拆分后的行
            for i, cust in enumerate(customer_data):
                new_row = row.copy()
                new_row['客户代码'] = cust['客户代码']
                new_row['客户名称'] = cust['客户名称']
                new_row['剩余生产量'] = allocations[i]
                df_expanded_list.append(new_row)

        # 合并所有行
        if df_expanded_list:
            df_expanded = pd.DataFrame(df_expanded_list)
        else:
            df_expanded = pd.DataFrame(columns=df.columns.tolist() + ['剩余生产量', '客户名称'])

        # 分组汇总剩余生产量
        grouped = df_expanded.groupby(['客户代码', '客户名称', '产品型号'], as_index=False).agg(
            totalproding=('剩余生产量', 'sum')
        )

        # 将分配详情转换为DataFrame
        allocation_df = pd.DataFrame(allocation_details)

        return grouped, allocation_df

    except Exception as e:
        print(f"处理文件出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame(columns=['客户代码', '客户名称', '产品型号', 'totalproding']), pd.DataFrame(columns=['配方编号', '产品型号', '客户代码', '客户名称', '分配数量'])

def update_production_trace(sales_df, production_file_path, sheetname,nextmonthsheetname, nm_prodingfile, hb_prodingfile,
                            NM_data_storage, HB_data_storage, nm_data_prodsend, hb_data_prodsend,
                            NM_all_storage, HB_all_storage,NM_return_data,HB_return_data):
    """
    更新生产跟踪文件并生成新版本（支持多列更新）
    :param sales_df: 销售数据DataFrame（需包含客户、产品型号、当月列、内蒙计划数、湖北计划数）
    :return: 新生成的文件路径
    """
    # 生成新文件名
    new_file_path = generate_new_filename(production_file_path)
    smart_decimal_format = '#,##0.###'

    try:
        # 复制模板文件
        shutil.copyfile(production_file_path, new_file_path)

        # 加载新文件的工作表
        wb = load_workbook(filename=new_file_path)
        if sheetname not in wb.sheetnames:
            raise ValueError(f"工作表 {sheetname} 不存在")
        ws = wb[sheetname]

        # 2. 处理下月工作表
        if nextmonthsheetname not in wb.sheetnames:
            raise ValueError(f"工作表 {nextmonthsheetname} 不存在")
        ws_next = wb[nextmonthsheetname]

        # 处理内蒙配比单文件
        nm_proding_data, nm_allocation_details = parse_rateorder_file(nm_prodingfile, new_file_path)

        # 处理湖北配比单文件
        hb_proding_data, hb_allocation_details = parse_rateorder_file(hb_prodingfile, new_file_path)

        # 提取当月列名和下月列名
        current_month_col = sales_df.columns[2]  # 当月销售计划列名
        next_month_col = sales_df.columns[5]  # 下月销售计划列名
        next_month_nm_col = sales_df.columns[6]  # 下月内蒙计划数列名
        next_month_hb_col = sales_df.columns[7]  # 下月湖北计划数列名

        # 构建销售数据索引
        # sales_dict = sales_df.set_index(['客户', '产品型号'])[
        #     [sales_df.columns[2], '内蒙计划数', '湖北计划数']].to_dict(orient='index')
        sales_dict = sales_df.set_index(['客户', '产品型号'])[
            [current_month_col, '内蒙计划数', '湖北计划数',
             next_month_col, next_month_nm_col, next_month_hb_col]
        ].to_dict(orient='index')

        # 构建配比单数据索引
        nm_proding_dict = nm_proding_data.set_index(['客户名称', '产品型号'])['totalproding'].to_dict()
        hb_proding_dict = hb_proding_data.set_index(['客户名称', '产品型号'])['totalproding'].to_dict()

        # 初始化配比单未匹配集合
        nm_unmatched_rate = set(nm_proding_dict.keys())
        hb_unmatched_rate = set(hb_proding_dict.keys())

        # 初始化其他未匹配数据容器
        nm_unmatched = NM_data_storage[['客户', '产品型号', '库存']].to_dict('records')
        hb_unmatched = HB_data_storage[['客户', '产品型号', '库存']].to_dict('records')
        nm_send_unmatched = nm_data_prodsend[['客户', '产品型号', '内蒙发货数']].to_dict('records')
        hb_send_unmatched = hb_data_prodsend[['客户', '产品型号', '湖北发货数']].to_dict('records')

        # 转换为集合以便快速删除
        nm_unmatched_set = {(item['客户'], item['产品型号']) for item in nm_unmatched}
        hb_unmatched_set = {(item['客户'], item['产品型号']) for item in hb_unmatched}
        nm_send_unmatched_set = {(item['客户'], item['产品型号']) for item in nm_send_unmatched}
        hb_send_unmatched_set = {(item['客户'], item['产品型号']) for item in hb_send_unmatched}

        # 构建索引字典
        nm_stock = NM_data_storage.drop_duplicates(subset=['客户', '产品型号'], keep='last')
        nm_stock_dict = nm_stock.set_index(['客户', '产品型号']).to_dict('index')

        hb_stock = HB_data_storage.drop_duplicates(subset=['客户', '产品型号'], keep='last')
        hb_stock_dict = hb_stock.set_index(['客户', '产品型号']).to_dict('index')

        nm_send = nm_data_prodsend.drop_duplicates(subset=['客户', '产品型号'], keep='last')
        nm_send_dict = nm_send.set_index(['客户', '产品型号']).to_dict('index')

        hb_send = hb_data_prodsend.drop_duplicates(subset=['客户', '产品型号'], keep='last')
        hb_send_dict = hb_send.set_index(['客户', '产品型号']).to_dict('index')

        # 遍历跟踪数据 处理当月工作表
        updated_rows = 0
        for row_idx, row in enumerate(ws.iter_rows(min_row=4), start=4):
            customer = row[0].value  # A列 - 客户名称
            product = row[1].value  # B列 - 产品型号

            # 终止条件：A/B列同时为空
            if customer is None and product is None:
                break

            key = (customer, product)

            # 匹配销售数据 - 列索引调整（新增包装方式列）
            if key in sales_dict:
                # 获取三列数据
                month_value = sales_dict[key][current_month_col]
                nm_plan = sales_dict[key]['内蒙计划数']
                hb_plan = sales_dict[key]['湖北计划数']

                # 更新三列（D=4, E=5, F=6）
                # D列：6月销售计划
                ws.cell(row=row_idx, column=4, value=month_value)

                # E列：内蒙计划数
                ws.cell(row=row_idx, column=5, value=nm_plan).number_format = '0'

                # F列：湖北计划数
                ws.cell(row=row_idx, column=6, value=hb_plan).number_format = '0'

                # 设置格式
                for col in [4, 5, 6]:
                    ws.cell(row=row_idx, column=col).alignment = Alignment(horizontal='center', vertical='center')
                    ws.cell(row=row_idx, column=col).font = Font(bold=True, size=10)

                updated_rows += 1

            # 处理内蒙发货 - 列索引调整（H=8）
            if key in nm_send_dict:
                ws.cell(row=row_idx, column=8,
                        value=nm_send_dict[key]['内蒙发货数']).number_format = smart_decimal_format
                ws.cell(row=row_idx, column=8).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=row_idx, column=8).font = Font(bold=True, size=10)
                if key in nm_send_unmatched_set:
                    nm_send_unmatched_set.remove(key)

            # 处理湖北发货 - 列索引调整（I=9）
            if key in hb_send_dict:
                ws.cell(row=row_idx, column=9,
                        value=hb_send_dict[key]['湖北发货数']).number_format = smart_decimal_format
                ws.cell(row=row_idx, column=9).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=row_idx, column=9).font = Font(bold=True, size=10)
                if key in hb_send_unmatched_set:
                    hb_send_unmatched_set.remove(key)

            # 处理内蒙库存 - 列索引调整（K=11）
            if key in nm_stock_dict:
                ws.cell(row=row_idx, column=11, value=nm_stock_dict[key]['库存']).number_format = smart_decimal_format
                ws.cell(row=row_idx, column=11).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=row_idx, column=11).font = Font(bold=True, size=10)
                if key in nm_unmatched_set:
                    nm_unmatched_set.remove(key)

            # 处理湖北库存 - 列索引调整（L=12）
            if key in hb_stock_dict:
                ws.cell(row=row_idx, column=12, value=hb_stock_dict[key]['库存']).number_format = smart_decimal_format
                ws.cell(row=row_idx, column=12).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=row_idx, column=12).font = Font(bold=True, size=10)
                if key in hb_unmatched_set:
                    hb_unmatched_set.remove(key)

            # 处理配比单生产数据
            rate_key = key
            if product == "SN-LTF":
                rate_key = (customer, "SN-P2C-1")

            # N列(14列) - 内蒙配比单生产汇总
            if rate_key in nm_proding_dict:
                ws.cell(row=row_idx, column=14, value=nm_proding_dict[rate_key]).number_format = smart_decimal_format
                ws.cell(row=row_idx, column=14).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=row_idx, column=14).font = Font(bold=True, size=10)
                # 从配比单未匹配集合中移除已匹配的项
                if rate_key in nm_unmatched_rate:
                    nm_unmatched_rate.remove(rate_key)

            # O列(15列) - 湖北配比单生产汇总
            if rate_key in hb_proding_dict:
                ws.cell(row=row_idx, column=15, value=hb_proding_dict[rate_key]).number_format = smart_decimal_format
                ws.cell(row=row_idx, column=15).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=row_idx, column=15).font = Font(bold=True, size=10)
                # 从配比单未匹配集合中移除已匹配的项
                if rate_key in hb_unmatched_rate:
                    hb_unmatched_rate.remove(rate_key)

        # 2. 收集当月工作表的数据用于下月库存计算
        prev_sheet_data = {}
        for row_idx, row in enumerate(ws.iter_rows(min_row=4), start=4):
            customer = row[0].value
            product = row[1].value
            if customer is None and product is None:
                break

            def to_float(value):
                if value is None:
                    return 0.0
                try:
                    return float(value)
                except (TypeError, ValueError):
                    return 0.0

            # # 读取内蒙相关列
            # nm_plan = row[4].value or 0  # E列 - 内蒙计划数
            # nm_sent = row[7].value or 0  # H列 - 内蒙发货数
            # nm_stock = row[10].value or 0  # K列 - 内蒙库存
            # nm_rate = row[13].value or 0  # N列 - 内蒙配比单生产
            # nm_transit = row[16].value or 0  # Q列 - 内蒙在途
            #
            # # 读取湖北相关列
            # hb_plan = row[5].value or 0  # F列 - 湖北计划数
            # hb_sent = row[8].value or 0  # I列 - 湖北发货数
            # hb_stock = row[11].value or 0  # L列 - 湖北库存
            # hb_rate = row[14].value or 0  # O列 - 湖北配比单生产
            # hb_transit = row[15].value or 0  # P列 - 湖北在途

            # 读取内蒙相关列
            nm_plan_val = to_float(row[4].value)  # E列 - 内蒙计划数
            nm_sent_val = to_float(row[7].value)  # H列 - 内蒙发货数
            nm_stock_val = to_float(row[10].value)  # K列 - 内蒙库存
            nm_rate_val = to_float(row[13].value)  # N列 - 内蒙配比单生产
            nm_transit_val = to_float(row[16].value)  # Q列 - 内蒙在途

            # 读取湖北相关列
            hb_plan_val = to_float(row[5].value)  # F列 - 湖北计划数
            hb_sent_val = to_float(row[8].value)  # I列 - 湖北发货数
            hb_stock_val = to_float(row[11].value)  # L列 - 湖北库存
            hb_rate_val = to_float(row[14].value)  # O列 - 湖北配比单生产
            hb_transit_val = to_float(row[15].value)  # P列 - 湖北在途

            prev_sheet_data[(customer, product)] = {
                'nm_plan': nm_plan_val,
                'nm_sent': nm_sent_val,
                'nm_stock': nm_stock_val,
                'nm_rate': nm_rate_val,
                'nm_transit': nm_transit_val,
                'hb_plan': hb_plan_val,
                'hb_sent': hb_sent_val,
                'hb_stock': hb_stock_val,
                'hb_rate': hb_rate_val,
                'hb_transit': hb_transit_val
            }

        # ===== 处理下月工作表 =====
        updated_rows_next = 0
        for row_idx, row in enumerate(ws_next.iter_rows(min_row=4), start=4):
            customer = row[0].value  # A列 - 客户名称
            product = row[1].value  # B列 - 产品型号

            # 终止条件：A/B列同时为空
            if customer is None and product is None:
                break

            key = (customer, product)

            # 匹配销售数据 - 只写入计划数据
            if key in sales_dict:
                # 获取三列数据（下月）
                next_month_value = sales_dict[key][next_month_col]
                next_month_nm = sales_dict[key][next_month_nm_col]
                next_month_hb = sales_dict[key][next_month_hb_col]

                # 更新三列（D=4, E=5, F=6）
                # D列：下月销售计划
                ws_next.cell(row=row_idx, column=4, value=next_month_value)

                # E列：下月内蒙计划数
                ws_next.cell(row=row_idx, column=5, value=next_month_nm).number_format = '0'

                # F列：下月湖北计划数
                ws_next.cell(row=row_idx, column=6, value=next_month_hb).number_format = '0'

                # 设置格式
                for col in [4, 5, 6]:
                    ws_next.cell(row=row_idx, column=col).alignment = Alignment(horizontal='center',
                                                                                vertical='center')
                    ws_next.cell(row=row_idx, column=col).font = Font(bold=True, size=10)

                updated_rows_next += 1

            if key in prev_sheet_data:
                data = prev_sheet_data[key]

                # 计算内蒙库存 (K列)
                nm_total = (data['nm_stock'] +
                            data['nm_rate'] +
                            data['nm_transit'] -
                            data['nm_plan'] +
                            data['nm_sent'])
                nm_total = max(0, nm_total)  # 小于0时取0

                # 计算湖北库存 (L列)
                hb_total = (data['hb_stock'] +
                            data['hb_rate'] +
                            data['hb_transit'] -
                            data['hb_plan'] +
                            data['hb_sent'])
                hb_total = max(0, hb_total)  # 小于0时取0

                # 写入计算结果
                ws_next.cell(row=row_idx, column=11, value=nm_total).number_format = smart_decimal_format
                ws_next.cell(row=row_idx, column=12, value=hb_total).number_format = smart_decimal_format

                # 设置格式
                for col in [11, 12]:
                    cell = ws_next.cell(row=row_idx, column=col)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(bold=True, size=10)


        # 更新未匹配数据列表
        nm_unmatched = [item for item in nm_unmatched if (item['客户'], item['产品型号']) in nm_unmatched_set]
        hb_unmatched = [item for item in hb_unmatched if (item['客户'], item['产品型号']) in hb_unmatched_set]
        nm_send_unmatched = [item for item in nm_send_unmatched if
                             (item['客户'], item['产品型号']) in nm_send_unmatched_set]
        hb_send_unmatched = [item for item in hb_send_unmatched if
                             (item['客户'], item['产品型号']) in hb_send_unmatched_set]

        # 准备配比单未匹配数据
        nm_rate_unmatched = [
            {'客户': key[0], '产品型号': key[1], '剩余生产量': value}
            for key, value in nm_proding_dict.items()
            if key in nm_unmatched_rate
        ]

        hb_rate_unmatched = [
            {'客户': key[0], '产品型号': key[1], '剩余生产量': value}
            for key, value in hb_proding_dict.items()
            if key in hb_unmatched_rate
        ]

        # 写入未匹配数据到新sheet（带格式设置）
        def write_unmatched_sheet(wb, data, sheet_name, columns):

            if not data:
                print(f"未匹配数据 '{sheet_name}' 为空，跳过创建sheet")
                return

            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws = wb.create_sheet(sheet_name)
            ws.append(columns)

            # 确定数值列索引
            numeric_columns = ['库存', '内蒙发货数', '湖北发货数','库存（吨）','分配数量']
            numeric_indices = [i for i, col in enumerate(columns) if col in numeric_columns]

            for item in data:
                row = [item.get(col, 0) for col in columns]
                ws.append(row)

                # 设置数值列格式
                for col_idx in numeric_indices:
                    cell = ws.cell(row=ws.max_row, column=col_idx + 1)  # openpyxl列从1开始
                    cell.number_format = smart_decimal_format

        def write_allocation_sheet(wb, details_df, sheet_name):
            if not details_df.empty:
                if sheet_name in wb.sheetnames:
                    del wb[sheet_name]
                ws = wb.create_sheet(sheet_name)

                # 添加标题
                headers = ['配方编号', '产品型号', '客户代码', '客户名称', '分配数量(吨)']
                ws.append(headers)

                # 写入数据
                for _, row in details_df.iterrows():
                    ws.append([
                        row['配方编号'],
                        row['产品型号'],
                        row['客户代码'],
                        row['客户名称'],
                        row['分配数量']
                    ])

                # 设置格式
                for row_idx in range(1, ws.max_row + 1):
                    for col in range(1, 6):
                        cell = ws.cell(row=row_idx, column=col)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.font = Font(size=10)
                        if row_idx > 1 and col == 5:  # 分配数量列
                            cell.number_format = '#,##0.###'

                # 设置列宽
                col_widths = [15, 20, 15, 20, 15]
                for idx, width in enumerate(col_widths, 1):
                    ws.column_dimensions[get_column_letter(idx)].width = width

        # 1. 内蒙未定归属库存（客户为空的库存）
        nm_undetermined = NM_all_storage[NM_all_storage['客户'].str.strip() == ''].copy()
        print("nm_undetermined",nm_undetermined)
        if not nm_undetermined.empty:
            # 添加类型列以便识别
            nm_undetermined['类型'] = '未定归属库存'
            nm_undetermined = nm_undetermined[['产品型号', '库存（吨）']]
            write_unmatched_sheet(wb, nm_undetermined.to_dict('records'),
                                  '内蒙未定归属库存', ['产品型号', '库存（吨）'])

        # 2. 湖北未定归属库存（客户为空的库存）
        hb_undetermined = HB_all_storage[HB_all_storage['客户'].str.strip() == ''].copy()

        print("hb_undetermined", hb_undetermined)
        if not hb_undetermined.empty:
            hb_undetermined['类型'] = '未定归属库存'
            hb_undetermined = hb_undetermined[['产品型号', '库存（吨）']]
            write_unmatched_sheet(wb, hb_undetermined.to_dict('records'),
                                  '湖北未定归属库存', ['产品型号', '库存（吨）'])



        # 写入所有未匹配数据到新sheet
        write_unmatched_sheet(wb, nm_unmatched, '内蒙未在销售计划库存', ['客户', '产品型号', '库存'])
        write_unmatched_sheet(wb, hb_unmatched, '湖北未在销售计划库存', ['客户', '产品型号', '库存'])
        write_unmatched_sheet(wb, nm_send_unmatched, '内蒙需要确认发货', ['客户', '产品型号', '内蒙发货数'])
        write_unmatched_sheet(wb, hb_send_unmatched, '湖北需要确认发货', ['客户', '产品型号', '湖北发货数'])
        write_unmatched_sheet(wb, nm_rate_unmatched, '内蒙配比单未处理', ['客户', '产品型号', '剩余生产量'])
        write_unmatched_sheet(wb, hb_rate_unmatched, '湖北配比单未处理', ['客户', '产品型号', '剩余生产量'])
        print("nm_allocation_details",nm_allocation_details)
        write_allocation_sheet(wb, nm_allocation_details, '内蒙多客户配比单分配统计')
        write_allocation_sheet(wb, hb_allocation_details, '湖北多客户配比单分配统计')

        # 3. 内蒙特殊库存
        if not NM_return_data.empty:
            print("NM_return_data", NM_return_data)
            # 添加类型列以便识别
            NM_return_data['类型'] = '特殊库存'
            # 只保留需要的列
            nm_special_cols = [col for col in NM_return_data.columns if col != '类型']
            nm_special_cols.insert(0, '类型')  # 将类型列放在第一列
            write_unmatched_sheet(wb, NM_return_data.to_dict('records'),
                                  '内蒙特殊库存', nm_special_cols)

        # 4. 湖北特殊库存
        if not HB_return_data.empty:
            print("HB_return_data2",HB_return_data)
            HB_return_data['类型'] = '特殊库存'
            hb_special_cols = [col for col in HB_return_data.columns if col != '类型']
            hb_special_cols.insert(0, '类型')  # 将类型列放在第一列
            write_unmatched_sheet(wb, HB_return_data.to_dict('records'),
                                  '湖北特殊库存', hb_special_cols)

        # 保存修改
        wb.save(new_file_path)
        print(f"成功更新 {updated_rows} 行数据到: {new_file_path}")
        return new_file_path

    except KeyError as e:
        if "内蒙计划数" in str(e) or "湖北计划数" in str(e):
            raise ValueError("sales_df必须包含'内蒙计划数'和'湖北计划数'列") from e
        else:
            raise
    except Exception as e:
        # 清理生成失败的文件
        if os.path.exists(new_file_path):
            os.remove(new_file_path)
        raise RuntimeError(f"文件更新失败: {str(e)}")





def resplitsalesplan(isCreate, original_file_path):
    if isCreate != 1:
        return
    smart_decimal_format = '#,##0.###'
    # 加载工作簿
    wb = openpyxl.load_workbook(original_file_path)

    # 检查工作表是否存在
    if "2025年销售计划" not in wb.sheetnames:
        print(f"错误：工作簿中不存在名为'2025年销售计划'的工作表")
        return

    # 获取原始工作表
    original_sheet = wb["2025年销售计划"]

    # 添加新列
    last_col = original_sheet.max_column
    original_sheet.cell(row=2, column=last_col + 1, value="内蒙")
    original_sheet.cell(row=2, column=last_col + 2, value="湖北")
    original_sheet.cell(row=2, column=last_col + 3, value="标识")

    # 设置标识列的值（注意Excel行号从1开始）
    flag_rows = [8, 33]
    for row in flag_rows:
        if row <= original_sheet.max_row:
            original_sheet.cell(row=row, column=last_col + 3).value = 1

    # 获取当前月份（6月）
    current_month = datetime.now().month
    month_col = None
    for col in range(6, last_col + 1):  # 从F列开始查找
        cell_value = original_sheet.cell(row=2, column=col).value
        if cell_value and str(current_month) + "月" in str(cell_value):
            month_col = col
            break

    if month_col is None:
        print("错误：未找到当月列（1-12月）")
        return

    # 填充内蒙和湖北列
    for row in range(3, original_sheet.max_row + 1):
        customer_name = original_sheet.cell(row=row, column=3).value
        # 检查是否为合计行
        if customer_name and "合计" in str(customer_name):
            continue
        flag_cell = original_sheet.cell(row=row, column=last_col + 3)
        month_value = original_sheet.cell(row=row, column=month_col).value

        original_sheet.cell(row=row, column=last_col + 3).number_format = smart_decimal_format

        if flag_cell.value == 1:
            # 复制到湖北列
            original_sheet.cell(row=row, column=last_col + 2).value = month_value
            original_sheet.cell(row=row, column=last_col + 2).number_format = smart_decimal_format
        else:
            # 复制到内蒙列
            original_sheet.cell(row=row, column=last_col + 1).value = month_value
            original_sheet.cell(row=row, column=last_col + 1).number_format = smart_decimal_format

    # 复制工作表（深拷贝样式）
    # 创建新工作表并复制所有内容
    def copy_sheet_with_style(source, target_name):
        target = wb.create_sheet(target_name)
        for row in source.iter_rows():
            for cell in row:
                new_cell = target.cell(
                    row=cell.row,
                    column=cell.column,
                    value=cell.value
                )
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = smart_decimal_format
                    new_cell.alignment = copy(cell.alignment)
        # 复制列宽
        for col in range(1, source.max_column + 1):
            col_letter = get_column_letter(col)
            # 确保源列维度存在
            if col_letter in source.column_dimensions:
                source_dim = source.column_dimensions[col_letter]

                # 确保目标列维度存在
                if col_letter not in target.column_dimensions:
                    target.column_dimensions[col_letter] = openpyxl.worksheet.dimensions.ColumnDimension(
                        target, index=col_letter
                    )

                # 复制宽度
                if hasattr(source_dim, 'width') and source_dim.width is not None:
                    target.column_dimensions[col_letter].width = source_dim.width
            #target.column_dimensions[col_letter].width = source.column_dimensions[col_letter].width
        return target

    # 创建工作表副本
    neimeng_sheet = copy_sheet_with_style(original_sheet, "2025年内蒙销售计划")
    hubei_sheet = copy_sheet_with_style(original_sheet, "2025年湖北销售计划")

    # === 隐藏E列与当月列之间的列 ===
    # E列是第5列，当月列是month_col
    if month_col > 6:  # 确保有列需要隐藏
        start_hide_col = 6  # F列开始
        end_hide_col = month_col - 1  # 当月列的前一列

        for sheet in [neimeng_sheet, hubei_sheet]:
            for col in range(start_hide_col, end_hide_col + 1):
                col_letter = get_column_letter(col)
                # 确保列维度对象存在
                if col_letter not in sheet.column_dimensions:
                    # 如果不存在则创建
                    sheet.column_dimensions[col_letter] = openpyxl.worksheet.dimensions.ColumnDimension(
                        sheet, index=col_letter
                    )
                sheet.column_dimensions[col_letter].hidden = True

    # 处理内蒙销售计划
    neimeng_sheet = wb["2025年内蒙销售计划"]

    # 查找最后一个"合计"行
    last_total_row = None
    for row in range(neimeng_sheet.max_row, 2, -1):
        if neimeng_sheet.cell(row=row, column=3).value == "合计":
            last_total_row = row
            break

    if last_total_row is None:
        last_total_row = neimeng_sheet.max_row

    # 从高到低删除标识为1的行（在3到合计行之间）
    flag_rows_to_delete = [row for row in range(3, last_total_row + 1)
                           if neimeng_sheet.cell(row=row, column=last_col + 3).value == 1]

    # 按从大到小排序
    flag_rows_to_delete.sort(reverse=True)

    for row in flag_rows_to_delete:
        neimeng_sheet.delete_rows(row)

    # 处理湖北销售计划
    hubei_sheet = wb["2025年湖北销售计划"]

    # 查找最后一个"合计"行
    last_total_row = None
    for row in range(hubei_sheet.max_row, 2, -1):
        if hubei_sheet.cell(row=row, column=3).value == "合计":
            last_total_row = row
            break

    if last_total_row is None:
        last_total_row = hubei_sheet.max_row

    # 从高到低删除非目标行（3到合计行之间）
    rows_to_delete = []
    for row in range(3, last_total_row + 1):
        flag_value = hubei_sheet.cell(row=row, column=last_col + 3).value
        customer_value = str(hubei_sheet.cell(row=row, column=3).value)

        # 保留条件：标识为1 或 包含"合计"
        if flag_value != 1 and "合计" not in customer_value:
            rows_to_delete.append(row)

    # 按从大到小排序
    rows_to_delete.sort(reverse=True)

    for row in rows_to_delete:
        hubei_sheet.delete_rows(row)

    # 保存修改
    wb.save(original_file_path)

    print("操作成功完成！")
    return 1

# 示例调用
resplitsalesplan(1, "销售2025年销售计划6.13-生产1.xlsx")