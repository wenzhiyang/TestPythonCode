# PLATFORM-SAFE: This file contains no executable risks
# PURPOSE: Educational demonstration only

"""
本文件仅用于教育目的
不包含真实凭证或恶意功能
所有敏感操作均已移除或禁用
"""




def GetSalesPlanInit(filePath, sheetname):
    # 销售给的原始文件,清洗合计行、按成品合并销售月计划，并排序
    # 加载工作簿和指定sheet
    wb = load_workbook(filePath, data_only=True)

    # 验证sheet是否存在
    if sheetname not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheetname}' 不存在于文件中")
    ws = wb[sheetname]



    # 确定有效列数：第二行（行号2）中从A列开始直到第一个空列
    max_col = 0
    for cell in ws[2]:  # 第二行
        if cell.value is None:
            break
        max_col += 1

    # 检查是否包含必要的列（D列和F列）
    if max_col < 6:
        raise ValueError("有效列不足，未包含F列（第6列）")

    # 提取第二行的列名（A列到有效列末尾）
    columns = [cell.value for cell in ws[2][:max_col]]

    data = []
    # 从第三行开始读取数据
    for row in ws.iter_rows(min_row=3):
        # 检查停止条件
        a_val = row[0].value if len(row) > 0 else None
        b_val = row[1].value if len(row) > 1 else None
        c_val = row[2].value if len(row) > 2 else None

        # 判断是否满足停止条件
        stop_condition1 = (a_val in (None, '')) and (b_val in (None, '')) and (c_val in (None, ''))
        stop_condition2 = (a_val in (None, '') and (b_val in (None, '')) and c_val == '代加工客户销售计划')
        if stop_condition1 or stop_condition2:
            break

        # 跳过包含“合计”的行
        if c_val and '合计' in str(c_val):
            continue

        # 处理当前行数据
        row_data = []
        for col_idx in range(max_col):
            cell = row[col_idx] if col_idx < len(row) else None
            cell_value = cell.value if cell is not None else None

            # 修改D列（索引3）的值
            if col_idx == 3 and cell_value == 'SN-LTF':
                cell_value = 'SN-P2C-1'

            row_data.append(cell_value)

        data.append(row_data)

    # 转换为DataFrame
    df = pd.DataFrame(data, columns=columns)
    original_columns = df.columns.tolist()  # 获取原始列顺序
    month_columns = [col for col in original_columns if re.match(r'^\d+月$', col)]
    if df.empty:
        return pd.DataFrame()

    # === 新增逻辑：自动补充未来3个月 ===
    # if month_columns:
    #     month_numbers = sorted([int(col[:-1]) for col in month_columns])
    #     max_month_num = month_numbers[-1]
    #     max_month_col = f"{max_month_num}月"
    #
    #     next_months = []
    #     current_year = datetime.now().year
    #     for i in range(1, 4):
    #         new_month_num = (max_month_num + i - 1) % 12 + 1
    #         new_year = current_year + (max_month_num + i - 1) // 12
    #         next_months.append(f"{new_month_num}月")
    #
    #     for m in next_months:
    #         if m not in df.columns:
    #             df[m] = df[max_month_col]
    #             month_columns.append(m)

    # 自定义排序规则
    custom_order = [
        'SN-P1C', 'SN-P2C', 'SN-P1L', 'SN-P1G', 'SN-P2C-1', 'SN-4D-1',
        'SN-H01', 'SN-H04', 'MAG-4CL', 'SN-P2H', 'MAG-106A', 'SN-10BL',
        'MAG-09', 'SN-09G', 'MAG-106', 'SN-P2', 'SN-BP12T', 'SN-BP12',
        'SN-BN1', 'SN-K1-3T', 'SN-C1M-A', 'SN-P1', 'SN-K1M', 'SN-P2C-2',
        'SN-P2C-GX'
    ]

    # 按D列分类排序
    df[columns[3]] = pd.Categorical(
        df[columns[3]],
        categories=custom_order,
        ordered=True
    )

    # 分组聚合
    group_col = columns[3]
    target_column = '每月目标安全库存'
    agg_columns = month_columns + [target_column]

    # 数值类型转换
    for col in agg_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 分组求和
    grouped_df = df.groupby(group_col, observed=False)[agg_columns].sum().reset_index()

    # 按自定义顺序排序
    grouped_df = grouped_df.sort_values(by=group_col)

    return grouped_df




sales_df_all = GetSalesPlanInit("销售2025年销售计划6.6-生产.xlsx","2025年销售计划")

print("sales_df_all", sales_df_all[['产品型号','6月']])

