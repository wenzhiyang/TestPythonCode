# PLATFORM-SAFE: This file contains no executable risks
# PURPOSE: Educational demonstration only

"""
本文件仅用于教育目的
不包含真实凭证或恶意功能
所有敏感操作均已移除或禁用
"""




def get_prod_send():
    # 新增客户名称替换函数
    def update_customer_names(df):
        # 湖州天能 → 天能
        mask_tianneng = df['客户'] == '湖州天能'
        df.loc[mask_tianneng, '客户'] = '天能'

        # 天津力神 + SN-K1M → 力神
        mask_lishen_tj = (df['客户'] == '天津力神') & (df['产品型号'] == 'SN-K1M')
        df.loc[mask_lishen_tj, '客户'] = '力神'

        # 滁州星恒 → 星恒
        mask_xingheng_cz = df['客户'] == '滁州星恒'
        df.loc[mask_xingheng_cz, '客户'] = '星恒'

        mask_xingheng_cz = df['客户'].str.contains('国轩', na=False)
        df.loc[mask_xingheng_cz, '客户'] = '国轩'

        # 天鹏锂能 → 天鹏
        mask_tianpeng = df['客户'] == '天鹏锂能'
        df.loc[mask_tianpeng, '客户'] = '天鹏'

        # 南通拓邦/惠州拓邦 → 拓邦
        mask_tuobang1 = df['客户'] == '南通拓邦'
        df.loc[mask_tuobang1, '客户'] = '拓邦'

        mask_tuobang2 = df['客户'] == '惠州拓邦'
        df.loc[mask_tuobang2, '客户'] = '拓邦'

        # 贵阳比亚迪 → 比亚迪
        mask_byd = df['客户'] == '贵阳比亚迪'
        df.loc[mask_byd, '客户'] = '比亚迪'

        # 天津聚元 → 聚元
        mask_juyuan = df['客户'] == '天津聚元'
        df.loc[mask_juyuan, '客户'] = '力神聚元'

        # 东莞力鹏 → 力鹏
        mask_lipeng = df['客户'] == '东莞力鹏'
        df.loc[mask_lipeng, '客户'] = '力鹏'

        # 时代华景 → 华景
        mask_huajing = df['客户'] == '时代华景'
        df.loc[mask_huajing, '客户'] = '华景'

        mask_tianyi = df['客户'] == '芜湖天弋'
        df.loc[mask_tianyi, '客户'] = '天弋'

        # 苏州星恒 → 星恒
        mask_xingheng_sz = df['客户'] == '苏州星恒'
        df.loc[mask_xingheng_sz, '客户'] = '星恒'

        # 深圳宇宙探索 → 宇宙探索
        mask_yuzhou = df['客户'] == '深圳宇宙探索'
        df.loc[mask_yuzhou, '客户'] = '宇宙'

        mask_yuzhou = df['客户'] == '宇宙探索'
        df.loc[mask_yuzhou, '客户'] = '宇宙'

        # 大连中比 → 中比
        mask_zhongbi = df['客户'] == '大连中比'
        df.loc[mask_zhongbi, '客户'] = '中比'

        # 深圳塞恩士 → 塞恩士
        mask_saienshi = df['客户'] == '深圳塞恩士'
        df.loc[mask_saienshi, '客户'] = '赛恩士'

        mask_saienshi = df['客户'] == '深圳赛恩士'
        df.loc[mask_saienshi, '客户'] = '赛恩士'

        # 东莞能优 → 能优
        mask_nengyou = df['客户'] == '东莞能优'
        df.loc[mask_nengyou, '客户'] = '能优'

        # 无锡力神 + SN-K1M → 力神
        mask_lishen_wx = (df['客户'] == '无锡力神') & (df['产品型号'] == 'SN-K1M')
        df.loc[mask_lishen_wx, '客户'] = '力神'

        mask_nengyou = (df['客户'] == '朗泰通') & (df['产品型号'] == 'SN-P2C-1')
        df.loc[mask_nengyou, '产品型号'] = 'SN-LTF'

        return df

    query = """
       select Model as '产品型号',customer as '客户',SUM(CASE WHEN MONTH(saleDate) = 1 THEN Sendnum ELSE 0 END)/1000 AS '1月发货数',SUM(CASE WHEN MONTH(saleDate) = 2 THEN Sendnum ELSE 0 END)/1000 AS '2月发货数',SUM(CASE WHEN MONTH(saleDate) = 3 THEN Sendnum ELSE 0 END)/1000 AS '3月发货数',SUM(CASE WHEN MONTH(saleDate) = 4 THEN Sendnum ELSE 0 END)/1000 AS '4月发货数',SUM(CASE WHEN MONTH(saleDate) = 5 THEN Sendnum ELSE 0 END)/1000 AS '5月发货数' from ProductionSend_temp ps where orderstatus != '退货'
group by Model,customer
HAVING SUM(Sendnum) > 0
       """
    # 执行查询
    cursor.execute(query)

    # 获取查询结果
    results = cursor.fetchall()

    # # 获取列名
    # columns = ['产品型号', '客户', '内蒙发货数']
    # # 创建 DataFrame
    # prodsend_data_nm = pd.DataFrame(results, columns=columns)
    prodsend_data_nm = pd.DataFrame(results, columns=['产品型号', '客户', '1月发货数', '2月发货数', '3月发货数', '4月发货数', '5月发货数'])

    # 应用客户名称替换
    prodsend_data_nm = update_customer_names(prodsend_data_nm)

    prodsend_data_nm = prodsend_data_nm.groupby(
        ['客户', '产品型号'],
        as_index=False
    )['1月发货数','2月发货数','3月发货数','4月发货数','5月发货数'].sum()

    query = """
           select Model as '产品型号',customer as '客户',SUM(CASE WHEN MONTH(saleDate) = 1 THEN Sendnum ELSE 0 END)/1000 AS '1月发货数',SUM(CASE WHEN MONTH(saleDate) = 2 THEN Sendnum ELSE 0 END)/1000 AS '2月发货数',SUM(CASE WHEN MONTH(saleDate) = 3 THEN Sendnum ELSE 0 END)/1000 AS '3月发货数',SUM(CASE WHEN MONTH(saleDate) = 4 THEN Sendnum ELSE 0 END)/1000 AS '4月发货数',SUM(CASE WHEN MONTH(saleDate) = 5 THEN Sendnum ELSE 0 END)/1000 AS '5月发货数' from HB_ProductionSend_temp ps 
    where orderstatus != '退货' group by Model,customer
    HAVING SUM(Sendnum) > 0
           """
    # 执行查询
    cursor.execute(query)

    # 获取查询结果
    results_hb = cursor.fetchall()

    # 获取列名
    # columns_hb = ['产品型号', '客户', '湖北发货数']
    # # 创建 DataFrame
    # prodsend_data_hb = pd.DataFrame(results_hb, columns=columns_hb)
    prodsend_data_hb = pd.DataFrame(results_hb, columns=['产品型号', '客户', '1月发货数', '2月发货数', '3月发货数', '4月发货数', '5月发货数'])

    # 应用客户名称替换
    prodsend_data_hb = update_customer_names(prodsend_data_hb)

    prodsend_data_hb = prodsend_data_hb.groupby(
        ['客户', '产品型号'],
        as_index=False
    )['1月发货数','2月发货数','3月发货数','4月发货数','5月发货数'].sum()

    return prodsend_data_nm, prodsend_data_hb

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

    def calculate_formula_value(cell):
        """处理公式单元格，尝试计算其值"""
        if cell.data_type == 'f':  # 公式类型
            try:
                # 尝试直接计算简单公式
                if cell.value.startswith('='):
                    expr = cell.value[1:]
                    # 安全计算数学表达式
                    return eval(expr, {"__builtins__": None}, {})
            except:
                pass
        return cell.value

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

            # 处理公式值
            if cell and cell.data_type == 'f':
                cell_value = calculate_formula_value(cell)

            # 修改D列（索引3）的值
            if col_idx == 3 and cell_value == 'SN-LTF':
                cell_value = 'SN-P2C-1'

            row_data.append(cell_value)

        data.append(row_data)

    # 转换为DataFrame
    df = pd.DataFrame(data, columns=columns)
    original_columns = df.columns.tolist()  # 获取原始列顺序
    month_columns = [col for col in original_columns if re.match(r'^\d+月$', col)]
    print("month_columns",month_columns)
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

