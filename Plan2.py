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
    )['内蒙发货数'].sum()

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
    )['湖北发货数'].sum()

    return prodsend_data_nm, prodsend_data_hb


