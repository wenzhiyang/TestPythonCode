# PLATFORM-SAFE: This file contains no executable risks
# PURPOSE: Educational demonstration only

"""
本文件仅用于教育目的
不包含真实凭证或恶意功能
所有敏感操作均已移除或禁用
"""



def getcalc():
    target_cols_71 = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD',
                      'AE', 'AF', 'AG',
                      'AH']

    # === 各工序数据计算 ===
    def calculate_product(row_numbers, operation=None):
        """通用乘积计算函数，支持特殊列处理"""
        total = 0.0

        for col_letter in target_cols_71:
            product = 1.0
            col_idx = column_index_from_string(col_letter)

            nindex = 0
            for row in row_numbers:
                # 特殊处理高温炭化工序的特殊列
                if operation == "高温炭化":
                    if col_letter in high_temp_special_cols:
                        # 获取相关列
                        main_col, dep_col = special_col_rules[col_letter]
                        main_col_idx = column_index_from_string(main_col)
                        dep_col_idx = column_index_from_string(dep_col)

                        # 读取主列和依赖列的值
                        main_val = safe_float(get_merged_value(sheet, row, main_col_idx))
                        dep_val = safe_float(get_merged_value(sheet, row, dep_col_idx))

                        if col_letter == "P" and (row == 75 or row == 68):
                            break
                        if col_letter == "R" and (row == 75 or row == 68):
                            break
                        if col_letter == "V" and (row == 75 or row == 68):
                            break
                        if col_letter == "Y" and (row == 75 or row == 68):
                            break
                        if col_letter == "AC" and (row == 75 or row == 68):
                            break

                        # 如果是第一个sheet，高炭不需要考虑 + 72行的逻辑
                        if sheet_idx == 0:
                            # 读取第72行的值
                            main_val72 = 0  # safe_float(get_merged_value(sheet, 72, main_col_idx))
                            dep_val72 = 0  # safe_float(get_merged_value(sheet, 72, dep_col_idx))

                            # 如果值大于0，则相加
                            if main_val > 0 and main_val72 > 0:
                                main_val += main_val72
                            if dep_val > 0 and dep_val72 > 0:
                                dep_val += dep_val72

                            if col_letter == "P" and row == 56:
                                continue
                            if col_letter == "V" and row == 62:
                                continue
                            if col_letter == "Y" and row == 62:
                                continue
                            if col_letter == "V" and row == 62:
                                continue
                            if col_letter == "AA" and row == 62:
                                continue

                            # 使用相加后的值
                            value = dep_val  # main_val +
                            # print("gaotan-idx0",main_val,dep_val)
                        else:

                            # 对于后续的sheet，使用71行的值,参考74行的值（结余值），如果结余值74大于 71值，则71值0，否则71值-74值
                            main_val74 = safe_float(get_merged_value(sheet, 74, main_col_idx))
                            dep_val74 = safe_float(get_merged_value(sheet, 74, dep_col_idx))

                            # 计算差值（只取正值）
                            main_val = max(0, main_val - main_val74)
                            dep_val = max(0, dep_val - dep_val74)

                            value = dep_val  # main_val +

                            if col_letter == "P" and row == 56:
                                continue
                            if col_letter == "V" and row == 62:
                                continue
                            if col_letter == "Y" and row == 62:
                                continue
                            if col_letter == "AA" and row == 62:
                                continue

                        product *= value
                        # print("calculate_product-gaotan", col_letter, col_idx,sheet_idx, row,dep_val,main_val,value,product)
                    else:
                        continue
                else:
                    # 普通列的处理
                    raw_value = get_merged_value(sheet, row, col_idx)
                    clean_value = safe_float(raw_value)  # 获得第71行col_idx列数据

                    if col_letter == "P" and (row == 75 or row == 68):
                        break
                    if col_letter == "R" and (row == 75 or row == 68):
                        break
                    if col_letter == "V" and (row == 75 or row == 68):
                        break
                    if col_letter == "Y" and (row == 75 or row == 68):
                        break
                    if col_letter == "AC" and (row == 75 or row == 68):
                        break
                    print("row-feigaotan1", row, col_letter, raw_value, clean_value)

                    # 如果是第一个sheet且是71行，检查是否有72行需要相加
                    if sheet_idx == 0:
                        # raw_value72 = get_merged_value(sheet, 72, col_idx)
                        # clean_value72 = safe_float(raw_value72)

                        if clean_value > 0:
                            clean_value += 0
                    else:
                        # 对于后续的sheet，使用71行的值减去74行的值
                        main_val74 = safe_float(get_merged_value(sheet, 74, col_idx))
                        if clean_value > 0 and main_val74 > 0 and clean_value < main_val74:
                            clean_value = 0
                        elif clean_value > 0 and main_val74 > 0 and clean_value >= main_val74:
                            clean_value = clean_value - main_val74
                        # 对于后续的行

                    if col_letter == "M" and (row == 40 or row == 56 or row == 62):
                        continue

                    if col_letter == "N" and (row == 40 or row == 56 or row == 62):
                        continue
                    if col_letter == "O" and (row == 40 or row == 56):
                        continue
                    if col_letter == "P" and (row == 40 or row == 56):
                        continue

                    product *= clean_value
                    print("row-feigaotan2", row, col_letter, raw_value, clean_value, product)
                    # if nindex ==0:
                    # print("calculate_product-feigaotan", col_letter, col_idx,  clean_value, product)
                nindex += 1

            if product > 0:
                total += product

        print("row-feigaotan3", total)
        return int(total)