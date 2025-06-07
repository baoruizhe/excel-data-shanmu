#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据处理程序（基于字节码反编译重构）
功能：处理订单数据，匹配商品信息，生成OMS系统上传文件
"""

import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

def clean_string(s):
    """清理字符串，只保留字母和数字"""
    return ''.join(e for e in s.lower() if e.isalnum())

def main():
    """主程序函数"""
    try:
        # 文件路径设置
        file1_path = "商品档案表.xlsx"  # 商品信息表
        file2_path = "销售清单导出表.xlsx"  # 销售订单表
        
        print("正在读取商品档案表...")
        
        # 读取商品档案表（表格1）
        try:
            df1 = pd.read_excel(file1_path)
            print(f"成功读取商品档案表，共 {len(df1)} 行数据")
        except Exception as e:
            print(f"错误: 无法读取文件 '{file1_path}': {e}")
            return
        
        print("正在读取销售清单导出表...")
        
        # 读取销售清单表（表格2）
        try:
            df2 = pd.read_excel(file2_path)
            print(f"成功读取销售清单表，共 {len(df2)} 行数据")
        except Exception as e:
            print(f"错误: 无法读取文件 '{file2_path}': {e}")
            return
        
        # 创建商品映射字典，用于快速查找
        product_mapping = {}
        for _, row in df1.iterrows():
            try:
                product_name = str(row['品名']).strip()
                if product_name:
                    # 使用清理后的品名作为键
                    clean_name = clean_string(product_name)
                    product_mapping[clean_name] = {
                        'product_name': product_name,
                        'product_code': str(row.get('商品编码', '')),
                        'unit_price': row.get('单价', 0),
                        'category': str(row.get('类别', '')),
                        'brand': str(row.get('品牌', ''))
                    }
            except Exception as e:
                print(f"处理商品档案行时出错: {e}")
                continue
        
        print(f"建立商品映射完成，共 {len(product_mapping)} 个商品")
        
        # 处理销售订单数据
        output_data = []
        failed_data = []
        
        print("开始处理销售订单数据...")
        
        for index, row in df2.iterrows():
            try:
                # 提取订单基本信息
                order_number = str(row.get('销售单号', '')).strip()
                item = str(row.get('商品', '')).strip()
                recipient = str(row.get('收货人', '')).strip()
                phone = str(row.get('手机', '')).strip()
                province = str(row.get('省份', '')).strip()
                city = str(row.get('市（区）', '')).strip()
                district = str(row.get('区（县）', '')).strip()
                address = str(row.get('收货地址', '')).strip()
                quantity = row.get('数量', 1)
                amount = row.get('应收合计', 0)
                
                # 在商品映射中查找匹配的商品
                clean_item = clean_string(item)
                if clean_item in product_mapping:
                    # 找到匹配的商品
                    product_info = product_mapping[clean_item]
                    product_name = product_info['product_name']
                    product_code = product_info['product_code']
                    unit_price = product_info['unit_price']
                    
                    # 构建输出数据行
                    output_row = {
                        '导入编号': '',
                        '网店订单号': order_number,
                        '网店子订单号': '',
                        '网店会员名': '',
                        '货主': '',
                        '经销商': '',
                        '分公司/站点': '',
                        '下单时间': '',
                        '审单时间': '',
                        '备货时间': '',
                        '预计发货时间': '',
                        '实际发货时间': '',
                        '收货人': recipient,
                        '性别': '',
                        '手机': phone,
                        '固定电话': '',
                        '国家': '',
                        '省份': province,
                        '市（区）': city,
                        '区（县）': district,
                        '收货地址': address,
                        '邮政编码': '',
                        '发货仓库': '',
                        '应收邮资': '',
                        '平台佣金': '',
                        '客付税额': '',
                        '应收合计': amount,
                        '平台补贴': '',
                        '客服备注': '',
                        '客户备注': '',
                        '销售渠道名称': '山姆',
                        '终端平台类型': '',
                        '终端网店单号': '',
                        '终端实付金额': '',
                        '结算方式': '',
                        '结算币种': '',
                        '货品名称': product_name,
                        '商品链接id': '',
                        '条码': '',
                        '货品编号': product_code,
                        '规格': '默认规格',
                        '是否赠品': '',
                        '批次号': '',
                        '数量': quantity,
                        '单价': unit_price,
                        '货品优惠': '',
                        '金额': '',
                        '网店子订单号': '',
                        '定制码': '',
                        '货品备注': '',
                        '发票抬头': '',
                        '发票类型': '',
                        '证件类型': '',
                        '证件号码': '',
                        '证件使用姓名': '',
                        '物流公司': '',
                        '物流单号': row.get('*物流单号', ''),
                        '支付单号': '',
                        '收款账户': '',
                        '业务员': '',
                        '跟单员': '',
                        '标记': ''
                    }
                    
                    output_data.append(output_row)
                    
                else:
                    # 没有找到匹配的商品
                    print(f"警告: 在表格1中没有找到品名 '{item.strip()}' 对应的商品信息")
                    
                    failed_row = {
                        '销售单号': order_number,
                        '商品': item.strip(),
                        '收货人': recipient,
                        '数量': quantity,
                        '错误信息': f"未找到品名 '{item.strip()}' 对应的商品信息"
                    }
                    
                    failed_data.append(failed_row)
                    
            except Exception as e:
                print(f"错误: 处理商品 '{item}' 时出错, 错误信息: {e}")
                continue
        
        # 计算订单总额
        order_totals = {}
        for data in output_data:
            order_number = data['网店订单号']
            if order_number in order_totals:
                order_totals[order_number] += data['应收合计']
            else:
                order_totals[order_number] = data['应收合计']
        
        # 更新应收合计
        for data in output_data:
            data['应收合计'] = order_totals[data['网店订单号']]
        
        # 创建输出DataFrame
        df = pd.DataFrame(output_data)
        
        # 确保某些列为字符串类型
        df['导入编号'] = df['导入编号'].astype(str)
        df['网店订单号'] = df['网店订单号'].astype(str)
        
        # 生成时间戳
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # 生成输出文件名
        file_name = f'附件四-上传OMS系统_{timestamp}.xlsx'
        
        # 保存到Excel文件
        df.to_excel(file_name, index=False)
        
        # 调整列宽
        wb = load_workbook(file_name)
        ws = wb.active
        
        # 自动调整列宽
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column].width = adjusted_width
        
        wb.save(file_name)
        
        # 保存匹配失败的数据
        if failed_data:
            failed_df = pd.DataFrame(failed_data)
            failed_df['销售单号'] = failed_df['销售单号'].astype(str)
            failed_filename = f'匹配表格1失败数据_{timestamp}.xlsx'
            failed_df.to_excel(failed_filename, index=False)
            print(f'匹配失败数据已保存为: {failed_filename}')
        
        print(f'文件已保存为: {file_name}')
        print('文件处理完毕，请按任意键关闭窗口...')
        os.system('pause')
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        print('请按任意键关闭窗口...')
        os.system('pause')

if __name__ == "__main__":
    main() 