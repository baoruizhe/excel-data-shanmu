import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
#打包用
from multiprocessing import freeze_support

def clean_string(s):
    """清洗函数：去除所有的空格、冒号、星号和其他符号"""
    return ''.join(e for e in s.lower() if e.isalnum())  # 转为小写并去除非字母数字的字符

def main():
    """主程序函数"""
    try:
        # 获取当前脚本运行目录
        current_directory = os.getcwd()

        # 构造文件路径 - 保持原有的文件名
        file1_path = os.path.join(current_directory, "附件一产品信息及编码-固定不变.xlsx")
        file2_path = os.path.join(current_directory, "附件二销售出库导出山姆原始表.xlsx")

        # 检查文件1是否存在
        if not os.path.exists(file1_path):
            print(f"错误: 未找到文件 {file1_path}")
            return

        # 读取表格1
        try:
            table1 = pd.read_excel(file1_path)
            if table1.empty:
                print(f"错误: 表格1 ({file1_path}) 内容为空，请检查文件数据。")
                return
            print(f"成功读取商品档案表，共 {len(table1)} 行数据")
        except Exception as e:
            print(f"错误: 无法读取表格1 ({file1_path})，错误信息: {e}")
            return

        print("表格1读取成功，开始处理数据...")

        # 检查文件2是否存在
        if not os.path.exists(file2_path):
            print(f"错误: 未找到文件 {file2_path}")
            return

        # 读取表格2
        try:
            table2 = pd.read_excel(file2_path)
            if table2.empty:
                print(f"错误: 表格2 ({file2_path}) 内容为空，请检查文件数据。")
                return
            print(f"成功读取销售清单表，共 {len(table2)} 行数据")
        except Exception as e:
            print(f"错误: 无法读取表格2 ({file2_path})，错误信息: {e}")
            return

        print("表格2读取成功，开始处理数据...")

        # 打印表格1的前几行，检查列名和数据
        print("表格1内容预览:")
        print(table1.head())

        # 创建商品映射字典，用于快速查找 - 应用excel1.0.4.py的逻辑改进
        product_mapping = {}
        for _, row in table1.iterrows():
            try:
                # 使用原有的列名格式
                product_name = str(row.iloc[0]).strip()  # 假设品名在第1列
                if product_name:
                    # 使用清理后的品名作为键
                    clean_name = clean_string(product_name)
                    product_mapping[clean_name] = {
                        'product_name': str(row.get('山姆系统套组品名', product_name)),
                        'product_code': str(row.get('套组货号', '')),
                        'unit_price': row.get('单价', 0),
                        'original_name': product_name
                    }
            except Exception as e:
                print(f"处理商品档案行时出错: {e}")
                continue

        print(f"建立商品映射完成，共 {len(product_mapping)} 个商品")

        # 初始化输出表格4数据列表
        output_data = []

        # 初始化匹配失败数据列表
        failed_data = []

        print("开始处理销售订单数据...")

        # 解析表格2的每一行数据 - 保持原有的列名格式
        for index, row in table2.iterrows():
            # 使用原有的列名
            order_number = row['*销售单号']
            recipient = row['收货人']
            phone = row['收货联系方式']
            province = row['收货地址-省']
            city = row['收货地址-市']
            district = row['收货地址-区']
            address = row['收货地址-详细地址']
            items = row['商品&数量']
            
            # 判断商品&数量列是否有多个商品，通过英文逗号分割 - 保持原有逻辑
            for item in items.split(','):
                try:
                    # 按 '*' 分割商品和数量
                    product, quantity = item.split('*')
                    quantity = int(quantity)

                    # 调试输出，检查品名内容
                    print(f"处理订单 {order_number}，商品: {item.strip()}，数量: {quantity}")

                    # 使用字典映射查找商品 - 应用excel1.0.4.py的改进逻辑
                    item_cleaned = clean_string(product.strip())
                    print(f"清理后的品名: {item_cleaned}")

                    if item_cleaned in product_mapping:
                        # 找到匹配的商品
                        product_info = product_mapping[item_cleaned]
                        product_name = product_info['product_name']
                        product_code = product_info['product_code']
                        unit_price = product_info['unit_price']

                        # 计算金额
                        amount = quantity * unit_price

                        print(f"查询品名: {item_cleaned}, 匹配成功")

                        # 构造输出表格4的一行 - 使用完整的字段结构
                        output_data.append({
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
                            '物流单号': row['*物流单号'],
                            '支付单号': '',
                            '收款账户': '',
                            '业务员': '',
                            '跟单员': '',
                            '标记': ''
                        })
                    else:
                        print(f"警告: 在表格1中没有找到品名 '{product.strip()}' 对应的商品信息")
                        
                        # 将匹配失败的数据保存到失败列表
                        failed_data.append({
                            '销售单号': order_number,
                            '商品': product.strip(),
                            '收货人': recipient,
                            '数量': quantity,
                            '错误信息': f"未找到品名 '{product.strip()}' 对应的商品信息"
                        })

                except Exception as e:
                    print(f"错误: 处理商品 '{item}' 时出错, 错误信息: {e}")
                    continue

        # 计算订单总额 - 应用excel1.0.4.py的改进逻辑
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

        # 转换为DataFrame
        df = pd.DataFrame(output_data)

        # 确保某些列为字符串类型
        df['导入编号'] = df['导入编号'].astype(str)
        df['网店订单号'] = df['网店订单号'].astype(str)

        # 获取当前时间并格式化为 YYYYMMDD_HHMMSS
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        # 生成文件名
        file_name = f'附件四-上传OMS系统_{timestamp}.xlsx'

        # 保存 DataFrame 为 Excel
        df.to_excel(file_name, index=False)

        # 调整列宽
        wb = load_workbook(file_name)
        ws = wb.active

        # 自动调整列宽
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # 获取列字母
            for cell in col:
                try:
                    # 处理 None 或 NaN 的情况
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            adjusted_width = (max_length + 2)  # 设置列宽，+2是为了留出一些空隙
            ws.column_dimensions[column].width = adjusted_width

        # 保存修改后的 Excel 文件
        wb.save(file_name)

        # 保存匹配失败的数据为一个新的文件
        if failed_data:
            failed_df = pd.DataFrame(failed_data)
            failed_df['销售单号'] = failed_df['销售单号'].astype(str)
            failed_filename = f"匹配表格1失败数据_{timestamp}.xlsx"
            failed_df.to_excel(failed_filename, index=False)
            print(f"匹配失败数据已保存为: {failed_filename}")

        print(f"文件已保存为: {file_name}")
        print("文件处理完毕，请按任意键关闭窗口...")
        os.system('pause')
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        print('请按任意键关闭窗口...')
        os.system('pause')

if __name__ == "__main__":
    main()