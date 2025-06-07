import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
#打包用
from multiprocessing import freeze_support

def main():
    """主程序函数"""
    try:
        # 获取当前脚本运行目录
        current_directory = os.getcwd()

        # 构造文件路径
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
        except Exception as e:
            print(f"错误: 无法读取表格2 ({file2_path})，错误信息: {e}")
            return

        print("表格2读取成功，开始处理数据...")

        # 读取Excel文件
        table1 = pd.read_excel(file1_path)
        table2 = pd.read_excel(file2_path)

        # 打印表格1的前几行，检查列名和数据
        print("表格1内容预览:")
        print(table1.head())

        # 初始化输出表格4数据列表
        output_data = []

        # 初始化匹配失败数据列表
        failed_data = []

        # 清洗函数：去除所有的空格、冒号、星号和其他符号
        def clean_string(s):
            return ''.join(e for e in s.lower() if e.isalnum())  # 转为小写并去除非字母数字的字符

        # 解析表格2的每一行数据
        for index, row in table2.iterrows():
            order_number = row['*销售单号']
            recipient = row['收货人']
            phone = row['收货联系方式']
            province = row['收货地址-省']
            city = row['收货地址-市']
            district = row['收货地址-区']
            address = row['收货地址-详细地址']
            items = row['商品&数量']
            
            # 判断商品&数量列是否有多个商品，通过英文逗号分割
            for item in items.split(','):
                try:
                    # 按从左往右最右边的 '*' 分割商品和数量
                    parts = item.rsplit('*', 1)  # 从右往左查找，只分割最后一个*
                    if len(parts) != 2:
                        print(f"警告: 商品格式不正确 '{item.strip()}'，跳过处理")
                        continue
                    product, quantity = parts
                    quantity = int(quantity)

                    # 调试输出，检查品名内容
                    print(f"处理订单 {order_number}，商品: {item.strip()}，数量: {quantity}")

                    # 清理表格2中的品名
                    item_cleaned = clean_string(item.strip())
                    print(f"清理后的品名: {item_cleaned}")

                    # 进行匹配，确保表格1中的品名也进行相同的清理
                    table1_A = table1.iloc[:, 0].str.strip().apply(clean_string)  # 假设品名在第1列（索引0）
                    print(f"表格2匹配表格1的品名: 表格2: {item_cleaned}，表格1: {table1_A}")
                    matched_row = table1[table1_A == item_cleaned]

                    # 输出调试信息，查看实际查询结果
                    print(f"查询品名: {item_cleaned}, 匹配结果: {matched_row.shape[0]} 行")

                    if not matched_row.empty:
                        matched_row = matched_row.iloc[0]
                        product_code = matched_row['套组货号']
                        product_name = matched_row['山姆系统套组品名']
                        unit_price = matched_row['单价']

                        # 计算金额
                        amount = quantity * unit_price

                        # 构造输出表格4的一行
                        output_data.append({
                            '导入编号': order_number,
                            '网店订单号': order_number,
                            '下单时间': '',  # 表格2没有下单时间，可以保持为空或根据需要填充
                            '付款时间': '',  # 表格2没有付款时间
                            '承诺发货时间': '',  # 表格2没有承诺发货时间
                            '客户账号': '',  # 无相关字段
                            '客户邮箱': '',  # 无相关字段
                            'QQ': '',  # 无相关字段
                            '收货人': recipient,
                            '手机': phone,
                            '固定电话': '',  # 无相关字段
                            '国家': '',  # 无相关字段
                            '省份': province,
                            '市（区）': city,
                            '区（县）': district,
                            '收货地址': address,
                            '邮政编码': '',  # 无相关字段
                            '发货仓库': '',  # 无相关字段
                            '应收邮资': '',  # 无相关字段
                            '平台佣金': '',  # 无相关字段
                            '客付税额': '',  # 无相关字段
                            '应收合计': amount,
                            '平台补贴': '',  # 无相关字段
                            '客服备注': '',  # 无相关字段
                            '客户备注': '',  # 无相关字段
                            '销售渠道名称': '山姆',
                            '终端平台类型': '',  # 无相关字段
                            '终端网店单号': '',  # 无相关字段
                            '终端实付金额': '',  # 无相关字段
                            '结算方式': '',  # 无相关字段
                            '结算币种': '',  # 无相关字段
                            '货品名称': product_name,
                            '商品链接id': '',  # 无相关字段
                            '条码': '',  # 无相关字段
                            '货品编号': product_code,
                            '规格': '默认规格',  # 默认规格
                            '是否赠品': '',  # 无相关字段
                            '批次号': '',  # 无相关字段
                            '数量': quantity,
                            '单价': unit_price,
                            '货品优惠': '',  # 无相关字段
                            '金额': '',  # 无相关字段
                            '网店子订单号': '',  # 无相关字段
                            '定制码': '',  # 无相关字段
                            '货品备注': '',  # 无相关字段
                            '发票抬头': '',  # 无相关字段
                            '发票类型': '',  # 无相关字段
                            '证件类型': '',  # 无相关字段
                            '证件号码': '',  # 无相关字段
                            '证件使用姓名': '',  # 无相关字段
                            '物流公司': '',  # 无相关字段
                            '物流单号': row['*物流单号'],
                            '支付单号': '',  # 无相关字段
                            '收款账户': '',  # 无相关字段
                            '业务员': '',  # 无相关字段
                            '跟单员': '',  # 无相关字段
                            '标记': ''  # 无相关字段
                        })
                    else:
                        print(f"警告: 在表格1中没有找到品名 '{item.strip()}' 对应的商品信息")
                        
                        # 将匹配失败的数据保存到失败列表
                        failed_data.append({
                            '销售单号': order_number,
                            '商品': item.strip(),
                            '收货人': recipient,
                            '数量': quantity,
                            '错误信息': f"未找到品名 '{item.strip()}' 对应的商品信息"
                        })

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


        # 转换为DataFrame
        df = pd.DataFrame(output_data)

        # 假设 df 是你最终的 DataFrame
        df['导入编号'] = df['导入编号'].astype(str)
        df['网店订单号'] = df['网店订单号'].astype(str)

        # 获取当前时间并格式化为 YYYYMMDD_HHMMSS
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        # 生成文件名
        file_name = f'附件四-上传OMS系统_{timestamp}.xlsx'

        # 保存 DataFrame 为 Excel
        df.to_excel(file_name, index=False)

        # 打开刚才保存的 Excel 文件
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
        # 提示用户按任意键关闭
        print("文件处理完毕，请按任意键关闭窗口...")
        input()  # 等待用户按任意键
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        print('请按任意键关闭窗口...')
        input()

if __name__ == "__main__":
    freeze_support()  # 支持multiprocessing打包为exe
    main()