# 仅适用筛选"自提"订单的数据统计

import pandas as pd
import numpy as np
import os
import re
from datetime import datetime
import time
import sys

def excel_data_analyzer():
    """
    Excel高级数据分析工具 - 增强版
    
    整合功能：
    1. 通用多条件筛选与统计
    2. 专业图书借还月度分析
    
    使用说明：
    - 加载文件后选择分析模式
    - 模式1：通用多条件筛选（适合任意数据）
    - 模式2：图书借还月度分析（专为图书订单设计）
    """
    
    # 打印欢迎信息
    print("\n" + "="*80)
    print("Excel高级数据分析工具".center(80))
    print(f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}".center(80))
    print("="*80 + "\n")
    
    # 文件路径处理
    file_path = input("请输入Excel文件路径: ").strip()
    
    # 验证文件存在性
    if not os.path.exists(file_path):
        print("\n❌ 错误: 文件不存在，请检查路径！")
        return
    
    # 读取Excel文件
    print("\n⏳ 正在读取Excel文件...")
    start_time = time.time()
    
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        original_rows, original_cols = df.shape
        load_time = time.time() - start_time
        
        # 显示文件基本信息
        print(f"\n✅ 文件加载成功！耗时: {load_time:.2f}秒")
        print(f"📊 数据维度: {original_rows} 行, {original_cols} 列")
        
        # 显示列信息
        print("\n🔍 数据列信息:")
        for i, col in enumerate(df.columns):
            # 获取列类型
            if pd.api.types.is_numeric_dtype(df[col]):
                col_type = "数值"
            elif pd.api.types.is_datetime64_any_dtype(df[col]):
                col_type = "日期"
            else:
                col_type = "文本"
            
            # 获取唯一值数量
            unique_count = df[col].nunique()
            
            # 获取样本值
            sample_values = df[col].dropna().unique()[:5]
            sample_str = ", ".join(map(str, sample_values))
            
            print(f"{i+1}. {col} ({col_type}, {unique_count}个唯一值)")
            print(f"  示例: [{sample_str}{'...' if len(sample_values) == 5 else ''}]")
        
        # ====================
        # 功能选择菜单
        # ====================
        print("\n" + "="*80)
        print("功能选择".center(80))
        print("="*80)
        print("1. 通用多条件筛选与统计")
        print("2. 图书借还月度分析")
        print("3. 退出")
        print("="*80)
        
        choice = input("请选择功能 (1-3): ").strip()
        
        if choice == '3':
            print("\n👋 感谢使用，再见！")
            return
        elif choice == '2':
            # 调用图书借还月度分析功能
            book_order_analysis(df)
        else:
            # 默认选择通用多条件筛选
            general_data_analysis(df)
        
    except Exception as e:
        print(f"\n❌ 发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

def general_data_analysis(df):
    """
    通用多条件筛选与统计功能
    
    功能：
    1. 多条件筛选（任意列组合）
    2. 多值筛选（支持逗号分隔）
    3. 数据汇总统计
    4. 导出指定列
    
    参数:
    df -- pandas DataFrame对象
    """
    original_rows, original_cols = df.shape
    
    # 筛选条件收集
    print("\n" + "="*80)
    print("多条件筛选设置 (输入 'done' 结束条件输入)".center(80))
    print("="*80)
    print("💡 提示: 多值条件用逗号分隔 (例: 5,6,7 或 借书,还书)")
    
    filters = {}
    while True:
        col_name = input("\n请输入要筛选的列名 (或输入 'done' 结束): ").strip()
        
        if col_name.lower() == 'done':
            break
            
        if col_name not in df.columns:
            print(f"❌ 错误: 列名 '{col_name}' 不存在，请重新输入！")
            continue
            
        # 显示列详细信息
        col_type = "数值" if pd.api.types.is_numeric_dtype(df[col_name]) else "文本"
        unique_count = df[col_name].nunique()
        sample_values = df[col_name].dropna().unique()[:5]
        sample_str = ", ".join(map(str, sample_values))
        print(f"  🔍 列类型: {col_type}, 唯一值数量: {unique_count}")
        print(f"  📋 示例值: [{sample_str}{'...' if len(sample_values) == 5 else ''}]")
        
        # 获取筛选值
        condition_input = input(f"请输入筛选值 (多值用逗号分隔): ").strip()
        
        # 处理多值输入
        values = [v.strip() for v in condition_input.split(',') if v.strip()]
        
        # 数值类型转换
        if pd.api.types.is_numeric_dtype(df[col_name]):
            try:
                values = [float(v) for v in values]
            except ValueError:
                print(f"❌ 错误: 列 '{col_name}' 是数值类型，请输入数字！")
                continue
        
        filters[col_name] = values
    
    # 应用筛选条件
    filtered_df = df.copy()
    if filters:
        print("\n🔧 正在应用筛选条件...")
        for col, values in filters.items():
            if pd.api.types.is_numeric_dtype(df[col]):
                # 数值列筛选
                mask = filtered_df[col].isin(values)
                print(f"  条件: {col} 在 {values} 中 → 匹配 {mask.sum()} 行")
            else:
                # 文本列筛选 (支持部分匹配)
                pattern = '|'.join([re.escape(str(v)) for v in values])
                mask = filtered_df[col].astype(str).str.contains(pattern, case=False, na=False)
                print(f"  条件: {col} 包含 {values} 中的值 → 匹配 {mask.sum()} 行")
            
            filtered_df = filtered_df[mask]
    
    # 显示筛选结果
    print("\n" + "="*80)
    print(f"筛选结果: {len(filtered_df)} 行数据 (原始数据: {original_rows} 行)".center(80))
    print("="*80)
    
    if len(filtered_df) == 0:
        print("\n⚠️ 没有匹配的数据！")
        return
    
    # 数据汇总统计
    print("\n📊 数据汇总统计")
    print("="*80)
    
    # 自动识别数值列进行汇总
    numeric_cols = filtered_df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        print("数值列汇总统计:")
        summary = filtered_df[numeric_cols].agg(['count', 'sum', 'mean', 'min', 'max'])
        print(summary)
    else:
        print("未检测到数值列")
    
    # 分类统计
    print("\n🔠 分类统计:")
    # 选择前5个分类列进行统计
    categorical_cols = [col for col in filtered_df.columns 
                       if not pd.api.types.is_numeric_dtype(filtered_df[col])
                       and filtered_df[col].nunique() < 50][:5]
    
    for col in categorical_cols:
        print(f"\n{col} 分类统计:")
        counts = filtered_df[col].value_counts()
        # 添加百分比
        percentages = counts / counts.sum() * 100
        stats = pd.DataFrame({
            '数量': counts,
            '占比(%)': percentages.round(1)
        })
        print(stats.head(10))  # 显示前10项
    
    # 导出选项
    print("\n" + "="*80)
    export = input("是否导出筛选结果? (y/n): ").strip().lower()
    if export == 'y':
        print("\n🔄 导出选项")
        print("="*80)
        print("可用列名: " + ", ".join(filtered_df.columns))
        
        # 获取用户选择的列
        selected_cols = input("\n请输入要导出的列名 (多个用逗号分隔，留空导出所有列): ").strip()
        
        if selected_cols:
            selected_cols = [col.strip() for col in selected_cols.split(',')]
            # 验证列名
            valid_cols = [col for col in selected_cols if col in filtered_df.columns]
            invalid_cols = [col for col in selected_cols if col not in filtered_df.columns]
            
            if invalid_cols:
                print(f"⚠️ 忽略无效列名: {', '.join(invalid_cols)}")
            
            if valid_cols:
                export_df = filtered_df[valid_cols]
                print(f"✅ 已选择 {len(valid_cols)} 列进行导出")
            else:
                print("⚠️ 没有有效列名，将导出所有列")
                export_df = filtered_df
        else:
            export_df = filtered_df
            print("✅ 将导出所有列")
        
        # 获取导出路径
        default_path = f"筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        export_path = input(f"请输入导出文件路径 (默认: {default_path}): ").strip()
        export_path = export_path if export_path else default_path
        
        try:
            # 确保文件扩展名正确
            if not export_path.lower().endswith('.xlsx'):
                export_path += '.xlsx'
            
            export_df.to_excel(export_path, index=False)
            print(f"\n✅ 筛选结果已导出至: {os.path.abspath(export_path)}")
            print(f"📊 导出数据: {export_df.shape[0]} 行, {export_df.shape[1]} 列")
        except Exception as e:
            print(f"\n❌ 导出失败: {str(e)}")
    
    print("\n" + "="*80)
    print("通用筛选分析完成!".center(80))
    print("="*80)

def book_order_analysis(df):
    """
    图书借还月度分析功能 - 增强版
    
    功能：
    1. 按月份分别统计借书和还书的书籍数量
    2. 统计借书和还书的人次（订单数量）
    3. 支持自定义选择月份范围
    4. 生成专业分析报告
    
    增强功能：
    - 智能月份识别（支持数值、文本、日期格式）
    - 详细的调试信息
    - 健壮的错误处理
    
    参数:
    df -- pandas DataFrame对象
    """
    # 自动检测关键列
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower()
        if '订单' in col_lower and '类型' in col_lower:
            col_mapping['order_type'] = col
        elif '数量' in col_lower or '册数' in col_lower:
            col_mapping['quantity'] = col
        elif '月份' in col_lower or 'month' in col_lower or '年月' in col_lower:
            col_mapping['month'] = col
        elif '订单' in col_lower and 'id' in col_lower:
            col_mapping['order_id'] = col
    
    # 验证关键列是否存在
    required_cols = ['order_type', 'quantity', 'month']
    missing_cols = [col for col in required_cols if col not in col_mapping]
    if missing_cols:
        print("\n❌ 错误: 未能自动识别以下关键列:")
        print(" - " + "\n - ".join(missing_cols))
        print("\n请手动指定列名:")
        for col in missing_cols:
            col_mapping[col] = input(f"请输入'{col.replace('_', ' ')}'对应的列名: ").strip()
    
    order_col = col_mapping['order_type']
    qty_col = col_mapping['quantity']
    month_col = col_mapping['month']
    
    # 验证订单类型
    order_types = df[order_col].unique()
    if '借书订单' not in order_types and '借书' not in order_types:
        print("\n⚠️ 注意: 订单类型列缺少'借书'值")
    if '还书订单' not in order_types and '还书' not in order_types:
        print("\n⚠️ 注意: 订单类型列缺少'还书'值")
    
    # 显示订单类型分布
    print("\n🔍 订单类型分布:")
    order_counts = df[order_col].value_counts()
    print(order_counts)
    
    # 验证数量列为数值型
    if not pd.api.types.is_numeric_dtype(df[qty_col]):
        print(f"\n⚠️ 注意: '{qty_col}'列为文本类型，尝试转换为数字...")
        df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce')
        if df[qty_col].isna().any():
            nan_count = df[qty_col].isna().sum()
            print(f"❌ 转换失败，有 {nan_count} 行无法转换为数字")
            print("将使用0填充无效值")
            df[qty_col] = df[qty_col].fillna(0)
    
    # ===========================================
    # 智能月份识别系统 - 支持多种格式
    # ===========================================
    print("\n🔍 正在处理月份数据...")
    
    # 创建月份副本用于处理
    month_series = df[month_col].copy()
    
    # 方法1：尝试直接转换为数值
    print("尝试方法1: 直接转换为数值...")
    try:
        month_series = pd.to_numeric(month_series, errors='coerce')
        if not month_series.isna().all():
            print("✅ 成功将月份列转换为数值")
            df['month_num'] = month_series
        else:
            print("❌ 方法1失败: 无法转换为数值")
    except Exception as e:
        print(f"❌ 方法1失败: {str(e)}")
    
    # 方法2：尝试从字符串中提取数字（如果方法1失败）
    if 'month_num' not in df.columns or df['month_num'].isna().all():
        print("\n尝试方法2: 从文本中提取月份数字...")
        try:
            # 增强提取逻辑：匹配多种可能的数字格式
            extracted = month_series.astype(str).str.extract(r'(\d{1,2})', expand=False)
            if extracted.notna().any():
                df['month_num'] = pd.to_numeric(extracted, errors='coerce')
                print("✅ 成功从文本中提取月份数字")
            else:
                print("❌ 方法2失败: 文本中未找到数字")
        except Exception as e:
            print(f"❌ 方法2失败: {str(e)}")
    
    # 方法3：尝试解析日期格式（如果前两种方法失败）
    if 'month_num' not in df.columns or df['month_num'].isna().all():
        print("\n尝试方法3: 解析日期格式...")
        try:
            # 尝试常见日期格式
            date_formats = ['%Y-%m', '%m/%d/%Y', '%Y/%m/%d', '%b %Y', '%B %Y', 
                           '%Y%m', '%m-%Y', '%Y年%m月', '%m月']
            parsed = None
            
            for fmt in date_formats:
                try:
                    temp_parsed = pd.to_datetime(month_series, format=fmt, errors='coerce')
                    if temp_parsed.notna().any():
                        parsed = temp_parsed
                        print(f"✅ 成功解析日期格式: {fmt}")
                        break
                except:
                    continue
            
            if parsed is not None:
                df['month_num'] = parsed.dt.month
                print("✅ 从日期中提取月份成功")
            else:
                print("❌ 方法3失败: 无法解析日期格式")
        except Exception as e:
            print(f"❌ 方法3失败: {str(e)}")
    
    # 最终回退：使用原始值作为月份
    if 'month_num' not in df.columns:
        print("\n⚠️ 使用原始月份值")
        df['month_num'] = month_series
    
    # 显示月份处理结果
    print("\n🔍 月份处理结果样本 (前5行):")
    print(df[[month_col, 'month_num']].head())
    
    # 获取实际月份范围
    valid_months = df[df['month_num'].notna()]['month_num']
    min_month = int(valid_months.min()) if not valid_months.empty else 0
    max_month = int(valid_months.max()) if not valid_months.empty else 0
    
    # 检查月份值是否在1-12范围内
    if min_month < 1 or max_month > 12:
        print(f"\n⚠️ 注意: 月份值范围 {min_month}-{max_month} 超出1-12的正常范围")
        print("这可能表示数据格式问题，但分析将继续进行")
    
    print(f"\n📅 数据包含月份范围: {min_month}月 至 {max_month}月")
    
    # 选择分析的月份
    print("\n" + "="*80)
    print("月份选择".center(80))
    print("="*80)
    print(f"可用月份: {min_month}-{max_month}")
    month_input = input("请输入要分析的月份(多个月份用逗号分隔, 输入'all'分析所有月份): ").strip()
    
    if month_input.lower() == 'all':
        months = list(range(min_month, max_month + 1))
        print(f"✅ 将分析所有月份: {min_month}月 至 {max_month}月")
    else:
        try:
            months = [int(m.strip()) for m in month_input.split(',') if m.strip()]
            # 验证月份是否在范围内
            invalid_months = [m for m in months if m < min_month or m > max_month]
            if invalid_months:
                print(f"⚠️ 忽略无效月份: {', '.join(map(str, invalid_months))}")
                months = [m for m in months if min_month <= m <= max_month]
            print(f"✅ 将分析月份: {', '.join(map(str, months))}")
        except:
            print("❌ 月份输入格式错误，将分析所有月份")
            months = list(range(min_month, max_month + 1))
            print(f"✅ 将分析所有月份: {min_month}月 至 {max_month}月")
    
    # 筛选数据 - 使用数值型月份列
    filtered_df = df[df['month_num'].isin(months)]
    print(f"\n🔍 筛选到 {len(filtered_df)} 条记录 (总记录: {len(df)})")
    
    if len(filtered_df) == 0:
        print("\n⚠️ 没有匹配的数据！")
        return
    
    # 按月统计
    print("\n" + "="*80)
    print("图书借还月度统计结果".center(80))
    print("="*80)
    
    # 创建统计表格
    results = []
    for month in months:
        month_data = filtered_df[filtered_df['month_num'] == month]
        
        # 借书统计（书籍数量）
        borrow_data = month_data[
            month_data[order_col].astype(str).str.contains('借书', case=False, na=False)
        ]
        borrow_qty = borrow_data[qty_col].sum()
        
        # 还书统计（书籍数量）
        return_data = month_data[
            month_data[order_col].astype(str).str.contains('还书', case=False, na=False)
        ]
        return_qty = return_data[qty_col].sum()
        
        # 借书人次统计（订单数量）
        borrow_orders = borrow_data.shape[0]
        
        # 还书人次统计（订单数量）
        return_orders = return_data.shape[0]
        
        # 净借书量
        net_qty = borrow_qty - return_qty
        
        # 将结果添加到统计表中
        results.append({
            '月份': month,
            '借书量(册)': borrow_qty,
            '还书量(册)': return_qty,
            '借书人次': borrow_orders,
            '还书人次': return_orders,
            '净借书量(册)': net_qty
        })
    
    # 创建结果DataFrame
    result_df = pd.DataFrame(results)
    
    # 添加总计行
    if not result_df.empty:
        total_row = {
            '月份': '总计',
            '借书量(册)': result_df['借书量(册)'].sum(),
            '还书量(册)': result_df['还书量(册)'].sum(),
            '借书人次': result_df['借书人次'].sum(),
            '还书人次': result_df['还书人次'].sum(),
            '净借书量(册)': result_df['净借书量(册)'].sum()
        }
        result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)
    
    # 打印结果
    print("\n📊 按月统计结果:")
    print(result_df.to_string(index=False))
    
    # 添加分析报告
    print("\n📈 趋势分析:")
    if not result_df.empty:
        # 1. 借书最多的月份
        max_borrow_idx = result_df['借书量(册)'].idxmax()
        max_borrow_month = result_df.loc[max_borrow_idx, '月份']
        max_borrow_value = result_df.loc[max_borrow_idx, '借书量(册)']
        
        # 2. 还书最多的月份
        max_return_idx = result_df['还书量(册)'].idxmax()
        max_return_month = result_df.loc[max_return_idx, '月份']
        max_return_value = result_df.loc[max_return_idx, '还书量(册)']
        
        # 3. 借书人次最多的月份
        max_borrow_orders_idx = result_df['借书人次'].idxmax()
        max_borrow_orders_month = result_df.loc[max_borrow_orders_idx, '月份']
        max_borrow_orders_value = result_df.loc[max_borrow_orders_idx, '借书人次']
        
        # 4. 还书人次最多的月份
        max_return_orders_idx = result_df['还书人次'].idxmax()
        max_return_orders_month = result_df.loc[max_return_orders_idx, '月份']
        max_return_orders_value = result_df.loc[max_return_orders_idx, '还书人次']
        
        # 5. 净借书最多的月份
        max_net_idx = result_df['净借书量(册)'].idxmax()
        max_net_month = result_df.loc[max_net_idx, '月份']
        max_net_value = result_df.loc[max_net_idx, '净借书量(册)']
        
        # 6. 净借书最少的月份
        min_net_idx = result_df['净借书量(册)'].idxmin()
        min_net_month = result_df.loc[min_net_idx, '月份']
        min_net_value = result_df.loc[min_net_idx, '净借书量(册)']
        
        print(f"• 借书量最高的月份: {max_borrow_month}月 ({max_borrow_value}册)")
        print(f"• 还书量最高的月份: {max_return_month}月 ({max_return_value}册)")
        print(f"• 借书人次最多的月份: {max_borrow_orders_month}月 ({max_borrow_orders_value}人次)")
        print(f"• 还书人次最多的月份: {max_return_orders_month}月 ({max_return_orders_value}人次)")
        print(f"• 净借书量最高的月份: {max_net_month}月 ({max_net_value}册)")
        print(f"• 净借书量最低的月份: {min_net_month}月 ({min_net_value}册)")
    else:
        print("⚠️ 没有统计数据")
    
    # 导出结果
    print("\n" + "="*80)
    export = input("是否导出统计结果? (y/n): ").strip().lower()
    if export == 'y':
        default_name = f"图书借还统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        export_path = input(f"请输入导出路径(默认: {default_name}): ").strip() or default_name
        
        if not export_path.endswith('.xlsx'):
            export_path += '.xlsx'
        
        try:
            # 创建Excel写入器
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                # 写入统计结果
                result_df.to_excel(writer, sheet_name='月度统计', index=False)
                
                # 写入原始数据（可选）
                include_raw = input("是否包含原始数据? (y/n): ").strip().lower()
                if include_raw == 'y':
                    # 创建副本，移除临时列
                    export_df = filtered_df.copy()
                    if 'month_num' in export_df.columns:
                        export_df = export_df.drop(columns=['month_num'])
                    export_df.to_excel(writer, sheet_name='原始数据', index=False)
                
                # 添加图表（可选）
                try:
                    if not result_df.empty and len(result_df) > 1:  # 需要多于一行数据（不包括总计）
                        from openpyxl.chart import BarChart, Reference
                        from openpyxl.chart.axis import DateAxis
                        workbook = writer.book
                        ws = workbook['月度统计']
                        
                        # 创建书籍数量统计图表
                        qty_chart = BarChart()
                        qty_chart.type = "col"
                        qty_chart.style = 10
                        qty_chart.title = "图书借还数量统计"
                        qty_chart.y_axis.title = '数量(册)'
                        qty_chart.x_axis.title = '月份'
                        
                        # 添加书籍数量数据
                        qty_data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=len(result_df)-1)  # 排除总计行
                        qty_categories = Reference(ws, min_col=1, min_row=2, max_row=len(result_df)-1)
                        
                        qty_chart.add_data(qty_data, titles_from_data=True)
                        qty_chart.set_categories(qty_categories)
                        
                        # 放置图表
                        ws.add_chart(qty_chart, "F2")
                        
                        # 创建人次统计图表
                        orders_chart = BarChart()
                        orders_chart.type = "col"
                        orders_chart.style = 11
                        orders_chart.title = "图书借还人次统计"
                        orders_chart.y_axis.title = '人次'
                        orders_chart.x_axis.title = '月份'
                        
                        # 添加人次数据
                        orders_data = Reference(ws, min_col=4, min_row=1, max_col=5, max_row=len(result_df)-1)
                        orders_categories = Reference(ws, min_col=1, min_row=2, max_row=len(result_df)-1)
                        
                        orders_chart.add_data(orders_data, titles_from_data=True)
                        orders_chart.set_categories(orders_categories)
                        
                        # 放置图表
                        ws.add_chart(orders_chart, "F18")
                except Exception as e:
                    print(f"⚠️ 图表创建失败: {str(e)}")
            
            print(f"\n✅ 统计结果已导出至: {os.path.abspath(export_path)}")
        except Exception as e:
            print(f"\n❌ 导出失败: {str(e)}")
    
    print("\n" + "="*80)
    print("图书借还分析完成!".center(80))
    print("="*80)

if __name__ == "__main__":
    # 主程序入口
    try:
        excel_data_analyzer()
    except KeyboardInterrupt:
        print("\n\n👋 用户中断程序，再见！")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 发生未处理的错误: {str(e)}")
        import traceback
        traceback.print_exc()

        