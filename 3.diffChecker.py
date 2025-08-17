import pandas as pd
import os

def compare_express_numbers(file_a, file_b, sheet_name, col_name):
    """
    比较两个Excel文件中快递单号列的差异
    """
    try:
        # 读取Excel文件，指定列为字符串类型
        df_a = pd.read_excel(file_a, sheet_name=sheet_name, dtype={col_name: str})
        df_b = pd.read_excel(file_b, sheet_name=sheet_name, dtype={col_name: str})
        
        # 提取单号列并去重
        nums_a = set(df_a[col_name].dropna().astype(str).str.strip())
        nums_b = set(df_b[col_name].dropna().astype(str).str.strip())
        
        # 计算差异
        only_in_a = nums_a - nums_b
        only_in_b = nums_b - nums_a
        common = nums_a & nums_b
        
        # 打印结果
        print("\n" + "="*50)
        print(f"文件A单号总数: {len(nums_a)}")
        print(f"文件B单号总数: {len(nums_b)}")
        print(f"共同单号数量: {len(common)}")
        print(f"A中独有单号数量: {len(only_in_a)}")
        print(f"B中独有单号数量: {len(only_in_b)}")
        print("="*50)
        
        # 将单号列表排序
        sorted_common = sorted(common)
        sorted_only_a = sorted(only_in_a)
        sorted_only_b = sorted(only_in_b)
        
        # 创建合并的排序列表
        all_sorted = sorted_common + sorted_only_a + sorted_only_b
        
        # 创建并排显示的数据
        display_data = []
        for num in all_sorted:
            a_display = num if num in nums_a else ""
            b_display = num if num in nums_b else ""
            display_data.append((a_display, b_display))
        
        # 打印并排比对结果
        print("\n" + "="*50)
        print("快递单号比对结果 (升序排列，显示首尾各10行):")
        print("="*50)
        print(f"{'A文件单号':<25} | {'B文件单号':<25}")
        print("-"*50)
        
        # 打印前10行
        for i in range(min(10, len(display_data))):
            a, b = display_data[i]
            print(f"{a:<25} | {b:<25}")
        
        # 打印省略信息（如果数据量大于20）
        if len(display_data) > 20:
            print(f"... 省略中间 {len(display_data)-20} 行 ...")
            
            # 打印后10行
            for i in range(len(display_data)-10, len(display_data)):
                a, b = display_data[i]
                print(f"{a:<25} | {b:<25}")
        elif len(display_data) > 10:
            # 如果数据量在11-20之间，直接打印剩余行
            for i in range(10, len(display_data)):
                a, b = display_data[i]
                print(f"{a:<25} | {b:<25}")
        
        print("="*50)
        
        if not only_in_a and not only_in_b:
            print("\n✅ 两个文件快递单号完全一致")
            return
        
        # 打印差异详情
        if only_in_a:
            print(f"\n🔍 A文件独有单号 ({len(only_in_a)}个):")
            print_sorted_numbers(only_in_a)
        
        if only_in_b:
            print(f"\n🔍 B文件独有单号 ({len(only_in_b)}个):")
            print_sorted_numbers(only_in_b)
    
    except Exception as e:
        print(f"\n❌ 发生错误: {str(e)}")
        print("请检查输入的文件路径、工作表名和列名是否正确")

def print_sorted_numbers(numbers, max_display=10):
    """打印排序后的单号列表，控制显示数量"""
    sorted_nums = sorted(numbers)
    
    # 打印前max_display个
    for num in sorted_nums[:max_display]:
        print(num)
    
    # 如果还有更多，显示统计信息
    if len(sorted_nums) > max_display:
        print(f"... 以及另外 {len(sorted_nums) - max_display} 个单号")

def get_user_input(prompt, default=None):
    """获取用户输入，支持默认值"""
    if default:
        user_input = input(f"{prompt} (默认: '{default}'): ").strip()
        return user_input if user_input else default
    return input(prompt).strip()

def print_welcome():
    """打印欢迎信息"""
    print("\n" + "="*50)
    print(" Excel快递单号比对工具 ".center(50, '★'))
    print("="*50)
    print("本工具用于比对两个Excel文件中的快递单号列")
    print("支持特性:")
    print("  - 自动处理长数字格式问题")
    print("  - 并排显示排序后的单号（首尾各10行）")
    print("  - 显示差异数量和具体单号")
    print("  - 支持自定义工作表和列名")
    print("="*50 + "\n")

if __name__ == "__main__":
    # 打印欢迎信息
    print_welcome()
    
    # 获取用户输入
    print("请指定第一个文件 (A文件):")
    file_a = get_user_input("  ➤ 文件路径: ")
    
    print("\n请指定第二个文件 (B文件):")
    file_b = get_user_input("  ➤ 文件路径: ")
    
    # 检查文件是否存在
    while not os.path.exists(file_a):
        print(f"\n❌ 文件不存在: {file_a}")
        file_a = get_user_input("  ➤ 请重新输入A文件路径: ")
    
    while not os.path.exists(file_b):
        print(f"\n❌ 文件不存在: {file_b}")
        file_b = get_user_input("  ➤ 请重新输入B文件路径: ")
    
    # 获取工作表和列名配置
    print("\n配置比对参数 (按Enter使用默认值)")
    sheet_name = get_user_input("  ➤ 工作表名称", "Sheet1")
    col_name = get_user_input("  ➤ 快递单号列名", "快递单号")
    
    print("\n" + "="*50)
    print(f"开始比对: {os.path.basename(file_a)} 和 {os.path.basename(file_b)}")
    print(f"工作表: '{sheet_name}', 列名: '{col_name}'")
    print("="*50)
    
    # 执行比对
    compare_express_numbers(file_a, file_b, sheet_name, col_name)
    
    # 结束信息
    print("\n" + "="*50)
    print("比对完成! 按Enter退出...")
    input()

    