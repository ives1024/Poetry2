import os
import win32com.client as win32  # 需要安装pywin32库

def convert_xls_to_xlsx(input_path):
    """
    将.xls文件转换为.xlsx格式（保留所有内容）
    :param input_path: 输入的.xls文件路径
    """
    # 检查输入文件是否存在
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"输入文件不存在: {input_path}")
    
    # 生成输出路径（替换扩展名）
    output_path = os.path.splitext(input_path)[0] + ".xlsx"
    
    # 通过Excel应用程序操作
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        # 设置不可见模式（静默操作）
        excel.Visible = False
        
        # 打开原始文件
        workbook = excel.Workbooks.Open(os.path.abspath(input_path))
        
        # 另存为新格式（FileFormat=51对应.xlsx）
        workbook.SaveAs(
            os.path.abspath(output_path),
            FileFormat=51,  # xlOpenXMLWorkbook
            ConflictResolution=2  # 覆盖现有文件
        )
        workbook.Close()
        
        print(f"转换成功: {input_path} -> {output_path}")
        
    except Exception as e:
        print(f"转换失败: {str(e)}")
    finally:
        # 确保退出Excel进程
        excel.Application.Quit()

if __name__ == "__main__":
    # 输入文件路径（示例：当前目录下的test.xls）
    input_file = "3m3513.xls"  # 修改为实际文件路径
    
    # 执行转换
    convert_xls_to_xlsx(input_file)