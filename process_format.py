
# -*- coding: utf-8 -*-
import pandas as pd
from tkinter import Tk, filedialog

VERSION = 0.2

def get_version():
    return VERSION

def read_excel_to_string_list():
    # 创建一个Tkinter窗口，但不显示
    root = Tk()
    root.withdraw()
    
    # 弹出文件选择框，让用户选择Excel文件
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
    )
    
    # 如果用户取消选择，直接退出
    if not file_path:
        print("未选择文件，程序退出。")
        return False
    
    # 使用pandas读取Excel文件
    try:
        df = pd.read_excel(file_path, dtype=str)  # 确保所有数据都以字符串形式读取
        # 将DataFrame转换为二维字符串列表
        string_list = df.values.tolist() 
        return string_list
    
    except Exception as e:
        print(f"读取Excel文件时出错：{e}")
        return False

def del_img(data):
    """
    规范化二维字符串列表，删除第4列（索引为3）包含'图片'的行。
    
    参数:
        data (list of list of str): 输入的二维字符串列表。
    
    返回:
        list of list of str: 规范化后的二维字符串列表。
    """
    print('正在删除类别为图片的行...')
    # 使用列表推导式过滤掉符合条件的行
    normalized_data = [row for row in data if len(row) < 4 or '图片' not in row[3]]
    print('删除成功!')
    return normalized_data

def read_replace():
    """
    读取 Excel 文件并转换为二维字符串列表，列数为 3 列。
    如果第三列的值不是 'Y' 或 'N'，则打印错误并返回 False。
    
    参数:
        file_path (str): Excel 文件路径。
    
    返回:
        bool: 如果第三列的值全部为 'Y' 或 'N'，返回 True；否则返回 False。
    """
    try:
        # 读取 Excel 文件
        print('正在读取替换词表...')
        df = pd.read_excel('替换词表.xlsx')  # 不将第一行作为列名

        # 确保数据是二维字符串列表，列数为 3 列
        if df.shape[1] < 3:
            df = df.reindex(columns=range(3))

        # 将数据转换为二维字符串列表，空单元格替换为空字符串
        data = df.iloc[:, :3].fillna('').astype(str).values.tolist()

        # 遍历第三列的所有字符串
        for row in data:
            third_column_value = row[2]  # 获取第三列的值
            if third_column_value not in ['Y', 'N', '']:  # 检查是否为 'Y' 或 'N'
                print(f"错误! 第三列中的'{third_column_value}'是无效值!")
                return False  # 返回 False 表示出错

        # 如果没有错误，返回 True
        print('替换词表读取成功!')
        return data

    except Exception as e:
        print(f"Error: {e}")
        return False  # 如果发生异常，返回 False

def main_function():
    re_list = read_replace()
    if re_list == False:
        return
    re_list.append(['�', '', 'N'])

    txt_list = read_excel_to_string_list()
    if txt_list == False:
        return
    txt_list = del_img(txt_list)
    #print(txt_list)
    

if __name__ == "__main__":
    main_function()