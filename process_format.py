
# -*- coding: utf-8 -*-
import pandas as pd
from tkinter import Tk, filedialog
import csv
import re

VERSION = 0.4
COLUMN_TYPE = 3 #类型所在的列数，从0开始数
COLUMN_SEND_DATA = 6 #发送内容所在的列数，从0开始数
COLUMN_REF_DATA = 7 #引用内容所在的列数，从0开始数
COLUMN_NAME = 4 #昵称内容所在的列数，从0开始数
COLUMN_TIME = 2 #时刻所在的列数，从0开始数
FONT_COLOR = '#D8DAD9' #灰色的HEX表示值
IS_TEST = True #是否是测试环境，使用时改为False

def get_version():
    return VERSION

def save_list_to_csv(data, filename):
    """
    将二维字符串列表保存为CSV文件。

    参数:
        data (list of list of str): 二维字符串列表，例如 [['a', 'b'], ['c', 'd']]
        filename (str): 要保存的CSV文件名，例如 'output.csv'
    """
    # 打开文件，准备写入
    if IS_TEST:
        with open(filename, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
        
            # 写入数据
            for row in data:
                writer.writerow(row)

        print(f"数据已成功保存到文件 {filename}")

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
    normalized_data = [row for row in data if len(row) < 4 or '图片' not in row[COLUMN_TYPE]]
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

def replace_list(data,re_list):
    # 遍历二维列表，处理第7列的字符串
    print('正在替换词...')
    result = []  # 用于存储处理后的二维列表
    for row in data:
        if isinstance(row[COLUMN_REF_DATA], str):  # 确保存在引用内容
            processed_string = replace_word(row[COLUMN_REF_DATA],re_list)  # 处理第8列的字符串（索引为7）
            row[COLUMN_REF_DATA] = processed_string  # 更新第8列的值

        if len(row) >= (COLUMN_SEND_DATA + 1):  # 确保当前行有至少7列
            processed_string = replace_word(row[COLUMN_SEND_DATA],re_list)  # 处理第7列的字符串（索引为6）
            if processed_string:  # 如果处理后的字符串不为空
                row[COLUMN_SEND_DATA] = processed_string  # 更新第7列的值
                result.append(row)  # 将处理后的行添加到结果列表
        else:
            result.append(row)  # 如果当前行不足7列，直接添加到结果列表
    print('替换完成!')
    return result

def replace_word(input_string, re_list):
    """
    替换字符串中指定的单词。

    :param input_string: 输入的原始字符串
    :param re_list: 二维字符串列表，每行包含3个元素：
    :return: 替换后的字符串
    """
    # 遍历替换列表
    for item in re_list:
        original_word, replacement_word, flag = item  # 分别提取每行的三个元素
        if flag == 'N':
            input_string = input_string.replace(original_word, replacement_word)
        else:
            pattern = re.escape(original_word) + r'{2,}'
        # 替换匹配到的内容为空字符串
            input_string = re.sub(pattern, replacement_word, input_string)
    
    return input_string

def normalize_2d_list(input_list):
    """
    遍历二维列表，删除字符串中的�字符。
    如果删除后为空，保留空字符串''。
    非字符串类型的元素保持不变。
    """
    print('正在删除无效字符...')
    normalized_list = []
    for row in input_list:
        new_row = []
        for item in row:
            # 检查是否为字符串类型
            if isinstance(item, str):
                # 删除字符串中的�字符
                cleaned_item = item.replace('�', '')
            else:
                # 如果不是字符串，保持原样
                cleaned_item = item
            new_row.append(cleaned_item)
        normalized_list.append(new_row)
    print('删除成功!')
    return normalized_list

def link_str(data):
    print('正在生成导入的Markdown文件...')
    string = ''
    for row in data:
        string = string + row[COLUMN_NAME] + '（' + row[COLUMN_TIME][:5] + '）：' + row[COLUMN_SEND_DATA]
        if isinstance(row[COLUMN_REF_DATA], str):
            string = string + '\n<font style="color:' + FONT_COLOR + ';">引用内容：' + row[COLUMN_REF_DATA] + '</font>\n'
        else:
            string = string + '\n'
    string = string.replace('\n','\n\n')
    return string

def main_function():
    re_list = read_replace()
    if re_list == False:
        return

    txt_list = read_excel_to_string_list()
    if txt_list == False:
        return
    save_list_to_csv(txt_list,'raw.csv')
    txt_list = del_img(txt_list)
    txt_list = normalize_2d_list(txt_list)
    txt_list = replace_list(txt_list,re_list)

    save_list_to_csv(txt_list,'fin.csv')
    content = link_str(txt_list)

    with open('import.md', "w", encoding="utf-8") as file:
        file.write(content)
    print('import.md生成成功!\n全部流程结束!')

    
if __name__ == "__main__":
    main_function()