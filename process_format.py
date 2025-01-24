
# -*- coding: utf-8 -*-
"""
Copyright (c) 2025, bin96
All rights reserved.

This script is licensed under the MIT License.
See LICENSE file for details.

Description:
The function of this script is to perform data processing
"""

import pandas as pd
from tkinter import Tk, filedialog
import csv
import re
from openai import OpenAI
import os

VERSION = 1.2
IS_TEST = False #是否是测试环境，使用时改为False
FONT_COLOR = '#D8DAD9' #灰色的HEX表示值

COLUMN_TYPE = 3         #类型所在的列数，从0开始数
COLUMN_SEND_DATA = 6    #发送内容所在的列数，从0开始数
COLUMN_REF_DATA = 7     #引用内容所在的列数，从0开始数
COLUMN_NAME = 4         #昵称内容所在的列数，从0开始数
COLUMN_TIME = 2         #时刻所在的列数，从0开始数
COLUMN_MULTI = 2        #连续才删除所在的列，从0开始数
COLUMN_HOST = 3         #是否未主持人所在的列，从0开始数

LINE_KEY = 0            #API KEY所在的行索引，行索引=行数-2
LINE_EN_AI = 1          #启用AI所在的行索引，行索引=行数-2
LINE_SUMM = 2           #生成内容概括所在的行索引，行索引=行数-2
LINE_CHAP = 3           #生成章节总结并嵌入正文所在的行索引，行索引=行数-2
LINE_SA = 4             #利用语义分析删除相关内容所在的行索引，行索引=行数-2
LINE_COMPARA = 5        #生成AI处理前后对比文档所在的行索引，行索引=行数-2
LINE_COST = 6           #输出API计价信息所在的行索引，行索引=行数-2
AI_MODEL = "moonshot-v1-8k" #KIMI API的模型


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
    #print('正在删除类别为图片的行...')
    # 使用列表推导式过滤掉符合条件的行
    normalized_data = [row for row in data if len(row) < 4 or '图片' not in row[COLUMN_TYPE]]
    print('类别为图片的行删除成功!')
    return normalized_data

def read_replace():
    try:
        # 读取Excel文件为DataFrame
        df = pd.read_excel('替换词表.xlsx', dtype=str)  # 将所有数据读取为字符串类型
        df.fillna('', inplace=True)  # 将所有空值替换为空字符串

        # 将DataFrame转换为二维字符串列表
        data_list = df.values.tolist()

        for row in data_list:
            multi_value = row[COLUMN_MULTI]
            host_value = row[COLUMN_HOST]

            if multi_value not in ['Y', '']:
                print(f"错误：'连续才删除'列的值 '{multi_value}' 不是 'Y' 或空字符串")
                return False
            if host_value not in ['Y', '']:
                print(f"错误：'是否为主持人'列的值 '{host_value}' 不是 'Y' 或空字符串")
                return False

        # 如果没有出错，返回二维字符串列表
        return data_list

    except Exception as e:
        print(f"处理替换词表.xlsx时发生错误：{e}")
        return False

def replace_list(data,re_list):
    # 遍历二维列表，处理第7列的字符串
    #print('正在替换词...')
    result = []  # 用于存储处理后的二维列表
    for row in data:
        processed_string = replace_word(row[COLUMN_NAME],process_re_list(re_list,True))
        row[COLUMN_NAME] = processed_string
            
        if isinstance(row[COLUMN_REF_DATA], str):  # 确保存在引用内容
            processed_string = replace_word(row[COLUMN_REF_DATA],process_re_list(re_list))  # 处理第8列的字符串（索引为7）
            row[COLUMN_REF_DATA] = processed_string  # 更新第8列的值

        if len(row) >= (COLUMN_SEND_DATA + 1):  # 确保当前行有至少7列
            processed_string = replace_word(row[COLUMN_SEND_DATA],process_re_list(re_list))  # 处理第7列的字符串（索引为6）
            if processed_string:  # 如果处理后的字符串不为空
                row[COLUMN_SEND_DATA] = processed_string  # 更新第7列的值
                result.append(row)  # 将处理后的行添加到结果列表
        else:
            result.append(row)  # 如果当前行不足7列，直接添加到结果列表
    print('替换词完成!')
    return result

def process_re_list(data, is_host = False):
    """
    处理二维字符串列表。
    如果is_host为True，删除第四列不为'Y'的行；
    如果is_host为False，删除第四列是'Y'的行。
    最后删除第四列。
    """
    # 筛选符合条件的行
    if is_host:
        # 保留第四列为'Y'的行
        filtered_data = [row for row in data if row[3] == 'Y']
    else:
        # 保留第四列不为'Y'的行
        filtered_data = [row for row in data if row[3] != 'Y']
    
    # 删除第四列
    result = [row[:3] + row[4:] for row in filtered_data if len(row) > 3]

    return result

def replace_word(input_string, re_list):

    # 遍历替换列表
    for item in re_list:
        original_word, replacement_word, flag= item
        if flag == 'Y':
            pattern = re.escape(original_word) + r'{2,}'
            input_string = re.sub(pattern, replacement_word, input_string)
        else:
            input_string = input_string.replace(original_word, replacement_word)
    
    return input_string

def normalize_2d_list(input_list):
    """
    遍历二维列表，删除字符串中的�字符。
    如果删除后为空，保留空字符串''。
    非字符串类型的元素保持不变。
    """
    #print('正在删除无效字符...')
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
    print('无效字符删除成功!')
    return normalized_list

def link_str(data):
    #print('正在生成导入的Markdown文件...')
    string = ''
    for row in data:
        string = string + row[COLUMN_NAME] + '（' + row[COLUMN_TIME][:5] + '）：' + row[COLUMN_SEND_DATA]
        if isinstance(row[COLUMN_REF_DATA], str):
            string = string + '\n<font style="color:' + FONT_COLOR + ';">引用内容：' + row[COLUMN_REF_DATA] + '</font>\n'
        else:
            string = string + '\n'
    string = string.replace('\n','\n\n')
    return string

def read_ai_cfg():
    try:
        # 读取Excel文件为DataFrame
        if IS_TEST:
            file_path = 'AI配置_Test.xlsx'
        else:
            file_path = 'AI配置.xlsx'
        
        # 读取“配置”sheet
        config_sheet = pd.read_excel(file_path, sheet_name='配置')

        # 读取“语义分析”sheet
        semantic_sheet = pd.read_excel(file_path, sheet_name='语义分析')

        # 将每个sheet的内容转化为字符串列表
        config_list = config_sheet.values.tolist()
        sa_list = semantic_sheet.values.tolist()

        ai_cfg = {}
        ai_cfg["key"] = config_list[LINE_KEY][1]
        ai_cfg["en"] = (config_list[LINE_EN_AI][1] == 'Y')
        ai_cfg["summ"] = (config_list[LINE_SUMM][1] == 'Y')
        ai_cfg["chap"] = (config_list[LINE_CHAP][1] == 'Y')
        ai_cfg["sa"] = (config_list[LINE_SA][1] == 'Y')
        ai_cfg["compara"] = (config_list[LINE_COMPARA][1] == 'Y')
        ai_cfg["cost"] = (config_list[LINE_COST][1] == 'Y')

        return ai_cfg,sa_list

    except Exception as e:
        print(f"处理AI配置.xlsx时发生错误：{e}")
        return False

def call_kimi_api(ai_cfg,user_message, file_path=None):
    """
    调用 Kimi API，支持可选的文件上传功能。

    参数：
    - api_key: 你的 Kimi API 密钥
    - user_message: 用户的提示或问题
    - file_path: 要上传的文件路径（可选）

    返回：
    - 回答内容
    - token 使用情况
    """
    # 初始化 OpenAI 客户端
    client = OpenAI(
        api_key=ai_cfg["key"],
        base_url="https://api.moonshot.cn/v1",
    )

    # 构建消息列表
    messages = []
    messages.append({"role": "user", "content": user_message})

    # 检查是否需要上传文件
    files = []
    if file_path:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件 {file_path} 不存在！")
        files.append(file_path)

    # 调用接口
    try:
        completion = client.chat.completions.create(
            model=AI_MODEL,
            messages=messages,
            files=files,  # 上传文件（如果有的话）
            temperature=0.3,
        )
    except Exception as e:
        return f"调用失败：{e}"

    return completion.choices[0].message.content, completion.usage.total_tokens

def create_md(file_name,data):
    with open(file_name, "w", encoding="utf-8") as file:
        file.write(data)

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

    ai_cfg,sa_list= read_ai_cfg()
    if ai_cfg["en"]:
        pass
    else:
        create_md("import.md",content)
        print('import.md生成成功!\n全部流程结束!')
    
if __name__ == "__main__":
    main_function()