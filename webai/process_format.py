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
import os
import time
import ollama
import socket
import sys

VERSION = 2.0
IS_TEST = False #是否是测试环境，使用时改为False
FONT_COLOR = '#8A8F8D' #灰色的HEX表示值

COLUMN_TYPE = 3         #类型所在的列数，从0开始数
COLUMN_SEND_DATA = 6    #发送内容所在的列数，从0开始数
COLUMN_REF_DATA = 7     #引用内容所在的列数，从0开始数
COLUMN_NAME = 4         #昵称内容所在的列数，从0开始数
COLUMN_TIME = 2         #时刻所在的列数，从0开始数
COLUMN_MULTI = 2        #连续才删除所在的列，从0开始数
COLUMN_HOST = 3         #是否未主持人所在的列，从0开始数

LINE_EN_AI = 0          #启用AI所在的行索引，行索引=行数-2
LINE_HOST= 1            #Ollama Host所在的行索引，行索引=行数-2
LINE_PORT= 2            #Ollama Port所在的行索引，行索引=行数-2
LINE_MODEL= 3           #Ollama Model所在的行索引，行索引=行数-2
LINE_SA = 4             #利用语义分析删除相关内容所在的行索引，行索引=行数-2
LINE_SCORE = 5          #显示语义分析打分及内容所在的行索引，行索引=行数-2
LINE_COMPARA = 6        #生成AI处理前后对比文档所在的行索引，行索引=行数-2
HTML_PATH = "debug_info.html"

LOCK_FILE = 'script.lock'

def check_tcp_connection(ip_address, port, timeout=2):
    """
    检查指定IP和端口的TCP连接是否通
    
    参数:
        ip_address: 要检查的IP地址
        port: 要检查的端口号
        timeout: 连接超时时间(秒)，默认2秒
    
    返回:
        如果连接成功，返回True，否则返回False
    """
    try:
        # 创建一个TCP socket
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            # 设置超时时间
            s.settimeout(timeout)
            
            # 尝试连接到指定的IP和端口
            result = s.connect_ex((ip_address, port))
            
            # connect_ex返回0表示连接成功，其他值表示失败
            return result == 0
    
    except socket.error as e:
        print(f"Socket error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")
    
    return False

# 自定义打印函数
def custom_print(*args, **kwargs):
    # 打印到控制台
    print(*args, **kwargs)
    # 打印到 HTML 文件
    with open(HTML_PATH, "a", encoding="utf-8") as html_file:
        # 将参数转换为字符串并写入文件
        html_file.write(" ".join(map(str, args)) + "\n")

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
    
    file_path = 'knfile/聊天记录.xlsx'
    
    # 如果用户取消选择，直接退出
    if not file_path:
        custom_print("未选择文件，程序退出。")
        return False
    
    # 使用pandas读取Excel文件
    try:
        df = pd.read_excel(file_path, dtype=str)  # 确保所有数据都以字符串形式读取
        # 将DataFrame转换为二维字符串列表
        string_list = df.values.tolist() 
        return string_list
    
    except Exception as e:
        custom_print(f"读取Excel文件时出错：{e}")
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
    custom_print('类别为图片的行删除成功!')
    return normalized_data

def read_replace():
    try:
        # 读取Excel文件为DataFrame
        df = pd.read_excel('knfile/替换词表.xlsx', dtype=str)  # 将所有数据读取为字符串类型
        df.fillna('', inplace=True)  # 将所有空值替换为空字符串

        # 将DataFrame转换为二维字符串列表
        data_list = df.values.tolist()

        for row in data_list:
            multi_value = row[COLUMN_MULTI]
            host_value = row[COLUMN_HOST]

            if multi_value not in ['Y', '']:
                custom_print(f"错误：'连续才删除'列的值 '{multi_value}' 不是 'Y' 或空字符串")
                return False
            if host_value not in ['Y', '']:
                custom_print(f"错误：'是否为主持人'列的值 '{host_value}' 不是 'Y' 或空字符串")
                return False

        # 如果没有出错，返回二维字符串列表
        return data_list

    except Exception as e:
        custom_print(f"处理替换词表.xlsx时发生错误：{e}")
        return False

def replace_list(data,re_list):

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
    custom_print('替换词完成!')
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
    custom_print('无效字符删除成功!')
    return normalized_list

def link_str(data):

    string = ''
    for row in data:
        string = string + row[COLUMN_NAME] + '(' + row[COLUMN_TIME][:5] + ')：' + row[COLUMN_SEND_DATA]
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
            file_path = 'knfile/AI配置_Test.xlsx'
        else:
            file_path = 'knfile/AI配置.xlsx'
        
        # 读取“配置”sheet
        config_sheet = pd.read_excel(file_path, sheet_name='配置')

        # 读取“语义分析”sheet
        semantic_sheet = pd.read_excel(file_path, sheet_name='语义分析')

        # 将每个sheet的内容转化为字符串列表
        config_list = config_sheet.values.tolist()
        sa_list = semantic_sheet.values.tolist()

        ai_cfg = {}
        ai_cfg["en"] = (config_list[LINE_EN_AI][1] == 'Y')
        ai_cfg["sa"] = (config_list[LINE_SA][1] == 'Y')
        ai_cfg["compara"] = (config_list[LINE_COMPARA][1] == 'Y')
        ai_cfg["host"] = config_list[LINE_HOST][1]
        ai_cfg["port"] = str(config_list[LINE_PORT][1])
        ai_cfg["model"] = config_list[LINE_MODEL][1]
        ai_cfg["score"] = (config_list[LINE_SCORE][1] == 'Y')

        return ai_cfg,sa_list

    except Exception as e:
        custom_print(f"处理AI配置.xlsx时发生错误：{e}")
        return False

def call_ollama(ai_cfg,content):
    host = ai_cfg["host"]
    port = ai_cfg["port"]
    client= ollama.Client(host=f"http://{host}:{port}")
    res=client.chat(model=ai_cfg["model"],messages=[{"role": "user","content":content}],options={"temperature":0})
    return res['message']['content']

def call_deepseek(ai_cfg,content):
    answer = call_ollama(ai_cfg,content)
    cleaned_text = re.sub(r'<think>.*?</think>', '', answer, flags=re.DOTALL)
    return cleaned_text
 
def create_md(file_name,data):
    with open(file_name, "w", encoding="utf-8") as file:
        file.write(data)

def delete_indices_from_list(indices, target_list):
    """
    根据给定的索引列表删除目标列表中对应索引的元素。

    参数:
        indices (list): 包含要删除的索引的列表。
        target_list (list): 要从中删除元素的目标列表。

    返回:
        list: 删除指定索引元素后的新列表。
    """
    # 使用集合去重，确保每个索引只处理一次
    unique_indices = set(indices)
    
    # 从后往前删除，避免索引错位
    for index in sorted(unique_indices, reverse=True):
        if 0 <= index < len(target_list):  # 确保索引在有效范围内
            del target_list[index]
        else:
            custom_print(f"警告：索引 {index} 超出目标列表范围，已忽略。")
    
    return target_list

def process_sa(ai_cfg,sa_list,txt_list):
    if ai_cfg["sa"] == False:
        return txt_list
    
    del_num = []
    for row in sa_list:
        custom_print('正在进行"' + row[0] + '"的语义分析及删除...')
        for index, row2 in enumerate(txt_list):
            message = '请判断下列文字是不是关于' + row[0] + '的信息,仅回答为0到10的数字,0为肯定不是,10为肯定是。注意，仅回答数字！文字为:' + row2[COLUMN_SEND_DATA]
            answer = call_deepseek(ai_cfg,message)

            # 提取数字
            number = re.search(r'\d+', answer)
            if number:
                if ai_cfg["score"]:
                    custom_print("得分：" + number.group())
                    custom_print(row2[COLUMN_SEND_DATA])
                if float(number.group()) >= float(row[1]):
                    custom_print('"' + row2[COLUMN_SEND_DATA] + '"已被删除')
                    del_num.append(index)
            else:
                custom_print("No number found.")
    
    txt_list = delete_indices_from_list(del_num,txt_list)
    custom_print('语义分析完成!')
    return txt_list
    
            
def main_function():

    if os.path.exists(LOCK_FILE):
        #custom_print("脚本正在运行，请勿重复调用")
        sys.exit()
    else:
        with open(LOCK_FILE, 'w') as f:
            f.write(str(os.getpid()))

    # 打开 HTML 文件并写入头部信息
    with open(HTML_PATH, "w") as html_file:
        html_file.write('<html><head><title>Debug Information</title><meta http-equiv="refresh" content="8"></head><body><pre>')

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
        txt_list = process_sa(ai_cfg,sa_list,txt_list)
        save_list_to_csv(txt_list,'ai.csv')
        if ai_cfg["compara"]:
            create_md("import_AI处理前.md",content)
            content = link_str(txt_list)
            create_md("import_AI处理后.md",content)
            custom_print('import_AI处理前/处理后.md生成成功!\n全部流程结束!')
        else:
            content = link_str(txt_list)
            create_md("import.md",content)
            custom_print('import.md生成成功!\n全部流程结束!')
    else:
        create_md("import.md",content)
        custom_print('import.md生成成功!\n全部流程结束!')
    
if __name__ == "__main__":
    main_function()