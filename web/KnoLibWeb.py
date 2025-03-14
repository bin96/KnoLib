# -*- coding: utf-8 -*-
"""
Copyright (c) 2025, bin96
All rights reserved.

This script is licensed under the MIT License.
See LICENSE file for details.

Description:
The function of this script is to perform data processing
"""

from flask import Flask, request, send_from_directory, Response
import os
import pandas as pd
from tkinter import Tk, filedialog
import csv
import re
import os
from functools import wraps

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

HTML_PATH = "logs/debug_info.html"
CHAT_PATH = 'uploads/聊天记录.xlsx'
RE_PATH = 'uploads/替换词表.xlsx'

app = Flask(__name__)


def check_auth(username, password):
    return username == 'know' and password == 'know'

def authenticate():
    return Response(
        'Could not verify your access level for that URL.\n'
        'You have to login with proper credentials', 401,
        {'WWW-Authenticate': 'Basic realm="Login Required"'})

def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# 设置上传文件的保存目录
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 设置结果文件的保存目录
RESULT_FOLDER = 'results'
if not os.path.exists(RESULT_FOLDER):
    os.makedirs(RESULT_FOLDER)

app.config['RESULT_FOLDER'] = RESULT_FOLDER

# 设置错误日志文件的保存目录
LOG_FOLDER = 'logs'
if not os.path.exists(LOG_FOLDER):
    os.makedirs(LOG_FOLDER)

app.config['LOG_FOLDER'] = LOG_FOLDER

@app.route('/')
@requires_auth
def index():
    if os.path.exists('results/import.md'):
        os.remove('results/import.md')

    if os.path.exists(HTML_PATH):
        os.remove(HTML_PATH)

    if os.path.exists(CHAT_PATH):
        os.remove(CHAT_PATH)

    if os.path.exists(RE_PATH):
        os.remove(RE_PATH) 
    return open('index.html', encoding='utf-8').read()

@app.route('/upload_chat_record', methods=['POST'])
def upload_chat_record():
    if 'file' not in request.files:
        return "没有文件部分"
    file = request.files['file']
    if file.filename == '':
        return "没有选择文件"
    if file:
        filename = file.filename
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        return f"聊天记录文件 {filename} 已上传成功"
    return "上传失败"

@app.route('/upload_replace', methods=['POST'])
def upload_replace():
    if 'file' not in request.files:
        return "没有文件部分"
    file = request.files['file']
    if file.filename == '':
        return "没有选择文件"
    if file:
        filename = file.filename
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        return f"替换成表文件 {filename} 已上传成功"
    return "上传失败"

@app.route('/execute_script', methods=['POST'])
def execute_script():
    # 在这里调用你的 Python 脚本
    # 假设脚本会处理上传的文件并生成结果文件
    if not os.path.exists(CHAT_PATH):
        return "聊天记录.xlsx 未上传!"
    if not os.path.exists(RE_PATH):
        return "替换词表.xlsx 未上传!"
    main_function()
    print("脚本已执行")
    return "脚本执行成功"

@app.route('/download')
def download_file():
    result_file = "import.md"
    return send_from_directory(app.config['RESULT_FOLDER'], result_file, as_attachment=True)

@app.route('/error_log')
def error_log():
    log_file = "debug_info.html"
    return send_from_directory(app.config['LOG_FOLDER'], log_file)

# 自定义打印函数
def custom_print(*args, **kwargs):
    # 打印到控制台
    print(*args, **kwargs)
    # 打印到 HTML 文件
    with open(HTML_PATH, "a") as html_file:
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

        custom_print(f"数据已成功保存到文件 {filename}")

def read_excel_to_string_list():
   
    # 使用pandas读取Excel文件
    try:
        df = pd.read_excel(CHAT_PATH, dtype=str)  # 确保所有数据都以字符串形式读取
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
    # 使用列表推导式过滤掉符合条件的行
    normalized_data = [row for row in data if len(row) < 4 or '图片' not in row[COLUMN_TYPE]]
    custom_print('类别为图片的行删除成功!')
    return normalized_data

def read_replace():
    try:
        # 读取Excel文件为DataFrame
        df = pd.read_excel(RE_PATH, dtype=str)  # 将所有数据读取为字符串类型
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
    custom_print('无效字符删除成功!')
    return normalized_list

def link_str(data):
    #print('正在生成导入的Markdown文件...')
    string = ''
    for row in data:
        string = string + row[COLUMN_NAME] + '(' + row[COLUMN_TIME][:5] + ')：' + row[COLUMN_SEND_DATA]
        if isinstance(row[COLUMN_REF_DATA], str):
            string = string + '\n<font style="color:' + FONT_COLOR + ';">引用内容：' + row[COLUMN_REF_DATA] + '</font>\n'
        else:
            string = string + '\n'
    string = string.replace('\n','\n\n')
    return string

def create_md(file_name,data):
    with open(file_name, "w", encoding="utf-8") as file:
        file.write(data)
            
def main_function():

    # 打开 HTML 文件并写入头部信息
    with open(HTML_PATH, "w") as html_file:
        html_file.write("<html><head><title>Debug Information</title></head><body><pre>")

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
    create_md("results/import.md",content)
    custom_print('import.md生成成功!\n全部流程结束!')

    # 写入 HTML 文件的尾部信息
    with open(HTML_PATH, "a") as html_file:
        html_file.write("</pre></body></html>")

    if os.path.exists(CHAT_PATH):
        # 删除文件
        os.remove(CHAT_PATH)

    if os.path.exists(RE_PATH):
        # 删除文件
        os.remove(RE_PATH)        

if __name__ == '__main__':
    app.run(port=5555)