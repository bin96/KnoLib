
# -*- coding: utf-8 -*-
import os
import requests
import shutil
import sys
import importlib.util

# 脚本B的本地路径
SCRIPT_B_PATH = "process_format.py"
# 更新文件的下载地址
UPDATE_URL = "https://gitee.com/bin96/KnoLib/raw/master/process_format.py"
# 服务器上的最新版本号地址
VERSION_URL = "https://gitee.com/bin96/KnoLib/raw/master/version"

def check_for_update():
    """检查是否有新版本"""
    try:
        current_version = 0
        try:
            spec = importlib.util.spec_from_file_location("script_b", SCRIPT_B_PATH)
            script_b = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(script_b)
            current_version = script_b.get_version()  # 调用脚本B中的main_function
        except Exception as e:
            print(f"process_format.py文件不正确或者不存在!")
        response = requests.get(VERSION_URL)
        latest_version = response.text.strip()
        return float(latest_version) > current_version
    except Exception as e:
        print(f"Error checking for updates: {e}")
        return False

def download_update():
    """下载最新版本的脚本B"""
    try:
        response = requests.get(UPDATE_URL)
        with open(SCRIPT_B_PATH, "wb") as f:
            f.write(response.content)
        print("更新完成.")
    except Exception as e:
        print(f"Error downloading update: {e}")

def run_script_b_function():
    """动态加载脚本B并运行其函数"""
    try:
        spec = importlib.util.spec_from_file_location("script_b", SCRIPT_B_PATH)
        script_b = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(script_b)
        script_b.main_function()  # 调用脚本B中的main_function
    except Exception as e:
        print(f"Error running Script B's function: {e}")

if __name__ == "__main__":
    if check_for_update():
        print("检查发现新版本.更新中...")
        download_update()
        run_script_b_function()
    else:
        print("当前为最新版本.")
        run_script_b_function()
    input('按任意键退出...')