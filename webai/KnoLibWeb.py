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
import json
import os
from functools import wraps

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

@app.route('/')
@requires_auth
def index():
    return open('index.html', encoding='utf-8').read()

@app.route('/upload_chat_record', methods=['POST'])
def upload_chat_record():
    return 

@app.route('/upload_replace', methods=['POST'])
def upload_replace():
    return 

@app.route('/execute_script', methods=['POST'])
def execute_script():
    lock_file = 'script.lock'
    if os.path.exists(lock_file):
        print("脚本正在运行，请勿重复调用")
        return "脚本正在运行，请勿重复调用"
    else:
        os.system(".venv/bin/python3 process_format.py")
        return " "
    
@app.route('/stop_script', methods=['POST'])
def stop_script():
    lock_file = 'script.lock'
    if os.path.exists(lock_file):
        os.remove(lock_file)
        print("脚本已强制停止！")
        return "脚本已经强制停止！"
    else:
        print("脚本未执行！")
        return "脚本未执行！"
    
@app.route('/start_server', methods=['POST'])
def start_server():
    with open('ipmi.json', 'r', encoding='utf-8') as file:
        data = json.load(file)
    os.system(f"ipmitool -H 127.0.0.1 -U {data['acct']} -P {data['pwd']} chassis power on")
    return "已执行开机操作"

@app.route('/download')
def download_file():
    return

@app.route('/error_log')
def error_log():
    log_file = "debug_info.html"
    return send_from_directory('./',log_file)

if __name__ == '__main__':
    app.run(port=5566)