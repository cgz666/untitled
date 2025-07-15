# coding=utf-8
import os
import sys
sys.path.append('F:/newtowerV2')
sys.path.append(r'F:/newtowerV2\venv\Lib\site-packages')
import shutil
import re
import pandas as pd
from utils.mysql_connecter import sql_orm
from config import IP_SERVICE,INDEX,SPIDER_PATH,TEMP_PATH_ONE_DAY
import datetime
from flask import Flask, request,render_template, send_file,redirect,session
from flask_cors import CORS
from flask_session import Session
from redis import StrictRedis
from run_thread.other_task import gen_fsu_static
from sqlalchemy import text
import tempfile
from websource.spider.down_foura.down_yisi_his import main as gen_static_yisi_down
import json

app = Flask(__name__, template_folder=f"{INDEX}service/templates")
CORS(app, resources=r'/*')
app.secret_key = 'xgxtt'
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = datetime.timedelta(seconds=1)
app.config['SESSION_TYPE'] = 'redis'
app.static_folder = 'templates'
# 创建Redis客户端并设置密码
redis_host = 'localhost'
redis_port = 6379
redis_password = 123456  # 按实际配置设置密码
redis_client = StrictRedis(host=redis_host, port=redis_port, password=redis_password)
app.config['SESSION_REDIS'] = redis_client
# 初始化会话扩展
Session(app)
down_path = f"{INDEX}service/websource/picture/"

def zip_file_and_send(folder,file_list):
    zip_folder=os.path.join(folder,'zip')
    for file in os.listdir(zip_folder):
        file = os.path.join(zip_folder, file)
        os.remove(file)
    for file in file_list:
        path=os.path.join(folder,file)
        zip_path=os.path.join(zip_folder,file)
        shutil.copy(path,zip_path)
    zip_path=f"{TEMP_PATH_ONE_DAY}{str(datetime.datetime.now().timestamp())}"
    shutil.make_archive(zip_path, 'zip', zip_folder)
    return send_file(zip_path + '.zip', as_attachment=True, cache_timeout=0)

@app.route('/index', methods=['get'])
def index():
    with sql_orm().session_scope() as temp:
        session,Base=temp
        pojo=Base.classes.update_downhour_log
        listt={}
        try:
            res=session.query(pojo).all()
            for item in res:
                listt[item.type]=item.time
            print(1)
        except Exception as e:print(e)
    return render_template('index.html',**listt)

@app.route('/get_4a_cookie', methods=['get'])
def get_aiot_cookie():
    try:
        res=sql_orm().get_cookie('foura')
        return json.dumps(res)
    except Exception as e:
        print(str(e))
if __name__ == '__main__':
    app.run(host=IP_SERVICE,debug=False,port=5000)
