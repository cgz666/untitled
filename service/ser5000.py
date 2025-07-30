from flask import Flask, send_file, render_template, jsonify
import os
import mysql.connector
from mysql.connector import Error
from utils.sql import sql_orm
from flask_cors import CORS
from config import FILES

app = Flask(__name__)  # 默认使用项目根目录下的 templates 文件夹
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 1
CORS(app, resources={r"/": {"origins": "http://localhost:5173"}})

# MySQL 数据库配置
DB_HOST = "10.19.6.250"
DB_USER = "root"
DB_PASSWORD = "123456"
DB_NAME = "tower"

def get_db_connection():
    connection = None
    try:
        connection = mysql.connector.connect(
            host=DB_HOST,
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME
        )
        print("MySQL Database connection successful")
    except Error as err:
        print(f"Error: '{err}'")
    return connection

def get_last_modified_time(filename):
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT time FROM update_downhour_log WHERE type = %s ORDER BY time DESC LIMIT 1"
    cursor.execute(query, (filename,))
    result = cursor.fetchone()
    cursor.close()
    conn.close()
    return result[0] if result else None

@app.route('/')
@app.route('/index.html')  # 添加对index.html的支持
def index():
    # 假设 FILES 是一个包含文件名和路径的字典
    FILE = FILES
    files = []
    for filename, file_info in FILE.items():
        last_modified = get_last_modified_time(filename)
        files.append({
            'name': filename,
            'description': f"{filename} (最近更新时间：{last_modified})",
            'last_modified': last_modified,
            'key': filename
        })
    return render_template('index.html', files=files)

@app.route('/download.html')  # 添加对download.html的支持
def download():
    # 假设 FILES 是一个包含文件名和路径的字典
    FILE = FILES
    files = []
    for filename, file_info in FILE.items():
        last_modified = get_last_modified_time(filename)
        files.append({
            'name': filename,
            'description': f"{filename} (最近更新时间：{last_modified})",
            'last_modified': last_modified,
            'key': filename
        })
    return render_template('download.html', files=files)


@app.route('/quxin_order_youji', methods=['get'])
def quxin_order_youji():
    path = r'F:\untitled\quxin_script\data_spider\gongdan_youji_pipei\output\工单及油机匹配率统计(06-26).xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('工单及油机匹配率统计' + str(e))

@app.route('/quxin_yidong_offline', methods=['get'])
def quxin_yidong_offline():
    path = r'F:\untitled\quxin_script\data_spider\yidong_tuifu_static\output\移动无运营商推送退服类告警统计(06-18).xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('移动无运营商推送退服类告警统计' + str(e))

@app.route('/Interface_order', methods=['get'])
def Interface_order():
    path = r'F:\untitled\four_a_script\Interface_result\output\结果.xls'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('运营商接口工单' + str(e))

@app.route('/battery_order', methods=['get'])
def battery_order():
    path = r'F:\untitled\four_a_script\Performance_query\output\结果.xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('开关电源电压设置修改情况通报' + str(e))

@app.route('/environment_order', methods=['get'])
def environment_order():
    path = r'F:\untitled\environment_script\data_spider\output\委托备案数据与环保厅备案数据匹配情况统计-结果_筛选后.xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('委托备案数据与环保厅备案数据匹配情况统计' + str(e))

@app.route('/online_rate_order', methods=['get'])
def online_rate_order():
    path = r'F:\untitled\four_a_script\Average_online_rate\output\智联设备平均在线率统计-结果.xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('智联设备平均在线率统计' + str(e))
@app.route('/Custom_Workorder_zl', methods=['get'])
def Custom_Workorder_zl():
    path = r'F:\untitled\four_a_script\Custom_Workorder_zl\output\自定义工单-结果.xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('自定义工单_智联' + str(e))

@app.route('/Signal_strength', methods=['get'])
def Signal_strength():
    path = r'F:\newtowerV2\websource\spider_download\Comprehensive_query\信号强度\信号强度_结果.xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('信号强度' + str(e))

@app.route('/Device_management', methods=['get'])
def Device_management():
    path = r'F:\newtower3.8\project\four_a_script\Device_management\output\设备信息-结果.xlsx'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('智联设备信息' + str(e))

@app.route('/Custom_Workorder_yys', methods=['get'])
def Custom_Workorder_yys():
    path = r'F:\newtower3.8\project\four_a_script\Custom_Workorder_yys\output\自定义工单.zip'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('自定义工单_运营商' + str(e))

@app.route('/direct_current', methods=['GET'])
def direct_current():
    path = r'F:\newtowerV2\websource\spider_download\Comprehensive_query\直流电压直流负载总电流\直流电压\全站址直流电压.xlsx'
    # 获取文件的修改时间
    modify_time = os.path.getmtime(path)
    # 返回文件下载和修改时间
    response = send_file(path, as_attachment=True)
    response.headers['X-Modify-Time'] = str(modify_time)
    return response

@app.route('/Unicom_analysis', methods=['get'])
def Unicom_analysis():
    path = r'F:\newtower3.8\project\four_a_script\Unicom_analysis\output\联通分析_结果.zip'
    try:
        return send_file(path, as_attachment=True)
    except Exception as e:
        print('联通分析' + str(e))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)