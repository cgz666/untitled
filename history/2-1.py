import requests
import os
from bs4 import BeautifulSoup
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, Float, String, MetaData, Table
import re
from test.数据库 import sql_orm

# 配置模块
class Config:
    url = "http://omms.chinatowercom.cn:9000/business/resMge/alarmMge/listAlarm.xhtml"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": "route=bd2771ee9b3ccafe1de29698abbf009e; ULTRA_U_K=; JSESSIONID=DD4CEE454E5C68763C4FD7262C5E143F; acctId=101029143; uid=dw-wangcx9; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.105.20:all8180; BIGipServerywjk_new_pool1=342433196.10275.0000; Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1745809498,1745939912,1745993521,1746429292; Hm_lpvt_f6097524da69abc1b63c9f8d19f5bd5b=1746429292; HMACCOUNT=4DF837C069A2D971; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzQ2NDM5MTg3LCJOQU5PU0VDT05EIjozNTMxNTEzMjI1MDM4ODQ3NH0.1caGyCz3a1GUcielsU5FueLrjrKsYhPjKN67mx1ExCY",
        "Host": "omms.chinatowercom.cn:9000",
        "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
    }
    data_list = [
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:unitHidden": "0099977",
            "queryForm:selectSignalSize": "0",
            "queryForm:alarmNameMax": "15",
            "queryForm:firststarttimeInputCurrentDate": "05/2025",
            "queryForm:firstendtimeInputCurrentDate": "05/2025",
            "queryForm:querystationstatus_hiddenValue": "2",
            "queryForm:querystationstatus": "2",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id4",
            "queryForm:btn1": "queryForm:btn1"
        },
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:unitHidden": "0099977",
            "queryForm:selectSignalSize": "0",
            "queryForm:alarmNameMax": "15",
            "queryForm:firststarttimeInputCurrentDate": "05/2025",
            "queryForm:firstendtimeInputCurrentDate": "05/2025",
            "queryForm:querystationstatus_hiddenValue": "2",
            "queryForm:querystationstatus": "2",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id4",
            "queryForm:j_id166": "queryForm:j_id166",
            "AJAX:EVENTS_COUNT": "1"
        },
        {
            "j_id1547": "j_id1547",
            "j_id1547:j_id1552": "当页",
            "javax.faces.ViewState": "j_id4"
        }
    ]
    save_path = "C:/Users/27569/Desktop/新建文件夹/活动告警.xlsx"  # 直接指定完整路径和文件名

# 数据库连接配置
db_config = {
    'host': 'localhost',  # 数据库主机地址
    'user': 'cgz',        # 数据库用户名
    'password': '12345678',  # 数据库密码
    'database': 'my_crawled_data',  # 数据库名称
    'charset': 'utf8mb4'  # 指定字符集
}

# 请求模块
def get_view_state(url, headers):
    """获取页面的javax.faces.ViewState值"""
    res = requests.get(url=url, headers=headers)
    soup = BeautifulSoup(res.text, 'html.parser')
    view_state_input = soup.find('input', id='javax.faces.ViewState')
    if view_state_input:
        return view_state_input.get('value')
    return None

def post_data(url, headers, data):
    """发送POST请求"""
    res = requests.post(url=url, data=data, headers=headers)
    return res

# 数据处理模块
def save_response_to_file(response, file_path):
    """将响应内容保存到文件"""
    if not os.path.exists(os.path.dirname(file_path)):
        os.makedirs(os.path.dirname(file_path))
    with open(file_path, "wb") as file:
        file.write(response.content)
    print("文件已成功保存到:", file_path)

# 主程序模块
def main():
    config = Config
    url = config.url
    headers = config.headers
    data_list = config.data_list
    save_path = config.save_path

    # 获取javax.faces.ViewState值
    view_state = get_view_state(url, headers)
    if not view_state:
        print("无法获取javax.faces.ViewState值，请检查URL和headers配置")
        return

    # 遍历数据列表，发送POST请求
    for i, data in enumerate(data_list, start=1):
        data["javax.faces.ViewState"] = view_state
        response = post_data(url, headers, data)

        # 如果是最后一个请求，保存响应内容到文件
        if i == len(data_list):
            save_response_to_file(response, save_path)

            with sql_orm().session_scope() as temp:
                session, Base = temp
                pojo = Base.classes.station
                rows = []
                for index, row in df.iterrows():
                    temp = pojo(**row.to_dict())
                    rows.append(temp)
                session.bulk_save_objects(rows)

if __name__ == "__main__":
    main()
