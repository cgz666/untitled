import requests
import os
from bs4 import BeautifulSoup
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, Float, String, MetaData, Table
import re

# 配置模块
class Config:
    url = "http://omms.chinatowercom.cn:9000/business/resMge/alarmHisHbaseMge/listHisAlarmHbase.xhtml"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": "Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1745809498,1745939912,1745993521,1746429292; route=509b7f7141a6484a386a59b699b26db9; JSESSIONID=219CF9D2451D62C0CAE8AF2753C1F9D2; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzQ2NjY4MjYyLCJOQU5PU0VDT05EIjozNTU0NDIwNjg0NjU4NjY4NH0.VgK9XU-ZSIQo0FgUUo3pBraDDBCh-tZDwSkJvDdzynA; acctId=101429653; uid=wx-huangwl14; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.105.18:all8380; BIGipServerywjk_new_pool1=42016172.10275.0000",
        "Host": "omms.chinatowercom.cn:9000",
        "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
    }
    data_list = [
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:proviceIdHidden": "0098364",
            "queryForm:cityIdHidden": "undefined",
            "queryForm:countryIdHidden": "undefined",
            "queryForm:faultidText": "",
            "queryForm:j_id17": "",
            "queryForm:j_id33": "",
            "queryForm:j_id37": "",
            "queryForm:firststarttimeInputDate": "2025-05-01 00:00",
            "queryForm:firststarttimeInputCurrentDate": "05/2025",
            "queryForm:firstendtimeInputDate": "2025-05-02 00:00",
            "queryForm:firstendtimeInputCurrentDate": "05/2025",
            "queryForm:queryalarmName": "一级低压脱离告警",
            "queryForm:deleteproviceIdHidden": "",
            "queryForm:deletecityIdHidden": "",
            "queryForm:deletecountryIdHidden": "",
            "queryForm:queryDeleteCountyName": "",
            "queryForm:querySpeId": "",
            "queryForm:querySpeIdShow": "",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "15",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id4",
            "queryForm:j_id55": "queryForm:j_id55"
        },
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:proviceIdHidden": "0098364",
            "queryForm:cityIdHidden": "undefined",
            "queryForm:countryIdHidden": "undefined",
            "queryForm:faultidText": "",
            "queryForm:j_id17": "",
            "queryForm:j_id33": "",
            "queryForm:j_id37": "",
            "queryForm:firststarttimeInputDate": "2025-05-01 00:00",
            "queryForm:firststarttimeInputCurrentDate": "05/2025",
            "queryForm:firstendtimeInputDate": "2025-05-02 00:00",
            "queryForm:firstendtimeInputCurrentDate": "05/2025",
            "queryForm:queryalarmName": "一级低压脱离告警",
            "queryForm:deleteproviceIdHidden": "",
            "queryForm:deletecityIdHidden": "",
            "queryForm:deletecountryIdHidden": "",
            "queryForm:queryDeleteCountyName": "",
            "queryForm:querySpeId": "",
            "queryForm:querySpeIdShow": "",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "15",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id4",
            "queryForm:j_id56": "queryForm:j_id56",
            "AJAX:EVENTS_COUNT": "1"
        },
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:proviceIdHidden": "0098364",
            "queryForm:cityIdHidden": "undefined",
            "queryForm:countryIdHidden": "undefined",
            "queryForm:faultidText": "",
            "queryForm:j_id17": "",
            "queryForm:j_id33": "",
            "queryForm:j_id37": "",
            "queryForm:firststarttimeInputDate": "2025-05-01 00:00",
            "queryForm:firststarttimeInputCurrentDate": "05/2025",
            "queryForm:firstendtimeInputDate": "2025-05-02 00:00",
            "queryForm:firstendtimeInputCurrentDate": "05/2025",
            "queryForm:queryalarmName": "一级低压脱离告警",
            "queryForm:deleteproviceIdHidden": "",
            "queryForm:deletecityIdHidden": "",
            "queryForm:deletecountryIdHidden": "",
            "queryForm:queryDeleteCountyName": "",
            "queryForm:querySpeId": "",
            "queryForm:querySpeIdShow": "",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "1",
            "queryForm:pageSizeText": "15",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id4",
            "queryForm:j_id60": "queryForm:j_id60"
        },
        {
            "j_id313": "j_id313",
            "j_id313:devExport": "全部",
            "javax.faces.ViewState": "j_id4"
        }
    ]
    save_path = "C:/Users/27569/Desktop/for test/历史告警.xlsx"  # 直接指定完整路径和文件名

# 数据库连接配置
db_config = {
    'host': '10.19.6.250',  # 数据库主机地址
    'user': 'root',        # 数据库用户名
    'password': '123456',  # 数据库密码
    'database': 'test',    # 数据库名称
    'charset': 'utf8mb4'   # 指定字符集
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

# 数据处理模块
def save_response_to_file(response, save_path):
    """将响应内容保存到文件"""
    if not os.path.exists(os.path.dirname(save_path)):
        os.makedirs(os.path.dirname(save_path))
    with open(save_path, "wb") as file:
        file.write(response.content)
    print(f"文件已成功保存到: {save_path}")

def clean_column_name(column_name):
    """清理列名，去除特殊字符"""
    return re.sub(r'\W+', '_', column_name).strip('_')

def create_table_from_excel(save_path, db_config):
    """从Excel文件中读取数据并保存到数据库"""
    try:
        # 读取Excel文件
        df = pd.read_excel(save_path)
        print("Excel文件读取成功")
        print(df.head())  # 打印前几行数据

        # 清理列名
        df.columns = [clean_column_name(col) for col in df.columns]

        # 创建数据库连接
        engine = create_engine(f"mysql+pymysql://{db_config['user']}:{db_config['password']}@{db_config['host']}/{db_config['database']}")

        # 获取表名，这里以文件名作为表名
        table_name = os.path.splitext(os.path.basename(save_path))[0]

        # 动态创建表结构
        metadata = MetaData()
        columns = []
        for column in df.columns:
            if pd.api.types.is_integer_dtype(df[column]):
                columns.append(Column(column, Integer))
            elif pd.api.types.is_float_dtype(df[column]):
                columns.append(Column(column, Float))
            else:
                columns.append(Column(column, String(255)))

        # 创建表
        table = Table(table_name, metadata, *columns)
        metadata.create_all(engine)

        # 插入数据
        df.to_sql(table_name, engine, if_exists='append', index=False)
        print("数据已成功插入到数据库")

    except Exception as e:
        print(f"操作失败: {e}")

def insert_data_to_db(df, db_config, table_name):
    """将数据插入数据库，并在失败时输出异常信息"""
    try:
        # 创建数据库连接
        engine = create_engine(f"mysql+pymysql://{db_config['user']}:{db_config['password']}@{db_config['host']}/{db_config['database']}")

        # 插入数据
        df.to_sql(table_name, engine, if_exists='append', index=False)
        print("数据已成功插入到数据库")
    except Exception as e:
        print(f"数据插入失败: {e}")

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
        response = requests.post(url=url, data=data, headers=headers)

        # 如果是最后一个请求，保存响应内容到文件
        if i == len(data_list):
            save_response_to_file(response, save_path)

    # 尝试读取文件并插入数据库
    try:
        # 读取Excel文件
        df = pd.read_excel(save_path)
        print("Excel文件读取成功")
        print(df.head())  # 打印前几行数据

        # 清理列名
        df.columns = [clean_column_name(col) for col in df.columns]

        # 获取表名，这里以文件名作为表名
        table_name = os.path.splitext(os.path.basename(save_path))[0]

        # 插入数据到数据库
        insert_data_to_db(df, db_config, table_name)

    except Exception as e:
        print(f"文件读取或数据插入失败: {e}")

if __name__ == "__main__":
    main()