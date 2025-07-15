import requests
import os
from bs4 import BeautifulSoup
import requests
import json

# 获取cookie的URL
cookie_url = "http://10.19.6.250:5000/get_4a_cookie"

# 发送GET请求获取cookie
res = requests.get(cookie_url)

# 检查请求是否成功
if res.status_code == 200:
    cookie_str = res.text.strip()  # 获取响应内容并去除两端的空白字符

    # 将JSON格式的Cookie字符串解析为字典
    cookie_dict = json.loads(cookie_str)

    # 将字典转换为Cookie字符串格式
    cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])


# 配置模块
class Config:
    url = "http://omms.chinatowercom.cn:9000/business/resMge/alarmMge/listAlarm.xhtml"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": f"{cookie_header}",
        "Host": "omms.chinatowercom.cn:9000",
        "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
    }
    data_list = [
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:selectSignalSize": "0",
            "queryForm:alarmNameMax": "15",
            "queryForm:firststarttimeInputCurrentDate": "04/2025",
            "queryForm:firstendtimeInputCurrentDate": "04/2025",
            "queryForm:querystationstatus_hiddenValue": "2",
            "queryForm:querystationstatus": "2",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id3",
            "queryForm:btn1": "queryForm:btn1"
        },
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:selectSignalSize": "0",
            "queryForm:alarmNameMax": "15",
            "queryForm:firststarttimeInputCurrentDate": "04/2025",
            "queryForm:firstendtimeInputCurrentDate": "04/2025",
            "queryForm:querystationstatus_hiddenValue": "2",
            "queryForm:querystationstatus": "2",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id3",
            "queryForm:j_id166": "queryForm:j_id166",
            "AJAX:EVENTS_COUNT": "1"
        },

        {
            "j_id1547": "j_id1547",
            "j_id1547:j_id1552": "当页",
            "javax.faces.ViewState": "j_id3"
        }
    ]
    save_path = r"E:\陈桂志\普通文件"  # 保存路径
    file_name = "notice2.xlsx"

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
def save_response_to_file(response, save_path, file_name):
    """将响应内容保存到文件"""
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    file_path = os.path.join(save_path,file_name)
    with open(file_name, "wb") as file:
        file.write(response.content)
    print("文件已成功保存到:", {file_path})


# 主程序模块
def main():
    url = Config.url
    headers = Config.headers
    data_list = Config.data_list
    save_path = Config.save_path
    file_name = Config.file_name

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
            save_response_to_file(response, save_path, file_name)


if __name__ == "__main__":
    main()