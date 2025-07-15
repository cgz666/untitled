import requests
import os
from bs4 import BeautifulSoup


# 配置模块
class Config:
    url = "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/fsuMge/listFsu.xhtml"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": "route=7634b5a86f92649382e66b47cf771393; ULTRA_U_K=; JSESSIONID=E612AD76AACBF01B21DE44713E66F5EF; acctId=101029143; uid=dw-wangcx9; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.105.4:all8380; BIGipServerywjk_new_pool1=42016172.10275.0000; Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1745809498,1745939912,1745993521; HMACCOUNT=4DF837C069A2D971; Hm_lpvt_f6097524da69abc1b63c9f8d19f5bd5b=1745993549; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzQ2MDA1ODc5LCJOQU5PU0VDT05EIjozNDg4MTgyOTE2NTg0MTU5Nn0.V2DZec2m2KP2QQN3OGdHt_jOHCgmfX0NB_ZACZRkw8U",
        "Host": "omms.chinatowercom.cn:9000",
        "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
    }
    data_list = [
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm:unitHidden": "0098364,0099977,0099978,0099979,0099980,0099981,0099982,0099983,0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449",
            "queryForm:queryFlag": "queryFlag",
            "queryForm:registstatusText_hiddenValue": "0",
            "queryForm:registstatusText": "0",
            "queryForm:queryStaStatusSelId_hiddenValue": "2",
            "queryForm:queryStaStatusSelId": "2",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "queryForm": "queryForm",
            "javax.faces.ViewState": "j_id38",
            "queryForm:j_id155": "queryForm:j_id155"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm:unitHidden": "0098364,0099977,0099978,0099979,0099980,0099981,0099982,0099983,0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449",
            "queryForm:queryFlag": "queryFlag",
            "queryForm:registstatusText_hiddenValue": "0",
            "queryForm:registstatusText": "0",
            "queryForm:queryStaStatusSelId_hiddenValue": "2",
            "queryForm:queryStaStatusSelId": "2",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "queryForm": "queryForm",
            "javax.faces.ViewState": "j_id38",
            "queryForm:j_id156": "queryForm:j_id156",
            "AJAX:EVENTS_COUNT": "1"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "j_id814": "j_id814",
            "javax.faces.ViewState": "j_id38",
            "j_id814:j_id817": "j_id814:j_id817"
        },

        {
            "j_id814": "j_id814",
            "j_id814:j_id816": "全部",
            "javax.faces.ViewState": "j_id38"
        }

    ]
    save_path = r"E:\陈桂志\普通文件"  # 保存路径
    file_name = "fsu离线.xlsx"

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
    config = Config
    url = config.url
    headers = config.headers
    data_list = config.data_list
    save_path = config.save_path
    file_name = config.file_name

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