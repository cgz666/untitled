import requests
import os
from bs4 import BeautifulSoup
import time

# 配置模块
class Config:
    url = "http://omms.chinatowercom.cn:9000/business/resMge/alarmHisHbaseMge/listHisAlarmHbase.xhtml"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": "Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1745809498,1745939912,1745993521,1746429292; route=01905252c4af4acb5aa8a1354115d784; ULTRA_U_K=; JSESSIONID=CDDB2DB6D2C8A709EAF741EB6DD72D37; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzQ2NjEyMjE1LCJOQU5PU0VDT05EIjozNzM2MjA2NTgxMjIwMjY5fQ.jLy_vLhhUnml0r3cZoN03SXT-Nb0kwyY6drRV9QkkEk; acctId=101429653; uid=wx-huangwl14; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.129.16:all8280; BIGipServerywjk_new_pool1=342433196.10275.0000",
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
          "queryForm:firstendtimeInputDate": "2025-05-07 00:00",
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
          "javax.faces.ViewState": "j_id6",
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
          "queryForm:firstendtimeInputDate": "2025-05-07 00:00",
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
          "javax.faces.ViewState": "j_id6",
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
          "queryForm:firstendtimeInputDate": "2025-05-07 00:00",
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
          "javax.faces.ViewState": "j_id6",
          "queryForm:j_id60": "queryForm:j_id60"
        },

        {
              "j_id313": "j_id313",
              "j_id313:devExport": "全部",
              "javax.faces.ViewState": "j_id6"
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
    file_path = os.path.join(save_path, file_name)
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

        # 添加延迟，避免请求过于频繁
        time.sleep(1)

if __name__ == "__main__":
    main()