import requests
import os
import pandas as pd
from bs4 import BeautifulSoup


# 配置模块
class Config:
    url = "http://omms.chinatowercom.cn:9000/business/resMge/siteMge/listSite.xhtml"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": "route=7634b5a86f92649382e66b47cf771393; ULTRA_U_K=; JSESSIONID=E612AD76AACBF01B21DE44713E66F5EF; acctId=101029143; uid=dw-wangcx9; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.105.4:all8380; BIGipServerywjk_new_pool1=42016172.10275.0000; Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1745809498,1745939912,1745993521; HMACCOUNT=4DF837C069A2D971; Hm_lpvt_f6097524da69abc1b63c9f8d19f5bd5b=1745993549; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzQ1OTk3OTg1LCJOQU5PU0VDT05EIjozNDg3NTEwMzUzMTE3NDA1MX0.Y7JBSrHs-m_28cu8wNyUzIEfvHowBeQeiJXUNlg3oBM",
        "Host": "omms.chinatowercom.cn:9000",
        "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
    }
    data_list1 = [
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:setvisualization": "http://180.153.49.102:18582/?page=preview&sceneId=",
            "queryForm:unitHidden": "0099977,0099978,0099979,0099980,0099981,0099982,0099983",
            "queryForm:queryFlag1": "queryFlag1",
            "queryForm:querystatusSelid_hiddenValue": "2",
            "queryForm:querystatusSelid": "2",
            "queryForm:isUnion": "N",
            "queryForm:nodeIsUnion": "N",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id18",
            "queryForm:siteQueryId": "queryForm:siteQueryId"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:setvisualization": "http://180.153.49.102:18582/?page=preview&sceneId=",
            "queryForm:unitHidden": "0099977,0099978,0099979,0099980,0099981,0099982,0099983",
            "queryForm:queryFlag1": "queryFlag1",
            "queryForm:querystatusSelid_hiddenValue": "2",
            "queryForm:querystatusSelid": "2",
            "queryForm:isUnion": "N",
            "queryForm:nodeIsUnion": "N",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "0",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id18",
            "queryForm:j_id195": "queryForm:j_id195",
            "AJAX:EVENTS_COUNT": "1"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:setvisualization": "http://180.153.49.102:18582/?page=preview&sceneId=",
            "queryForm:unitHidden": "0099977,0099978,0099979,0099980,0099981,0099982,0099983",
            "queryForm:queryFlag1": "queryFlag1",
            "queryForm:querystatusSelid_hiddenValue": "2",
            "queryForm:querystatusSelid": "2",
            "queryForm:isUnion": "N",
            "queryForm:nodeIsUnion": "N",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "1",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id18",
            "queryForm:j_id189": "queryForm:j_id189"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:setvisualization": "http://180.153.49.102:18582/?page=preview&sceneId=",
            "queryForm:unitHidden": "0099977,0099978,0099979,0099980,0099981,0099982,0099983",
            "queryForm:queryFlag1": "queryFlag1",
            "queryForm:querystatusSelid_hiddenValue": "2",
            "queryForm:querystatusSelid": "2",
            "queryForm:isUnion": "N",
            "queryForm:nodeIsUnion": "N",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "1",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id18",
            "queryForm:j_id190": "queryForm:j_id190",
            "AJAX:EVENTS_COUNT": "1"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "j_id1517": "j_id1517",
            "j_id1517:exportProperty": "0",
            "j_id1517:j_id1528": "2025",
            "javax.faces.ViewState": "j_id18",
            "j_id1517:j_id1536": "j_id1517:j_id1536"
        },

        {
            "j_id1517": "j_id1517",
            "j_id1517:exportProperty": "0",
            "j_id1517:j_id1528": "2025",
            "j_id1517:j_id1535": "全部",
            "javax.faces.ViewState": "j_id18"
        }
    ]
    data_list2 =[
        {
            "AJAXREQUEST": [
                "_viewRoot",
                "_viewRoot"
            ],
            "queryForm": [
                "queryForm",
                "queryForm"
            ],
            "queryForm:setvisualization": [
                "http://180.153.49.102:18582/?page=preview&sceneId=",
                "http://180.153.49.102:18582/?page=preview&sceneId="
            ],
            "queryForm:unitHidden": [
                "0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449",
                "0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449"
            ],
            "queryForm:queryFlag1": [
                "queryFlag1",
                "queryFlag1"
            ],
            "queryForm:querystatusSelid_hiddenValue": [
                "2",
                "2"
            ],
            "queryForm:querystatusSelid": [
                "2",
                "2"
            ],
            "queryForm:isUnion": [
                "N",
                "N"
            ],
            "queryForm:nodeIsUnion": [
                "N",
                "N"
            ],
            "queryForm:msg": [
                "0",
                "0"
            ],
            "queryForm:currPageObjId": [
                "1",
                "1"
            ],
            "queryForm:pageSizeText": [
                "35",
                "35"
            ],
            "javax.faces.ViewState": [
                "j_id18",
                "j_id18"
            ],
            "queryForm:siteQueryId": "queryForm:siteQueryId",
            "queryForm:j_id195": "queryForm:j_id195",
            "AJAX:EVENTS_COUNT": "1"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:setvisualization": "http://180.153.49.102:18582/?page=preview&sceneId=",
            "queryForm:unitHidden": "0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449",
            "queryForm:queryFlag1": "queryFlag1",
            "queryForm:querystatusSelid_hiddenValue": "2",
            "queryForm:querystatusSelid": "2",
            "queryForm:isUnion": "N",
            "queryForm:nodeIsUnion": "N",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "1",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id18",
            "queryForm:j_id189": "queryForm:j_id189"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:setvisualization": "http://180.153.49.102:18582/?page=preview&sceneId=",
            "queryForm:unitHidden": "0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449",
            "queryForm:queryFlag1": "queryFlag1",
            "queryForm:querystatusSelid_hiddenValue": "2",
            "queryForm:querystatusSelid": "2",
            "queryForm:isUnion": "N",
            "queryForm:nodeIsUnion": "N",
            "queryForm:msg": "0",
            "queryForm:currPageObjId": "1",
            "queryForm:pageSizeText": "35",
            "javax.faces.ViewState": "j_id18",
            "queryForm:j_id190": "queryForm:j_id190",
            "AJAX:EVENTS_COUNT": "1"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "j_id1517": "j_id1517",
            "j_id1517:exportProperty": "0",
            "j_id1517:j_id1528": "2025",
            "javax.faces.ViewState": "j_id18",
            "j_id1517:j_id1536": "j_id1517:j_id1536"
        },

        {
            "j_id1517": "j_id1517",
            "j_id1517:exportProperty": "0",
            "j_id1517:j_id1528": "2025",
            "j_id1517:j_id1535": "全部",
            "javax.faces.ViewState": "j_id18"
        }

    ]
    save_path = r"E:\陈桂志\普通文件"  # 保存路径
    file_name1 = "notice1.xlsx"
    file_name2 = "notice2.xlsx"
    final_file_name = "final_notice.xlsx"  # 合并后的文件名

# 合并文件模块
def merge_files(file_path1, file_path2, final_file_path):
    """合并两个Excel文件"""
    df1 = pd.read_excel(file_path1)
    df2 = pd.read_excel(file_path2)
    merged_df = pd.concat([df1, df2], ignore_index=True)
    merged_df.to_excel(final_file_path, index=False)
    print(f"文件已成功合并到: {final_file_path}")

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
    data_list1 = config.data_list1   #第一个数据列表，包含前7个地级市
    data_list2 = config.data_list2   #第二个数据列表，包含后7个地级市
    save_path = config.save_path
    file_name1 = config.file_name1
    file_name2 = config.file_name2
    final_file_name = config.final_file_name     #合并两个响应文件


    # 获取javax.faces.ViewState值
    view_state = get_view_state(url, headers)
    if not view_state:
        print("无法获取javax.faces.ViewState值，请检查URL和headers配置")
        return

    # 使用第一个数据列表发送请求并保存响应文件
    for data in data_list1:
        data["javax.faces.ViewState"] = view_state
        response = post_data(url, headers, data)
    save_response_to_file(response, save_path, file_name1)

    # 使用第二个数据列表发送请求并保存响应文件
    for data in data_list2:
        data["javax.faces.ViewState"] = view_state
        response = post_data(url, headers, data)
    save_response_to_file(response, save_path, file_name2)

    # 合并两个文件
    file_path1 = os.path.join(save_path, file_name1)
    file_path2 = os.path.join(save_path, file_name2)
    final_file_path = os.path.join(save_path, final_file_name)
    merge_files(file_path1, file_path2, final_file_path)


if __name__ == "__main__":
    main()