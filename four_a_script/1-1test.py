import requests
import os
import json
from datetime import datetime
from bs4 import BeautifulSoup

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


class Config:
    url = "http://omms.chinatowercom.cn:9000/billDeal/monitoring/list/billList.xhtml"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": f"{cookie_header}",
        "Host": "omms.chinatowercom.cn:9000",
        "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
    }

    data_list2 = [
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:msg": "0",
            "queryForm:queryBillId": "",
            "queryForm:queryBillSn": "",
            "queryForm:isQueryHis": "N",
            "queryForm:queryStationId": "",
            "queryForm:deviceidText": "",
            "queryForm:addOrEditAreaNameId": "",
            "queryForm:aid": "",
            "queryForm:queryUnitId": "",
            "queryForm:j_id48": "",
            "queryForm:queryDWCompany": "",
            "queryForm:queryDWCompanyName": "",
            "queryForm:queryAlarmId": "",
            "queryForm:queryAlarmName": "",
            "queryForm:j_id58": "",
            "queryForm:starttimeInputDate": "2025-06-01 00:00",
            "queryForm:starttimeInputCurrentDate": "06/2025",
            "queryForm:starttimeTimeHours": "00",
            "queryForm:starttimeTimeMinutes": "00",
            "queryForm:endtimeInputDate": "2025-06-09 00:00",
            "queryForm:endtimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeInputDate": "2025-06-01 11:30",
            "queryForm:revertstarttimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeTimeHours": "00",
            "queryForm:revertstarttimeTimeMinutes": "00",
            "queryForm:revertendtimeInputDate": "2025-06-09 00:00",
            "queryForm:revertendtimeInputCurrentDate": "06/2025",
            "queryForm:revertendtimeTimeHours": "00",
            "queryForm:revertendtimeTimeMinutes": "00",
            "queryForm:dealstarttimeInputDate": "",
            "queryForm:dealstarttimeInputCurrentDate": "06/2025",
            "queryForm:dealendtimeInputDate": "",
            "queryForm:dealendtimeInputCurrentDate": "06/2025",
            "queryForm:sitesource_hiddenValue": "",
            "queryForm:querystationstatus_hiddenValue": "",
            "queryForm:billStatus_hiddenValue": "",
            "queryForm:faultSrc_hiddenValue": "移动运营商接口,联通运营商接口,电信运营商接口",
            "queryForm:faultSrc": "移动运营商接口",
            "queryForm:isHasten_hiddenValue": "",
            "queryForm:alarmlevel_hiddenValue": "",
            "queryForm:faultDevType_hiddenValue": "",
            "queryForm:isOverTime_hiddenValue": "",
            "queryForm:isReplyOver_hiddenValue": "",
            "queryForm:subOperatorHid_hiddenValue": "",
            "queryForm:operatorLevel_hiddenValue": "",
            "queryForm:turnSend_hiddenValue": "",
            "queryForm:sortSelect_hiddenValue": "",
            "queryForm:faultTypeId_hiddenValue": "",
            "queryForm:queryCrewProvinceId": "",
            "queryForm:queryCrewCityId": "",
            "queryForm:queryCrewAreaId": "",
            "queryForm:queryCrewVillageId": "",
            "queryForm:hideFlag": "",
            "queryForm:queryCrewVillageName": "",
            "queryForm:refreshTime": "",
            "queryForm:isTurnBack_hiddenValue": "",
            "queryForm:deleteproviceIdHidden": "",
            "queryForm:deletecityIdHidden": "",
            "queryForm:deletecountryIdHidden": "",
            "queryForm:queryDeleteCountyName": "",
            "queryForm:isTransitNodeId_hiddenValue": "",
            "queryForm:j_id139": "",
            "queryForm:j_id143": "",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id6",
            "queryForm:btn": "queryForm:btn"
        },
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:msg": "0",
            "queryForm:queryBillId": "",
            "queryForm:queryBillSn": "",
            "queryForm:isQueryHis": "N",
            "queryForm:queryStationId": "",
            "queryForm:deviceidText": "",
            "queryForm:addOrEditAreaNameId": "",
            "queryForm:aid": "",
            "queryForm:queryUnitId": "",
            "queryForm:j_id48": "",
            "queryForm:queryDWCompany": "",
            "queryForm:queryDWCompanyName": "",
            "queryForm:queryAlarmId": "",
            "queryForm:queryAlarmName": "",
            "queryForm:j_id58": "",
            "queryForm:starttimeInputDate": "2025-06-01 00:00",
            "queryForm:starttimeInputCurrentDate": "06/2025",
            "queryForm:starttimeTimeHours": "00",
            "queryForm:starttimeTimeMinutes": "00",
            "queryForm:endtimeInputDate": "2025-06-09 00:00",
            "queryForm:endtimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeInputDate": "2025-06-01 11:30",
            "queryForm:revertstarttimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeTimeHours": "00",
            "queryForm:revertstarttimeTimeMinutes": "00",
            "queryForm:revertendtimeInputDate": "2025-06-09 00:00",
            "queryForm:revertendtimeInputCurrentDate": "06/2025",
            "queryForm:revertendtimeTimeHours": "00",
            "queryForm:revertendtimeTimeMinutes": "00",
            "queryForm:dealstarttimeInputDate": "",
            "queryForm:dealstarttimeInputCurrentDate": "06/2025",
            "queryForm:dealendtimeInputDate": "",
            "queryForm:dealendtimeInputCurrentDate": "06/2025",
            "queryForm:sitesource_hiddenValue": "",
            "queryForm:querystationstatus_hiddenValue": "",
            "queryForm:billStatus_hiddenValue": "",
            "queryForm:faultSrc_hiddenValue": "移动运营商接口,联通运营商接口,电信运营商接口",
            "queryForm:faultSrc": "移动运营商接口",
            "queryForm:isHasten_hiddenValue": "",
            "queryForm:alarmlevel_hiddenValue": "",
            "queryForm:faultDevType_hiddenValue": "",
            "queryForm:isOverTime_hiddenValue": "",
            "queryForm:isReplyOver_hiddenValue": "",
            "queryForm:subOperatorHid_hiddenValue": "",
            "queryForm:operatorLevel_hiddenValue": "",
            "queryForm:turnSend_hiddenValue": "",
            "queryForm:sortSelect_hiddenValue": "",
            "queryForm:faultTypeId_hiddenValue": "",
            "queryForm:queryCrewProvinceId": "",
            "queryForm:queryCrewCityId": "",
            "queryForm:queryCrewAreaId": "",
            "queryForm:queryCrewVillageId": "",
            "queryForm:hideFlag": "",
            "queryForm:queryCrewVillageName": "",
            "queryForm:refreshTime": "",
            "queryForm:isTurnBack_hiddenValue": "",
            "queryForm:deleteproviceIdHidden": "",
            "queryForm:deletecityIdHidden": "",
            "queryForm:deletecountryIdHidden": "",
            "queryForm:queryDeleteCountyName": "",
            "queryForm:isTransitNodeId_hiddenValue": "",
            "queryForm:j_id139": "",
            "queryForm:j_id143": "",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id6",
            "queryForm:j_id150": "queryForm:j_id150",
            "AJAX:EVENTS_COUNT": "1"
        },
            {
                "j_id1945": "j_id1945",
                "j_id1945:j_id1947": "N",
                "j_id1945:devExport": "全部",
                "javax.faces.ViewState": "j_id6"
            }
    ]
    data_list1 = [
        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:msg": "0",
            "queryForm:queryBillId": "",
            "queryForm:queryBillSn": "",
            "queryForm:isQueryHis": "W",
            "queryForm:queryStationId": "",
            "queryForm:deviceidText": "",
            "queryForm:addOrEditAreaNameId": "",
            "queryForm:aid": "",
            "queryForm:queryUnitId": "",
            "queryForm:j_id48": "",
            "queryForm:queryDWCompany": "",
            "queryForm:queryDWCompanyName": "",
            "queryForm:queryAlarmId": "",
            "queryForm:queryAlarmName": "",
            "queryForm:j_id58": "",
            "queryForm:starttimeInputDate": "2025-06-01 00:00",
            "queryForm:starttimeInputCurrentDate": "06/2025",
            "queryForm:starttimeTimeHours": "00",
            "queryForm:starttimeTimeMinutes": "00",
            "queryForm:endtimeInputDate": "2025-06-09 00:00",
            "queryForm:endtimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeInputDate": "2025-06-01 11:30",
            "queryForm:revertstarttimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeTimeHours": "00",
            "queryForm:revertstarttimeTimeMinutes": "00",
            "queryForm:revertendtimeInputDate": "2025-06-09 00:00",
            "queryForm:revertendtimeInputCurrentDate": "06/2025",
            "queryForm:revertendtimeTimeHours": "00",
            "queryForm:revertendtimeTimeMinutes": "00",
            "queryForm:dealstarttimeInputDate": "",
            "queryForm:dealstarttimeInputCurrentDate": "06/2025",
            "queryForm:dealendtimeInputDate": "",
            "queryForm:dealendtimeInputCurrentDate": "06/2025",
            "queryForm:sitesource_hiddenValue": "",
            "queryForm:querystationstatus_hiddenValue": "",
            "queryForm:billStatus_hiddenValue": "",
            "queryForm:faultSrc_hiddenValue": "移动运营商接口,联通运营商接口,电信运营商接口",
            "queryForm:faultSrc": "移动运营商接口",
            "queryForm:isHasten_hiddenValue": "",
            "queryForm:alarmlevel_hiddenValue": "",
            "queryForm:faultDevType_hiddenValue": "",
            "queryForm:isOverTime_hiddenValue": "",
            "queryForm:isReplyOver_hiddenValue": "",
            "queryForm:subOperatorHid_hiddenValue": "",
            "queryForm:operatorLevel_hiddenValue": "",
            "queryForm:turnSend_hiddenValue": "",
            "queryForm:sortSelect_hiddenValue": "",
            "queryForm:faultTypeId_hiddenValue": "",
            "queryForm:queryCrewProvinceId": "",
            "queryForm:queryCrewCityId": "",
            "queryForm:queryCrewAreaId": "",
            "queryForm:queryCrewVillageId": "",
            "queryForm:hideFlag": "",
            "queryForm:queryCrewVillageName": "",
            "queryForm:refreshTime": "",
            "queryForm:isTurnBack_hiddenValue": "",
            "queryForm:deleteproviceIdHidden": "",
            "queryForm:deletecityIdHidden": "",
            "queryForm:deletecountryIdHidden": "",
            "queryForm:queryDeleteCountyName": "",
            "queryForm:isTransitNodeId_hiddenValue": "",
            "queryForm:j_id139": "",
            "queryForm:j_id143": "",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id6",
            "queryForm:btn": "queryForm:btn"
        },

        {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:msg": "0",
            "queryForm:queryBillId": "",
            "queryForm:queryBillSn": "",
            "queryForm:isQueryHis": "W",
            "queryForm:queryStationId": "",
            "queryForm:deviceidText": "",
            "queryForm:addOrEditAreaNameId": "",
            "queryForm:aid": "",
            "queryForm:queryUnitId": "",
            "queryForm:j_id48": "",
            "queryForm:queryDWCompany": "",
            "queryForm:queryDWCompanyName": "",
            "queryForm:queryAlarmId": "",
            "queryForm:queryAlarmName": "",
            "queryForm:j_id58": "",
            "queryForm:starttimeInputDate": "2025-06-01 00:00",
            "queryForm:starttimeInputCurrentDate": "06/2025",
            "queryForm:starttimeTimeHours": "00",
            "queryForm:starttimeTimeMinutes": "00",
            "queryForm:endtimeInputDate": "2025-06-09 00:00",
            "queryForm:endtimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeInputDate": "2025-06-01 11:30",
            "queryForm:revertstarttimeInputCurrentDate": "06/2025",
            "queryForm:revertstarttimeTimeHours": "00",
            "queryForm:revertstarttimeTimeMinutes": "00",
            "queryForm:revertendtimeInputDate": "2025-06-09 00:00",
            "queryForm:revertendtimeInputCurrentDate": "06/2025",
            "queryForm:revertendtimeTimeHours": "00",
            "queryForm:revertendtimeTimeMinutes": "00",
            "queryForm:dealstarttimeInputDate": "",
            "queryForm:dealstarttimeInputCurrentDate": "06/2025",
            "queryForm:dealendtimeInputDate": "",
            "queryForm:dealendtimeInputCurrentDate": "06/2025",
            "queryForm:sitesource_hiddenValue": "",
            "queryForm:querystationstatus_hiddenValue": "",
            "queryForm:billStatus_hiddenValue": "",
            "queryForm:faultSrc_hiddenValue": "移动运营商接口,联通运营商接口,电信运营商接口",
            "queryForm:faultSrc": "移动运营商接口",
            "queryForm:isHasten_hiddenValue": "",
            "queryForm:alarmlevel_hiddenValue": "",
            "queryForm:faultDevType_hiddenValue": "",
            "queryForm:isOverTime_hiddenValue": "",
            "queryForm:isReplyOver_hiddenValue": "",
            "queryForm:subOperatorHid_hiddenValue": "",
            "queryForm:operatorLevel_hiddenValue": "",
            "queryForm:turnSend_hiddenValue": "",
            "queryForm:sortSelect_hiddenValue": "",
            "queryForm:faultTypeId_hiddenValue": "",
            "queryForm:queryCrewProvinceId": "",
            "queryForm:queryCrewCityId": "",
            "queryForm:queryCrewAreaId": "",
            "queryForm:queryCrewVillageId": "",
            "queryForm:hideFlag": "",
            "queryForm:queryCrewVillageName": "",
            "queryForm:refreshTime": "",
            "queryForm:isTurnBack_hiddenValue": "",
            "queryForm:deleteproviceIdHidden": "",
            "queryForm:deletecityIdHidden": "",
            "queryForm:deletecountryIdHidden": "",
            "queryForm:queryDeleteCountyName": "",
            "queryForm:isTransitNodeId_hiddenValue": "",
            "queryForm:j_id139": "",
            "queryForm:j_id143": "",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id6",
            "queryForm:j_id150": "queryForm:j_id150",
            "AJAX:EVENTS_COUNT": "1"
        },
        {
            "j_id1945": "j_id1945",
            "j_id1945:j_id1947": "N",
            "j_id1945:devExport": "全部",
            "javax.faces.ViewState": "j_id6"
        }
    ]
    INDEX = os.getcwd()
    save_path = os.path.join(INDEX, "output")  # 修改为当前相对路径下的output文件夹
    file_name1 = os.path.join(save_path, "当前工单.xlsx")  # 修改为当前工单.xlsx
    file_name2 = os.path.join(save_path, "当天前历史工单.xlsx")  # 修改为当天前历史工单.xlsx

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

# 通用处理数据列表的函数
def process_data_list(config, data_list, file_name):
    url = config.url
    headers = config.headers

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
            save_response_to_file(response, file_name)

# 主程序模块
def main():
    config = Config
    process_data_list(config, config.data_list1, config.file_name1)
    process_data_list(config, config.data_list2, config.file_name2)

if __name__ == "__main__":
    main()