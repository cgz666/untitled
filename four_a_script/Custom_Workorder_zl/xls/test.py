import requests
import os
import json
import pythoncom
from bs4 import BeautifulSoup
import time
import win32com.client as win32
from datetime import datetime

class interface_result():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.url = "http://omms.chinatowercom.cn:9000/billDeal/monitoring/list/billList.xhtml"
        self.headers = {
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
        self.data_list = [
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
                "queryForm:queryUnitId": "0098364,0099977,0099978,0099979,0099980,0099981,0099982,0099983,0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449",
                "queryForm:j_id48": "",
                "queryForm:queryDWCompany": "",
                "queryForm:queryDWCompanyName": "",
                "queryForm:queryAlarmId": "",
                "queryForm:queryAlarmName": "",
                "queryForm:j_id58": "",
                "queryForm:starttimeInputDate": "2025-06-26 00:00",
                "queryForm:starttimeInputCurrentDate": "06/2025",
                "queryForm:starttimeTimeHours": "00",
                "queryForm:starttimeTimeMinutes": "00",
                "queryForm:endtimeInputDate": "2025-06-26 23:59",
                "queryForm:endtimeInputCurrentDate": "06/2025",
                "queryForm:endtimeTimeHours": "23",
                "queryForm:endtimeTimeMinutes": "59",
                "queryForm:revertstarttimeInputDate": "2025-06-26 00:00",
                "queryForm:revertstarttimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeTimeHours": "00",
                "queryForm:revertstarttimeTimeMinutes": "00",
                "queryForm:revertendtimeInputDate": "2025-06-26 23:59",
                "queryForm:revertendtimeInputCurrentDate": "06/2025",
                "queryForm:revertendtimeTimeHours": "23",
                "queryForm:revertendtimeTimeMinutes": "59",
                "queryForm:dealstarttimeInputDate": "",
                "queryForm:dealstarttimeInputCurrentDate": "06/2025",
                "queryForm:dealendtimeInputDate": "",
                "queryForm:dealendtimeInputCurrentDate": "06/2025",
                "queryForm:sitesource_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "",
                "queryForm:billStatus_hiddenValue": "",
                "queryForm:faultSrc_hiddenValue": "铁塔集团动环网管,移动运营商接口,联通运营商接口,电信运营商接口",
                "queryForm:faultSrc": "铁塔集团动环网管",
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
                "queryForm:isTransitNode显示更多": ""
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
                "queryForm:queryUnitId": "0098364,0099977,0099978,0099979,0099980,0099981,0099982,0099983,0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449",
                "queryForm:j_id48": "",
                "queryForm:queryDWCompany": "",
                "queryForm:queryDWCompanyName": "",
                "queryForm:queryAlarmId": "",
                "queryForm:queryAlarmName": "",
                "queryForm:j_id58": "",
                "queryForm:starttimeInputDate": "2025-06-26 00:00",
                "queryForm:starttimeInputCurrentDate": "06/2025",
                "queryForm:starttimeTimeHours": "00",
                "queryForm:starttimeTimeMinutes": "00",
                "queryForm:endtimeInputDate": "2025-06-26 23:59",
                "queryForm:endtimeInputCurrentDate": "06/2025",
                "queryForm:endtimeTimeHours": "23",
                "queryForm:endtimeTimeMinutes": "59",
                "queryForm:revertstarttimeInputDate": "2025-06-26 00:00",
                "queryForm:revertstarttimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeTimeHours": "00",
                "queryForm:revertstarttimeTimeMinutes": "00",
                "queryForm:revertendtimeInputDate": "2025-06-26 23:59",
                "queryForm:revertendtimeInputCurrentDate": "06/2025",
                "queryForm:revertendtimeTimeHours": "23",
                "queryForm:revertendtimeTimeMinutes": "59",
                "queryForm:dealstarttimeInputDate": "",
                "queryForm:dealstarttimeInputCurrentDate": "06/2025",
                "queryForm:dealendtimeInputDate": "",
                "queryForm:dealendtimeInputCurrentDate": "06/2025",
                "queryForm:sitesource_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "",
                "queryForm:billStatus_hiddenValue": "",
                "queryForm:faultSrc_hiddenValue": "铁塔集团动环网管,移动运营商接口,联通运营商接口,电信运营商接口",
                "queryForm:faultSrc": "铁塔集团动环网管",
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
                "queryForm:isTransitNode显示更多": ""
        },
            {
                "j_id1945": "j_id1945",
                "j_id1945:j_id1947": "N",
                "j_id1945:devExport": "全部",
                "javax.faces.ViewState": "j_id6"
            }

        ]
        self.city_list = [
            "0099977,0108557,0108559,0108560,0108562,0108563,0108564,0108565,0108566,2735327428",
            "0099978,0108569,0108570,0108571,0108572,0108573,0108574,0108575,0108576,0108577,0108578,2788568264,2788570112,2788568312,2788567212,2788568343",
            "0099979",
            "0099980",
            "0099981",
            "0099982",
            "0099983",
            "0099984",
            "0099985",
            "0099986",
            "0099987",
            "0099988",
            "0099989",
            "0099990,0108648,0108649,0108650,0108651",
            "2710377449,2710377453"
        ]

        INDEX = r"F:\newtower3.8\project\four_a_script\interface_result"
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "爬取运营商接口工单-结果11.xlsx")
        self.file_name1 = os.path.join(self.save_path, "当前工单11.xls")
        self.file_name2 = os.path.join(self.save_path, "当天前历史工单11.xls")
        self.model_path =  os.path.join(self.save_path, "运营商接口工单完成情况统计表11.xlsx")
        self.temp_files = [os.path.join(self.save_path, f"temp_{i}.xls") for i in range(len(self.city_list))]

        self.current_date = datetime.now()
        self.first_day_of_month = self.current_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        self.first_day_of_month_str = self.first_day_of_month.strftime("%Y-%m-%d %H:%M")
        self.current_day = self.current_date.replace(hour=0, minute=0)
        self.current_day_str = self.current_day.strftime("%Y-%m-%d %H:%M")
        self.current_month_str = self.current_date.strftime("%m/%Y")
    def get_view_state(self,url, headers):
        """获取页面的javax.faces.ViewState值"""
        res = requests.get(url=url, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        view_state_input = soup.find('input', id='javax.faces.ViewState')
        if view_state_input:
            return view_state_input.get('value')
        return None

    # 通用处理数据列表的函数
    def process_data_list(self):
        url = self.url
        headers = self.headers
        view_state = self.get_view_state(url, headers)


        if not view_state:
            print("无法获取javax.faces.ViewState值，请检查URL和headers配置")
            return
        for i, data in enumerate(self.data_list, start=1):
            if i in [1, 2]:  # 处理前两个数据项

                data["javax.faces.ViewState"] = view_state
                response = requests.post(url=url, data=data, headers=headers)
                if i == len(self.data_list):
                    with open(self.file_name1, "wb") as file:
                        file.write(response.content)
                    print("当前工单已保存到:", self.file_name1)




    def main(self):
        self.process_data_list()

if __name__ == "__main__":
    interface_result().main()