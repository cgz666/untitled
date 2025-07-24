import requests
import os
import json
from bs4 import BeautifulSoup
import pandas as pd

class Charging_current():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.url = "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml"
        self.headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Cache-Control": "max-age=0",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Origin": "http://omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }
        self.data_list = [
            {
                "AJAXREQUEST": "_viewRoot",
                "stationListFormB:nameText": "江州区理工学校12号公寓楼",
                "stationListFormB:stationidText": "",
                "stationListFormB:queryStatusId": "2",
                "stationListFormB:currPageObjId": "0",
                "stationListFormB": "stationListFormB",
                "autoScroll": "",
                "javax.faces.ViewState": "j_id22",
                "stationListFormB:j_id764": "stationListFormB:j_id764",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm2B": "queryForm2B",
                "queryForm2B:aname": "",
                "queryForm2B:aid": "5C244C87B3D4F279E0530D801DACF07C",
                "queryForm2B:panel2OpenedState": "",
                "javax.faces.ViewState": "j_id22",
                "queryForm2B:j_id720": "queryForm2B:j_id720",
                "AJAX:EVENTS_COUNT": "1"
            },

            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm3B": "queryForm3B",
                "queryForm3B:ifRestrict": "true",
                "queryForm3B:mname": "",
                "queryForm3B:did": "45140240600740",
                "queryForm3B:panel3OpenedState": "",
                "javax.faces.ViewState": "j_id22",
                "queryForm3B:j_id543": "queryForm3B:j_id543",
                "AJAX:EVENTS_COUNT": "1"
            },

            {
                "AJAXREQUEST": "_viewRoot",
                "queryFormB": "queryFormB",
                "queryFormB:aid": "5C244C87B3D4F279E0530D801DACF07C",
                "queryFormB:siteProvinceId": "0098364",
                "queryFormB:queryFlag": "3",
                "queryFormB:unitHidden1": "",
                "queryFormB:unitTypeHidden": "",
                "queryFormB:siteNameId": "江州区理工学校12号公寓楼",
                "queryFormB:deviceName": "江州区理工学校12号公寓楼/开关电源01",
                "queryFormB:did": "45140240600740",
                "queryFormB:midName": "直流电压",
                "queryFormB:mid": "0406111001",
                "queryFormB:queryFsuId": "",
                "queryFormB:midType": "遥测",
                "queryFormB:starttimeInputDate": "2025-07-21 11:17",
                "queryFormB:starttimeInputCurrentDate": "07/2025",
                "queryFormB:endtimeInputDate": "2025-07-22 11:17",
                "queryFormB:endtimeInputCurrentDate": "07/2025",
                "queryFormB:querySiteSourceCode": "",
                "queryFormB:ifRestrict": "true",
                "queryFormB:currPageObjId": "0",
                "queryFormB:pageSizeText": "35",
                "queryFormB:panelOpenedState": "",
                "javax.faces.ViewState": "j_id22",
                "queryFormB:j_id186": "queryFormB:j_id186"
            },

            {
                "AJAXREQUEST": "_viewRoot",
                "queryFormB": "queryFormB",
                "queryFormB:aid": "5C244C87B3D4F279E0530D801DACF07C",
                "queryFormB:siteProvinceId": "0098364",
                "queryFormB:queryFlag": "3",
                "queryFormB:unitHidden1": "",
                "queryFormB:unitTypeHidden": "",
                "queryFormB:siteNameId": "江州区理工学校12号公寓楼",
                "queryFormB:deviceName": "江州区理工学校12号公寓楼/开关电源01",
                "queryFormB:did": "45140240600740",
                "queryFormB:midName": "直流电压",
                "queryFormB:mid": "0406111001",
                "queryFormB:queryFsuId": "",
                "queryFormB:midType": "遥测",
                "queryFormB:starttimeInputDate": "2025-07-21 11:17",
                "queryFormB:starttimeInputCurrentDate": "07/2025",
                "queryFormB:endtimeInputDate": "2025-07-22 11:17",
                "queryFormB:endtimeInputCurrentDate": "07/2025",
                "queryFormB:querySiteSourceCode": "",
                "queryFormB:ifRestrict": "true",
                "queryFormB:currPageObjId": "0",
                "queryFormB:pageSizeText": "35",
                "queryFormB:panelOpenedState": "",
                "javax.faces.ViewState": "j_id22",
                "queryFormB:j_id187": "queryFormB:j_id187",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "j_id432": "j_id432",
                "j_id432:j_id434": "全部",
                "javax.faces.ViewState": "j_id22"
            }
        ]

        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "爬取运营商接口工单-结果.xlsx")
        self.file_name = os.path.join(self.save_path, "三相电流.xlsx")

    def get_view_state(self):
        """获取页面的javax.faces.ViewState值"""
        res = requests.post(url=self.url, headers=self.headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        view_state_input = soup.find('input', id='javax.faces.ViewState')
        if view_state_input:
            return view_state_input.get('value')
        return None

    def process_data_list(self):
        view_state = self.get_view_state()
        if not view_state:
            print("无法获取javax.faces.ViewState值，请检查URL和headers配置")
            return

        # 第一次请求
        data1 = self.data_list[0]
        data1["javax.faces.ViewState"] = view_state
        response1 = requests.post(url=self.url, data=data1, headers=self.headers)
        soup1 = BeautifulSoup(response1.text, 'html.parser')

        # 从第一次请求的响应中提取selectFlag的id值
        select_flag_input = soup1.find('input', attrs={'name': 'selectFlagB'})
        if select_flag_input:
            aid_value = select_flag_input.get('id')
            print(f"获取到的selectFlag的id值: {aid_value}")
        else:
            print("无法获取selectFlag的id值")
            return

        # 将获取到的id值填入第二个data的queryForm2B:aid
        data2 = self.data_list[1]
        data2["queryForm2B:aid"] = aid_value
        data2["javax.faces.ViewState"] = view_state
        response2 = requests.post(url=self.url, data=data2, headers=self.headers)
        soup2 = BeautifulSoup(response2.text, 'html.parser')

        # 从第二个请求的响应中提取包含"开关电源01"的id值
        did_input = soup2.find('input', attrs={'value': '江州区理工学校12号公寓楼/开关电源01'})
        if did_input:
            did_value = did_input.get('id')
            print(f"获取到的包含'开关电源01'的id值: {did_value}")
        else:
            print("无法获取包含'开关电源01'的id值")
            return

        # 处理后续请求
        for i in range(2, len(self.data_list)):
            data = self.data_list[i].copy()  # 创建副本避免修改原始数据

            # 更新aid和did值
            if "queryFormB" in data:  # 对于queryFormB表单
                data["queryFormB:aid"] = aid_value
                data["queryFormB:did"] = did_value
            elif "queryForm3B" in data:  # 对于queryForm3B表单
                data["queryForm3B:did"] = did_value

            data["javax.faces.ViewState"] = view_state

            # 发送请求
            response = requests.post(url=self.url, data=data, headers=self.headers)

            # 如果是最后一个请求，保存响应内容
            if i == len(self.data_list) - 1:
                with open(self.file_name, 'wb') as f:
                    f.write(response.content)
                print(f"响应内容已保存到: {self.file_name}")

                # 检查文件是否有效
                if os.path.getsize(self.file_name) < 1024:  # 小于1KB认为可能无效
                    print("警告: 下载的文件可能无效，文件大小过小")
            else:
                # 更新view_state为最新响应中的值
                soup = BeautifulSoup(response.text, 'html.parser')
                new_view_state = soup.find('input', id='javax.faces.ViewState')
                if new_view_state:
                    view_state = new_view_state.get('value')

    def main(self):
        self.process_data_list()

if __name__ == "__main__":
    Charging_current().main()