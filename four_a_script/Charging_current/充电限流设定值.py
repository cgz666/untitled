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

        self.url = "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/realTimePerformanceMge/realTimeperfdata.xhtml"
        self.headers = {
            "Accept": "*/*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Origin": "http://omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/realTimePerformanceMge/realTimeperfdata.xhtml",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }
        self.data_list = [
            {
                "AJAXREQUEST": "_viewRoot",
                "stationListForm:nameText": "南宁马山县里当乡棋盘基站无线机房",
                "stationListForm:stationidText": "",
                "stationListForm:queryStatusId": "",
                "stationListForm:stationcode": "",
                "stationListForm:currPageObjId": "0",
                "stationListForm": "stationListForm",
                "autoScroll": "",
                "javax.faces.ViewState": "j_id6",
                "stationListForm:j_id284": "stationListForm:j_id284",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm2": "queryForm2",
                "queryForm2:j_id123": "",
                "queryForm2:aid": "",
                "queryForm2:panel2OpenedState": "",
                "javax.faces.ViewState": "j_id6",
                "queryForm2:j_id125": "queryForm2:j_id125",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:addOrEditAreaNameId": "",
                "queryForm:aid": "",
                "queryForm:fsuid": "",
                "queryForm:deviceName": "",
                "queryForm:did": "",
                "queryForm:midName": "电池充电限流设定值",
                "queryForm:mid": "0406301001",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "j_id6",
                "queryForm:j_id21": "queryForm:j_id21"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:addOrEditAreaNameId": "",
                "queryForm:aid": "",
                "queryForm:fsuid": "",
                "queryForm:deviceName": "",
                "queryForm:did": "",
                "queryForm:midName": "电池充电限流设定值",
                "queryForm:mid": "0406301001",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "j_id6",
                "queryForm:j_id22": "queryForm:j_id22",
                "AJAX:EVENTS_COUNT": "1"
            }
        ]

        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "爬取运营商接口工单-结果.xlsx")
        self.file_name = os.path.join(self.save_path, "充电电流设定值.xlsx")

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
        select_flag_input = soup1.find('input', attrs={'name': 'selectFlag'})
        if select_flag_input:
            aid_value = select_flag_input.get('id')
            print(f"获取到的selectFlag的id值: {aid_value}")
        else:
            print("无法获取selectFlag的id值")
            return

        # 将获取到的id值填入第二个data的queryForm2:aid
        data2 = self.data_list[1]
        data2["queryForm2:aid"] = aid_value
        data2["javax.faces.ViewState"] = view_state
        response2 = requests.post(url=self.url, data=data2, headers=self.headers)
        soup2 = BeautifulSoup(response2.text, 'html.parser')

        # 从第二个请求的响应中提取包含"开关电源01"的id值
        did_input = soup2.find('input', attrs={'value': '南宁马山县里当乡棋盘基站无线机房/开关电源01'})
        if did_input:
            did_value = did_input.get('id')
            print(f"获取到的包含'开关电源01'的id值: {did_value}")
        else:
            print("无法获取包含'开关电源01'的id值")
            return

        # 将获取到的aid和did值填入第三个和第四个data
        data3 = self.data_list[2]
        data3["queryForm:aid"] = aid_value
        data3["queryForm:did"] = did_value
        data3["javax.faces.ViewState"] = view_state
        response3 = requests.post(url=self.url, data=data3, headers=self.headers)
        soup3 = BeautifulSoup(response3.text, 'html.parser')

        data4 = self.data_list[3]
        data4["queryForm:aid"] = aid_value
        data4["queryForm:did"] = did_value
        data4["javax.faces.ViewState"] = view_state
        response4 = requests.post(url=self.url, data=data4, headers=self.headers)
        soup4 = BeautifulSoup(response4.text, 'html.parser')

        # 从最后一个请求的响应中提取数据并保存到Excel文件
        self.extract_and_save_data(soup4)

    def extract_and_save_data(self, soup):
        """从HTML中提取数据并保存到Excel文件"""
        table = soup.find('table', class_='rich-table')
        if not table:
            print("无法找到数据表格")
            return

        rows = table.find_all('tr', class_='rich-table-row')
        data = []

        for row in rows:
            cells = row.find_all('td', class_='rich-table-cell')
            if len(cells) < 11:
                continue

            # 提取数据
            station_maintenance_id = cells[2].text.strip()
            station_resource_code = cells[3].text.strip()
            device = cells[4].text.strip()
            signal_id = cells[5].text.strip()
            monitoring_point = cells[6].text.strip()
            measured_value = cells[7].text.strip()
            unit = cells[8].text.strip()
            status = cells[9].text.strip()

            data.append({
                "站址": "",
                "站址名备注": "",
                "站址运维ID": station_maintenance_id,
                "站址资源编码": station_resource_code,
                "设备": device,
                "信号量ID": signal_id,
                "监控点": monitoring_point,
                "实测值": measured_value,
                "单位": unit,
                "状态": status
            })

        # 保存数据到Excel文件
        if data:
            df = pd.DataFrame(data)
            df.to_excel(self.file_name, index=False)
            print(f"数据已保存到文件: {self.file_name}")
        else:
            print("没有数据可保存")

    def main(self):
        self.process_data_list()

if __name__ == "__main__":
    Charging_current().main()