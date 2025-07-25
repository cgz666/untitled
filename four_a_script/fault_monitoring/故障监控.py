import requests
import json
import time
from datetime import datetime
import pandas as pd
import os
from bs4 import BeautifulSoup

class fault_monitoring():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.url = "http://omms.chinatowercom.cn:9000/business/resMge/faultAlarmMge/listFaultActive.xhtml"
        self.headers = {
          "Accept": "*/*",
          "Accept-Encoding": "gzip, deflate",
          "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
          "Connection": "keep-alive",
          "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
          "Cookie": f"{cookie_header}",
          "Host": "omms.chinatowercom.cn:9000",
          "Origin": "http://omms.chinatowercom.cn:9000",
          "Referer": "http://omms.chinatowercom.cn:9000/business/resMge/faultAlarmMge/listFaultActive.xhtml",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
}
        self.data_list = [
            {
                "AJAXREQUEST": "_viewRoot",
                "hisQueryForm": "hisQueryForm",
                "hisQueryForm:unitHidden": "",
                "hisQueryForm:unitHid": "",
                "hisQueryForm:queryDay": "30",
                "hisQueryForm:queryFaultMids_hiddenValue": "退服场景",
                "hisQueryForm:queryFaultMids": "退服场景",
                "hisQueryForm:queryFaultDetail": "",
                "hisQueryForm:queryFaultDetailName": "",
                "hisQueryForm:queryLevel_hiddenValue": "",
                "hisQueryForm:j_id201": "",
                "hisQueryForm:j_id205": "",
                "hisQueryForm:j_id209": "",
                "hisQueryForm:j_id213": "",
                "hisQueryForm:j_id217": "",
                "hisQueryForm:j_id221": "",
                "hisQueryForm:firststarttimeInputDate": "2025-07-01 00:00",
                "hisQueryForm:firststarttimeInputCurrentDate": "07/2025",
                "hisQueryForm:firstendtimeInputDate": "2025-07-24 00:00",
                "hisQueryForm:firstendtimeInputCurrentDate": "07/2025",
                "hisQueryForm:j_id229": "",
                "hisQueryForm:recoverstarttimeInputDate": "",
                "hisQueryForm:recoverstarttimeInputCurrentDate": "07/2025",
                "hisQueryForm:recoverendtimeInputDate": "",
                "hisQueryForm:recoverendtimeInputCurrentDate": "07/2025",
                "hisQueryForm:j_id237": "",
                "hisQueryForm:queryFsuStatus_hiddenValue": "",
                "hisQueryForm:currPageObjId": "1",
                "hisQueryForm:pageSizeText": "35",
                "javax.faces.ViewState": "j_id6",
                "hisQueryForm:j_id245": "hisQueryForm:j_id245",
                "AJAX:EVENTS_COUNT": "1"
            },

            {
                "AJAXREQUEST": "_viewRoot",
                "hisQueryForm": "hisQueryForm",
                "hisQueryForm:unitHidden": "",
                "hisQueryForm:unitHid": "",
                "hisQueryForm:queryDay": "30",
                "hisQueryForm:queryFaultMids_hiddenValue": "退服场景",
                "hisQueryForm:queryFaultMids": "退服场景",
                "hisQueryForm:queryFaultDetail": "",
                "hisQueryForm:queryFaultDetailName": "",
                "hisQueryForm:queryLevel_hiddenValue": "",
                "hisQueryForm:j_id201": "",
                "hisQueryForm:j_id205": "",
                "hisQueryForm:j_id209": "",
                "hisQueryForm:j_id213": "",
                "hisQueryForm:j_id217": "",
                "hisQueryForm:j_id221": "",
                "hisQueryForm:firststarttimeInputDate": "2025-07-01 00:00",
                "hisQueryForm:firststarttimeInputCurrentDate": "07/2025",
                "hisQueryForm:firstendtimeInputDate": "2025-07-24 00:00",
                "hisQueryForm:firstendtimeInputCurrentDate": "07/2025",
                "hisQueryForm:j_id229": "",
                "hisQueryForm:recoverstarttimeInputDate": "",
                "hisQueryForm:recoverstarttimeInputCurrentDate": "07/2025",
                "hisQueryForm:recoverendtimeInputDate": "",
                "hisQueryForm:recoverendtimeInputCurrentDate": "07/2025",
                "hisQueryForm:j_id237": "",
                "hisQueryForm:queryFsuStatus_hiddenValue": "",
                "hisQueryForm:currPageObjId": "1",
                "hisQueryForm:pageSizeText": "35",
                "javax.faces.ViewState": "j_id6",
                "hisQueryForm:j_id249": "hisQueryForm:j_id249"
            },
            {
                "j_id407": "j_id407",
                "j_id407:j_id409": "全部",
                "javax.faces.ViewState": "j_id6"
            }
        ]
        self.city_list = [
            "0099977",
            "0099978",
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
            "0099990",
        ]
        self.INDEX = os.getcwd()
        self.save_path = os.path.join(self.INDEX, "xls")
        self.output_path = os.path.join(self.INDEX, "output")
        self.output_name = os.path.join(self.output_path, "故障监控.xlsx")
        self.temp_files = [os.path.join(self.save_path, f"temp_{i}.xls") for i in range(len(self.city_list))]

    def get_view_state(self, url, headers):
        """获取页面的javax.faces.ViewState值"""
        res = requests.get(url=url, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        view_state_input = soup.find('input', id='javax.faces.ViewState')
        if view_state_input:
            return view_state_input.get('value')
        return None

    def spider(self):
        url = self.url
        headers = self.headers
        view_state = self.get_view_state(url, headers)

        for city_idx, city_code in enumerate(self.city_list):
            print(f"正在爬取故障监控工单 {city_idx + 1}/{len(self.city_list)}")

            # 处理前两个数据项
            for i, data in enumerate(self.data_list, start=1):
                if i in [1, 2]:
                    data["hisQueryForm:unitHidden"] = city_code
                    data["hisQueryForm:unitHid"] = city_code
                data["javax.faces.ViewState"] = view_state
                response = requests.post(url=url, data=data, headers=headers)
                if i == len(self.data_list):
                    with open(self.temp_files[city_idx], "wb") as file:
                        file.write(response.content)
                    print(f"城市组 {city_idx + 1} 的临时文件已成功保存到: {self.temp_files[city_idx]}")
            time.sleep(2)

    def merge_excel_files(self):
        """将所有临时文件合并为一个Excel文件"""
        all_data = []
        for file_path in self.temp_files:
            if os.path.exists(file_path):
                try:
                    df = pd.read_excel(file_path)
                    all_data.append(df)
                except Exception as e:
                    print(f"读取文件 {file_path} 时出错: {e}")
        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)
            combined_df.to_excel(self.output_name, index=False)
            print(f"所有文件已成功合并到: {self.output_name}")
        else:
            print("没有找到任何有效的临时文件进行合并。")

    def main(self):
        self.spider()
        self.merge_excel_files()
if __name__ == "__main__":
    fault_monitoring().main()