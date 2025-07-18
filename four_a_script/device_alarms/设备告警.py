import requests
import json
import pythoncom
import pandas as pd
import win32com.client as win32
import os

class device_alarms():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.base_url = "https://zlzywg.chinatowercom.cn:8070/api/semp/sempAlarm/queryAlarmActive"
        self.headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Authorization": "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsiIl0sInVzZXJfbmFtZSI6IndlYl9tYW5hZ2V8d3gtaHVhbmd3bDE0Iiwic2NvcGUiOlsiYWxsIl0sImV4cCI6MTc1MjkxNTY2MCwidXNlcklkIjoxMTM0NzEzLCJqdGkiOiJXNlRmOWRsZWd3dUdWbkk1LS1nUFo1S2ZILVkiLCJjbGllbnRfaWQiOiJ3ZWJfbWFuYWdlIn0.O75PPlwDQ4HlDidUb1HuzEUQp6QhYQylV8sjz2YabzBlWBpyvFUIrQJ3mdvS94oE2-JGo3KSevRhbUwYywHSQtsA96qTzWhNhfxyuAr1_nzAkV9vvvmc7C1E-pylLf2YJbX9exqo97ykAIHn0QeAJftkE_uzFH6m8admyvpuTK7t2sPfC_6KtuM7hlIyePGhGM0KMKsxpyBPGHOcKodzOmrnTKncs_fMmuhivxLJ9Kx-WUkugLlzShMvzTsvAwM51-iHnoitxxRhdHZ-rWIGYDBsPzUIYxO3FnaQOJ6Qt3I7ADPE_xkZNFKh2m0vjXnO9QIapeTrxWoecEDu8UV84Q",
            "Connection": "keep-alive",
            "Content-Type": "application/json;charset=UTF-8",
            "Cookie": f"{cookie_header}",
            "Host": "zlzywg.chinatowercom.cn:8070",
            "Origin": "https://zlzywg.chinatowercom.cn:8070",
            "Referer": "https://zlzywg.chinatowercom.cn:8070/alarmcenter/alarmMonitor",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
            "sec-ch-ua": "\"Chromium\";v=\"136\", \"Microsoft Edge\";v=\"136\", \"Not.A/Brand\";v=\"99\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\""
}
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "设备告警-结果.xlsx")
        self.file_name1 = os.path.join(self.save_path, "活动告警.xlsx")
        self.file_name2 = os.path.join(self.save_path, "关闭告警.xlsx")
        self.file_name3 = os.path.join(self.save_path, "恢复告警.xlsx")

    def fetch_and_save_alarms(self):
        page = 1
        page_size = 100
        write_header = True

        columns = [
            "告警流水号", "告警名称", "告警类型", "设备名称", "站址名称", "首次时间", "更新时间", "设备编码", "站址编码",
            "管理区域（省）", "管理区域（市）", "管理区域（区）", "业务类型", "资源类型", "设备种类", "设备类型", "告警来源"
        ]

        with pd.ExcelWriter(self.file_name1, engine='openpyxl') as writer:
            while True:
                params = {
                    "pageNum": page,
                    "pageSize": page_size,
                    "admProvinceCode": "450000",
                    "alarmTimeStart": "2025-07-12 00:00:00",
                    "alarmTimeEnd": "2025-07-18 23:59:59",
                    "alarmStatus": "0"
                }

                try:
                    response = requests.post(url=self.base_url, headers=self.headers, json=params)
                    response.raise_for_status()  # 检查请求是否成功
                    data = response.json()
                except Exception as e:
                    print(f"请求异常: {e}")
                    break

                if data["code"] == 200:
                    if not data["data"]["data"]:
                        print(f"第 {page} 页没有数据")
                        break

                    rows = []
                    for item in data["data"]["data"]:
                        row = [
                            item.get("alarmMsgId", ""),
                            item.get("alarmTitle", ""),
                            item.get("alarmTypeName", ""),
                            item.get("devName", ""),
                            item.get("stationName", ""),
                            item.get("alarmTime", ""),
                            item.get("createTime", ""),
                            item.get("devCode", ""),
                            item.get("stationCode", ""),
                            item.get("admProvinceName", ""),
                            item.get("admCityName", ""),
                            item.get("admCountyName", ""),
                            item.get("devBusinessType", ""),
                            item.get("devResType", ""),
                            item.get("devType", ""),
                            item.get("devChildType", ""),
                            item.get("alarmSourceName", "")
                        ]
                        rows.append(row)

                    df = pd.DataFrame(rows, columns=columns)
                    df.to_excel(writer, sheet_name='活动告警', index=False, header=write_header)
                    print(f"已写入第 {page} 页数据")

                    total = data["data"]["total"]
                    total_pages = (total + page_size - 1) // page_size
                    write_header = False

                    if page >= total_pages:
                        print(f"所有数据写入完成，共 {page} 页")
                        break
                else:
                    print(f"请求失败：{data.get('msg', '未知错误')}")
                    break

                page += 1

    def main(self):
        self.fetch_and_save_alarms()

if __name__ == "__main__":
    device_alarms().main()