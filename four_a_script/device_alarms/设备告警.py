import requests
import json
import pandas as pd
import time
import os

class device_alarms():
    def __init__(self):
        # 获取Cookie
        cookie_url = "http://clound.gxtower.cn:3980/tt/get_aiot_cookie"
        res = requests.get(cookie_url)
        cookie = res.text.strip()
        print(cookie)
        # 基础配置 - 区分活动告警和历史告警URL
        self.active_url = "https://zlzywg.chinatowercom.cn:8070/api/semp/sempAlarm/queryAlarmActive"  # 活动告警（0）
        self.history_url = "https://zlzywg.chinatowercom.cn:8070/api/semp/sempAlarm/queryAlarmHistory"  # 历史告警（1/2）

        self.headers = {
            "Accept": "application/json, text/plain, */*",
            "Authorization": f"{cookie}",
            "Connection": "keep-alive",
            "Content-Type": "application/json;charset=UTF-8",
            "Cookie": "HWWAFSESID=2e792732886d588dc3; HWWAFSESTIME=1753059407266",
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

        # 路径配置
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        os.makedirs(self.save_path, exist_ok=True)
        self.file_name1 = os.path.join(self.save_path, "活动告警.xlsx")
        self.file_name2 = os.path.join(self.save_path, "关闭告警.xlsx")
        self.file_name3 = os.path.join(self.save_path, "恢复告警.xlsx")

        # 状态映射配置：包含状态值、对应的URL、保存文件、名称
        self.status_mapping = {
            "0": {
                "url": self.active_url,
                "file": self.file_name1,
                "name": "活动告警"
            },
            2: {
                "url": self.history_url,
                "file": self.file_name2,
                "name": "关闭告警"
            },
            1: {
                "url": self.history_url,
                "file": self.file_name3,
                "name": "恢复告警"
            }
        }

    def fetch_and_save_alarms(self):
        page_size = 100
        columns = [
            "告警流水号", "告警名称", "告警类型", "设备名称", "站址名称", "首次时间", "更新时间", "设备编码", "站址编码",
            "管理区域（省）", "管理区域（市）", "管理区域（区）", "业务类型", "资源类型", "设备种类", "设备类型", "告警来源", "告警状态"
        ]

        # 遍历不同的状态值（0:活动告警，1:关闭告警，2:恢复告警）
        for status in ["0", 2, 1]:
            config = self.status_mapping[status]
            print(f"\n===== 开始获取【{config['name']}】数据（状态值: {status}）=====")
            status_data = []

            has_more_pages = True
            page = 1

            while has_more_pages:
                # 构建带分页参数的URL（不同状态使用不同的基础URL）
                url = f"{config['url']}?pageNum={page}&pageSize={page_size}"
                # 请求参数（保持0为字符串，1/2为整数）
                params = {
                    "admProvinceCode": "450000",
                    "alarmStatus": status,
                    "alarmTimeStart": "2025-06-01 00:00:00",
                    "alarmTimeEnd": "2025-07-21 00:00:00",
                }

                try:
                    # 发送请求（根据状态调用不同的URL）
                    response = requests.post(url=url, headers=self.headers, json=params)
                    response.raise_for_status()
                    data = response.json()

                    # 提取数据记录（兼容不同格式的响应结构）
                    records = None
                    if "data" in data:
                        if "records" in data["data"]:
                            records = data["data"]["records"]
                        elif isinstance(data["data"], list):
                            records = data["data"]
                        elif "data" in data["data"] and isinstance(data["data"]["data"], list):
                            records = data["data"]["data"]

                    if not records:
                        print(f"第 {page} 页没有数据，终止当前状态爬取")
                        has_more_pages = False
                        break  # 无数据时终止

                    # 解析记录，添加状态列
                    for item in records:
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
                            item.get("alarmSourceName", ""),
                            status  # 添加状态列，方便核对
                        ]
                        status_data.append(row)

                    # 输出进度信息
                    print(f"已获取第 {page} 页数据，本页共 {len(records)} 条，累计 {len(status_data)} 条")

                    # 检查总页数，提前终止（避免无效请求）
                    total = data.get("total") or data.get("data", {}).get("total")
                    if total:
                        total_pages = (total + page_size - 1) // page_size
                        if page >= total_pages:
                            print(f"已获取全部 {total_pages} 页数据，终止当前状态爬取")
                            has_more_pages = False
                            break

                    page += 1

                except Exception as e:
                    print(f"第 {page} 页请求失败: {str(e)[:100]}")  # 简化错误信息
                    has_more_pages = False
                    break  # 异常时终止

                time.sleep(1)  # 间隔请求，避免频繁访问

            # 写入对应状态的Excel文件
            if status_data:
                df = pd.DataFrame(status_data, columns=columns)
                df.to_excel(config["file"], sheet_name=config["name"], index=False)
                print(f"\n【{config['name']}】爬取完成，共 {len(status_data)} 条，已保存至：{config['file']}")
            else:
                print(f"\n【{config['name']}】未获取到任何数据\n")

    def main(self):
        self.fetch_and_save_alarms()

if __name__ == "__main__":
    device_alarms().main()