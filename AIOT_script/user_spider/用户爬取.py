import requests
import json
import time
import pandas as pd
import os

class user_spider():
    def __init__(self):
        cookie_url = "http://clound.gxtower.cn:3980/tt/get_aiot_cookie"
        res = requests.get(cookie_url)
        cookie = res.text.strip()
        self.url1 = "https://energy-iot.chinatowercom.cn/api/admin/system/user/list"
        self.headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Authorization": f"{cookie}",
            "Connection": "keep-alive",
            "Content-Type": "application/json;charset=UTF-8",
            "Cookie": "HWWAFSESID=5763e73c92a5fa8982; HWWAFSESTIME=1752800874701; dc04ed2361044be8a9355f6efb",
            "Host": "energy-iot.chinatowercom.cn",
            "Origin": "https://energy-iot.chinatowercom.cn",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
            "sec-ch-ua": "\"Chromium\";v=\"136\", \"Microsoft Edge\";v=\"136\", \"Not.A/Brand\";v=\"99\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\""
        }

        self.url3 = "https://energy-iot.chinatowercom.cn/api/device/device/page"
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "结果.xlsx")

    def spider_id(self):
        """爬取用户ID"""
        page = 1
        page_size = 100
        total = 0
        ids = []
        while True:
            params = {
                "pageNum": page,
                "pageSize": page_size,
                "phonenumber": "",
                "page": page
            }
            response = requests.post(url=self.url1, headers=self.headers, json=params)
            response.raise_for_status()  # 检查请求是否成功
            response_data = response.json()

            rows = response_data.get("rows", [])
            user_ids = [user["userId"] for user in rows]
            ids.extend(user_ids)
            print(f"提取到的用户ID列表（第{page}页）：{user_ids}")

            # 更新总记录数
            if total == 0:
                total = response_data.get("total", 0)
            if len(rows) == 0 or page * page_size >= total:
                break
            page += 1

            # 添加延时避免请求过快
            time.sleep(0.5)

        print(f"所有用户ID已提取完成，总数：{len(ids)}")
        return ids

    def spider_user(self, user_ids):
        """获取多个用户详情并保存为字典"""
        if not user_ids:
            print("没有用户ID可获取详情")
            return {}

        user_details = {}
        print(f"开始获取{len(user_ids)}个用户的详情...")
        for i, user_id in enumerate(user_ids, 1):
            url = f"https://energy-iot.chinatowercom.cn/api/admin/system/user/base/detail/{user_id}"
            response = requests.get(url=url, headers=self.headers)
            response.raise_for_status()  # 检查请求是否成功
            data = response.json()

            # 从data子对象中提取信息
            user_data = data.get("data", {})
            phone = user_data.get("phonenumber", "")
            name = user_data.get("realName", "")

            # 如果没有名字，则跳过该用户，不做任何记录
            if not name:
                continue

            user_details[name] = phone
            print(f"({i}/{len(user_ids)}) 获取用户 {user_id} 详情成功: {name} - {phone}")

            # 添加延时避免请求过快
            time.sleep(0.3)
        return user_details

    def spider_dev(self, user_info):
        """爬取设备信息，并根据用户信息匹配"""
        # 准备爬取设备信息
        business_types = [1, 2, 4]  # 需要遍历的业务类型
        all_matched_data = []  # 用于存储所有匹配后的数据

        for business_type in business_types:
            print(f"开始爬取业务类型为 {business_type} 的设备信息...")
            page = 1
            page_size = 100
            total = 0

            while True:
                params = {
                    "devType": "",
                    "accessPointId": "",
                    "pageNum": page,
                    "pageSize": page_size,
                    "businessType": str(business_type),  # 当前业务类型
                    "deptIds": []
                }

                response = requests.post(url=self.url3, headers=self.headers, json=params)
                response.raise_for_status()  # 检查请求是否成功
                response_data = response.json()

                rows = response_data.get("rows", [])
                for row in rows:
                    chartered_airplane_name = row.get("charteredAirplaneName", "")
                    phone = user_info.get(chartered_airplane_name, "")
                    if phone:
                        matched_data = {
                            "包机人": chartered_airplane_name,
                            "手机号码": phone,
                            "点位编码": row.get("stationPubCode", ""),
                            "设备编码": row.get("devCode", ""),
                            "业务类型": row.get("businessTypeName", ""),
                        }
                        all_matched_data.append(matched_data)
                        print(f"匹配成功: {matched_data}")

                # 更新总记录数
                if total == 0:
                    total = response_data.get("total", 0)
                if len(rows) == 0 or page * page_size >= total:
                    break
                page += 1

                time.sleep(0.5)

        # 保存匹配后的数据到文件
        df = pd.DataFrame(all_matched_data)
        df.to_excel(self.output_name, index=False, sheet_name="设备信息")
        print(f"匹配后的设备信息已保存到 {self.output_name}，共保存 {len(all_matched_data)} 条记录")

    def main(self):
        user_ids = self.spider_id()
        user_info = self.spider_user(user_ids)
        print("用户信息（用户名：电话号码）：")
        print(user_info)
        self.spider_dev(user_info)

if __name__ == "__main__":
    user_spider().main()