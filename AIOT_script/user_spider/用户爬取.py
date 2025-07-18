import requests
import json
import time
import pandas as pd
import os

class user_spider():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])
        self.url1 = "https://energy-iot.chinatowercom.cn/api/admin/system/user/list"
        self.headers1 = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Authorization": "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsiIl0sInVzZXJfbmFtZSI6IndlYl9tYW5hZ2V8cWprLWNoZW5neiIsInNjb3BlIjpbImFsbCJdLCJleHAiOjE3NTI4NDQ0MjIsInVzZXJJZCI6MzQyMDEsImp0aSI6IjM1YmNkOTQ0LTFlMDQtNGE1Ni1hZTYzLTJiM2U0MmYxMjUwZiIsImNsaWVudF9pZCI6IndlYl9tYW5hZ2UifQ.jVJ0PXMl0w6UmYm172zTQA5deH0lP5c4IVFNOxwRg1yTLdZ4tUGWPRMxPGGkgjyX6Qw3Nrbx5irVVkrcdvVC0fCly8JnE54uDYne-nPhJQj3-mavr0C3-7ADam13j7FQreU3g8I_i1vBDsRyFbn-atp0fnLjKwP_yrYoQ8d69jURgGZroROvZSSeUUtAq7WX6tMjzExG6Qeue6o8dxLAhgejNVChsMCOvGarAfhiQDW9ikjbHUeLNZrbZqxGtleiorMF29ysw7YZ3DQKM3zP5EU5HvbPVVqw9eAajFSAaL00CfqCVxogd5umHQlNmL42WxqOfNyzNjiJhM9Pga3ZUw",
            "Connection": "keep-alive",
            "Content-Type": "application/json;charset=UTF-8",
            "Cookie": f"{cookie_header}",
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
        self.headers2 ={
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Authorization": "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsiIl0sInVzZXJfbmFtZSI6IndlYl9tYW5hZ2V8cWprLWNoZW5neiIsInNjb3BlIjpbImFsbCJdLCJleHAiOjE3NTI4NDQ0MjIsInVzZXJJZCI6MzQyMDEsImp0aSI6IjM1YmNkOTQ0LTFlMDQtNGE1Ni1hZTYzLTJiM2U0MmYxMjUwZiIsImNsaWVudF9pZCI6IndlYl9tYW5hZ2UifQ.jVJ0PXMl0w6UmYm172zTQA5deH0lP5c4IVFNOxwRg1yTLdZ4tUGWPRMxPGGkgjyX6Qw3Nrbx5irVVkrcdvVC0fCly8JnE54uDYne-nPhJQj3-mavr0C3-7ADam13j7FQreU3g8I_i1vBDsRyFbn-atp0fnLjKwP_yrYoQ8d69jURgGZroROvZSSeUUtAq7WX6tMjzExG6Qeue6o8dxLAhgejNVChsMCOvGarAfhiQDW9ikjbHUeLNZrbZqxGtleiorMF29ysw7YZ3DQKM3zP5EU5HvbPVVqw9eAajFSAaL00CfqCVxogd5umHQlNmL42WxqOfNyzNjiJhM9Pga3ZUw",
            "Connection": "keep-alive",
            "Cookie": f"{cookie_header}",
            "Host": "energy-iot.chinatowercom.cn",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
            "sec-ch-ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"'
        }
        self.url3 = "https://energy-iot.chinatowercom.cn/api/device/device/page"
        self.headers3 ={
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Authorization": "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsiIl0sInVzZXJfbmFtZSI6IndlYl9tYW5hZ2V8cWprLWNoZW5neiIsInNjb3BlIjpbImFsbCJdLCJleHAiOjE3NTI4NTA3NzIsInVzZXJJZCI6MzQyMDEsImp0aSI6ImEyZDE0Yjc0LTI5MjctNGM4ZS1iNzdhLWVjODEwOGUyY2VjMSIsImNsaWVudF9pZCI6IndlYl9tYW5hZ2UifQ.s9EfVv4iEyTovtlPM1BABa8Jjwlp9HX6KOmTVPSReFlqyJ9QblwHc9ph8cLHHjC8KwfVmPrkW71H4s86fc_ZB8Zk2Aw3u3INCQ_rUYhxcVRsYyIwoA8qyBinhEplNOLX0npgk3gxF6PoqyLPdEpuG7AuCKQDrwFxRY4ZzPkdUEzIlKO_xTF9HuCKk-pjW-o8QboHGKP387Sa9B9LnWUvBcPQ4fiGzdSH2UVuGgWkLufTiJQ1x2g5b52q_J6ElTxclc9q7Pb7-8wBuagvfWKY9w1SdQl4JLq2rMxDAYyisfHca-WqeF1XpJKi07A5OHi3TjPEvscFP927UZrNurTbhA",
            "Connection": "keep-alive",
            "Content-Length": "91",
            "Content-Type": "application/json;charset=UTF-8",
            "Cookie": f"{cookie_header}",
            "Host": "energy-iot.chinatowercom.cn",
            "Origin": "https://energy-iot.chinatowercom.cn",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
            "sec-ch-ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"'
        }
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "结果.xlsx")
        self.file_name1 = os.path.join(self.save_path, "用户信息.xlsx")

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
            try:
                response = requests.post(url=self.url1, headers=self.headers1, json=params)
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

            except requests.exceptions.RequestException as e:
                print(f"请求出错: {e}")
                break

        print(f"所有用户ID已提取完成，总数：{len(ids)}")
        return ids

    def spider_user(self, user_ids):
        """获取多个用户详情并保存到Excel"""
        if not user_ids:
            print("没有用户ID可获取详情")
            return

        user_details = []
        failed_ids = []

        print(f"开始获取{len(user_ids)}个用户的详情...")
        for i, user_id in enumerate(user_ids, 1):
            url = f"https://energy-iot.chinatowercom.cn/api/admin/system/user/base/detail/{user_id}"
            try:
                response = requests.get(url=url, headers=self.headers2)
                response.raise_for_status()  # 检查请求是否成功
                data = response.json()

                # 从data子对象中提取信息
                user_data = data.get("data", {})
                phone = user_data.get("phonenumber", "")
                name = user_data.get("realName", "")

                # 如果没有名字，则跳过该用户，不做任何记录
                if not name:
                    continue

                user_details.append({
                    "userId": user_id,
                    "realName": name,
                    "phonenumber": phone
                })

                print(f"({i}/{len(user_ids)}) 获取用户 {user_id} 详情成功: {name} - {phone}")

                # 添加延时避免请求过快
                time.sleep(0.3)

            except requests.exceptions.RequestException as e:
                print(f"({i}/{len(user_ids)}) 获取用户 {user_id} 详情失败: {e}")
                failed_ids.append(user_id)
            except Exception as e:
                print(f"({i}/{len(user_ids)}) 处理用户 {user_id} 数据时出错: {e}")
                failed_ids.append(user_id)

        # 保存到Excel
        if user_details:
            df = pd.DataFrame(user_details)
            df.to_excel(self.file_name1, index=False, sheet_name="用户信息")
            print(f"用户详情已保存到 {self.file_name1}，共保存 {len(user_details)} 条记录")

            if failed_ids:
                print(f"共有 {len(failed_ids)} 个用户获取失败: {failed_ids}")
        else:
            print("没有获取到任何用户详情")

    def spider_dev(self):
        """爬取设备信息，并根据用户信息匹配"""
        user_df = pd.read_excel(self.file_name1)
        user_info = user_df.to_dict(orient="records")

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

                try:
                    response = requests.post(url=self.url3, headers=self.headers3, json=params)
                    response.raise_for_status()  # 检查请求是否成功
                    response_data = response.json()

                    rows = response_data.get("rows", [])
                    for row in rows:
                        chartered_airplane_name = row.get("charteredAirplaneName", "")
                        for user in user_info:
                            if user["realName"] == chartered_airplane_name:
                                matched_data = {
                                    "包机人": user["realName"],
                                    "手机号码": user["phonenumber"],
                                    "点位编码": row.get("stationPubCode", ""),
                                    "设备编码": row.get("devCode", ""),
                                    "业务类型": row.get("businessTypeName", ""),
                                }
                                all_matched_data.append(matched_data)
                                # 每次匹配成功时输出该条数据
                                print(f"匹配成功: {matched_data}")
                                break

                    # 更新总记录数
                    if total == 0:
                        total = response_data.get("total", 0)
                    if len(rows) == 0 or page * page_size >= total:
                        break
                    page += 1

                    time.sleep(0.5)

                except requests.exceptions.RequestException as e:
                    print(f"请求出错: {e}")
                    break

        # 保存匹配后的数据到文件
        df = pd.DataFrame(all_matched_data)
        df.to_excel(self.output_name, index=False, sheet_name="设备信息")
        print(f"匹配后的设备信息已保存到 {self.output_name}，共保存 {len(all_matched_data)} 条记录")


    def main(self):
        # user_ids = self.spider_id()
        # self.spider_user(user_ids)
        self.spider_dev()

if __name__ == "__main__":
    user_spider().main()