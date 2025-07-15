import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

class main:
    def __init__(self):
        # 初始化API的URL
        self.url = "https://energy-iot.chinatowercom.cn/api/device/device/page"
        # 获取授权令牌
        self.Authorization = requests.get('http://clound.gxtower.cn:3980/tt/get_aiot_cookie').text
        # 设置HTTP请求头
        self.headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Authorization": self.Authorization,
            "Connection": "keep-alive",
            "Content-Type": "application/json;charset=UTF-8",
            "Host": "energy-iot.chinatowercom.cn",
            "Origin": "https://energy-iot.chinatowercom.cn",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
        }
        # 初始化请求参数
        self.params = {"devType": "", "accessPointId": "", "pageNum": 1, "pageSize": 10, "businessType": "1", "deptIds": []}
        # 定义业务类型列表
        self.business_types = [
            {"type": "1", "name": "备电"},
            {"type": "2", "name": "换电"},
            {"type": "4", "name": "低速充电"}
        ]

    def get_page_data(self, bt, page_size, current_page):
        # 复制请求参数并设置页码和页大小
        params = self.params.copy()
        params['pageNum'] = current_page
        params['pageSize'] = page_size
        params['businessType'] = bt['type']
        # 发送POST请求获取数据
        response = requests.post(self.url, json=params, headers=self.headers)
        response.raise_for_status()
        # 返回JSON响应
        return response.json()

    def process_page(self, bt, page_size, current_page):
        # 获取单页数据
        data = self.get_page_data(bt, page_size, current_page)
        all_data = []
        # 遍历数据行，提取并格式化信息
        for item in data["rows"]:
            all_data.append({
                "业务类型": bt['name'],
                "点位编码": item.get("stationPubCode", ""),
                "设备名称": item.get("devName", ""),
                "设备编码": item.get("devCode", ""),
                "设备状态": item.get("statusName", "")
            })
        # 打印进度信息
        print(f"完成{bt['name']}{current_page}")
        return all_data

    def process_business_type(self, bt):
        # 初始化数据列表
        all_data = []
        page_size = 1000
        current_page = 1
        # 获取第一页数据以确定总记录数
        data = self.get_page_data(bt, 10, current_page)
        total_records = data.get("total", 0)
        total_pages = (total_records + page_size - 1) // page_size
        # 打印业务类型和总记录数信息
        print(f"业务类型 {bt['name']} 共发现 {total_records} 条记录，需要获取 {total_pages} 页数据")

        # 使用线程池并行处理多页数据
        with ThreadPoolExecutor(max_workers=20) as executor:
            futures = [executor.submit(self.process_page, bt, page_size, page) for page in range(1, total_pages + 1)]
            for future in futures:
                all_data.extend(future.result())
        return all_data

    def down_aiot(self):
        # 初始化总数据列表
        all_data = []
        # 遍历所有业务类型
        for bt in self.business_types:
            all_data.extend(self.process_business_type(bt))
        # 如果有数据，则保存到Excel文件
        if all_data:
            df = pd.DataFrame(all_data).drop_duplicates()
            filename = "F:/newtowerV2/message/wuye_taizhang/设备信息/设备信息.xlsx"
            df.to_excel(filename, index=False)
            return filename