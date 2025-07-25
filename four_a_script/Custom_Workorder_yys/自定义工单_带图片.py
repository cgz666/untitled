import requests
import json
import zipfile
from datetime import datetime
import pandas as pd
import os
import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

class Custom_Workorder_yys_photo():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.url1 = "http://omms.chinatowercom.cn:9000/portal/SelfTaskController/getSelfTaskList"
        self.url2 = "http://omms.chinatowercom.cn:9000/portal/SelfTaskController/exportExcelAndImage"

        self.headers = {
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Content-Type": "application/json",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Origin": "http://omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/portal/iframe.html?modules/selfTask/views/taskListIndex",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
            "X-Requested-With": "XMLHttpRequest"
        }
        self.INDEX = r'F:\untitled\four_a_script\custom_workorder_yys\xls'
        self.save_path1 = os.path.join(self.INDEX, "当前工单")
        self.save_path2 = os.path.join(self.INDEX, "历史工单")
        self.output_path = os.path.join(self.INDEX, "output")
        self.output_name = os.path.join(self.output_path, "自定义工单.zip")

        # 2. 初始化带重试机制的Session
        self.session = self._init_session()
        # 用于统计总文件数和当前下载进度
        self.total_files = 0
        self.current_file = 0

    def _init_session(self):
        """创建带重试机制的Session，处理网络波动"""
        session = requests.Session()
        retry_strategy = Retry(
            total=3,  # 最大重试次数
            backoff_factor=1,  # 重试间隔：1s, 2s, 4s...
            status_forcelist=[429, 500, 502, 503, 504],  # 触发重试的状态码
            allowed_methods=["GET", "POST"]  # 允许重试的请求方法
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("http://", adapter)  # 为HTTP请求挂载适配器
        session.mount("https://", adapter)
        return session

    def spider1(self):
        if os.path.exists(self.save_path1):
            for file_name in os.listdir(self.save_path1):
                file_path = os.path.join(self.save_path1, file_name)
                if os.path.isfile(file_path):
                    os.remove(file_path)  # 删除文件
                elif os.path.isdir(file_path):
                    os.rmdir(file_path)  # 删除空的子文件夹
            print(f"已清空 {self.save_path1} 文件夹")
        # 逻辑同原代码，改用self.session.post，并添加超时和间隔
        page = 1
        page_size = 15
        data = {
            "queryType": "1",
            "orgId": "0098364",
            "BUSI_TYPE": "1",
            "status": [8],
            "templateName": "联通调度",
            "yunjianStatus": [8],
            "pageName": "taskListIndex",
            "page": page,
            "rows": page_size
        }
        results = []
        while True:
            try:
                response = self.session.post(
                    url=self.url1,
                    headers=self.headers,
                    json=data,
                )
                response.raise_for_status()
            except Exception as e:
                print(f"获取工单列表失败（page={page}）：{e}，跳过当前页")
                break

            response_data = response.json()
            if not response_data.get("rows"):
                break  # 无数据时退出循环

            for item in response_data["rows"]:
                result = {
                    "SITE_NAME": item.get("SITE_NAME", ""),
                    "ID": item.get("ID", ""),
                    "DO_END_TIME": item.get("DO_END_TIME", "")
                }
                if result["DO_END_TIME"]:
                    try:
                        result["DO_END_TIME"] = datetime.strptime(
                            result["DO_END_TIME"], "%Y-%m-%d %H:%M:%S.%f"
                        ).strftime("%Y-%m-%d")
                    except Exception as e:
                        print(f"日期格式错误：{e}，使用原始值")
                results.append(result)

            page += 1
            data["page"] = page
            time.sleep(1)  # 增加1秒间隔，减轻服务器压力

        # 更新总文件数
        self.total_files += len(results)
        print(f"当前工单模块共有 {len(results)} 个文件需要下载")

        # 下载文件（带重试）
        self._download_files(results, self.save_path1, query_type="1", status="8")

    def spider2(self):
        if os.path.exists(self.save_path2):
            for file_name in os.listdir(self.save_path2):
                file_path = os.path.join(self.save_path2, file_name)
                if os.path.isfile(file_path):
                    os.remove(file_path)  # 删除文件
                elif os.path.isdir(file_path):
                    os.rmdir(file_path)  # 删除空的子文件夹
            print(f"已清空 {self.save_path2} 文件夹")
        page = 1
        page_size = 15
        data = {
            "queryType": "2",
            "orgId": "0098364",
            "BUSI_TYPE": "1",
            "status": [11],
            "templateName": "联通调度",
            "yunjianStatus": [11],
            "pageName": "taskListIndex",
            "page": page,
            "rows": page_size
        }
        results = []
        while True:
            try:
                response = self.session.post(
                    url=self.url1,
                    headers=self.headers,
                    json=data,
                )
                response.raise_for_status()
            except Exception as e:
                print(f"获取工单列表失败（page={page}）：{e}，跳过当前页")
                break

            response_data = response.json()
            if not response_data.get("rows"):
                break

            for item in response_data["rows"]:
                result = {
                    "SITE_NAME": item.get("SITE_NAME", ""),
                    "ID": item.get("ID", ""),
                    "DO_END_TIME": item.get("DO_END_TIME", "")
                }
                if result["DO_END_TIME"]:
                    try:
                        result["DO_END_TIME"] = datetime.strptime(
                            result["DO_END_TIME"], "%Y-%m-%d %H:%M:%S.%f"
                        ).strftime("%Y-%m-%d")
                    except Exception as e:
                        print(f"日期格式错误：{e}，使用原始值")
                results.append(result)

            page += 1
            data["page"] = page
            time.sleep(1)

        # 更新总文件数
        previous_total = self.total_files
        self.total_files += len(results)
        print(f"历史工单模块共有 {len(results)} 个文件需要下载")

        # 下载文件（带重试）
        self._download_files(results, self.save_path2, query_type="2", status="11")

    def _download_files(self, results, save_path, query_type, status):
        """通用文件下载方法，带重试逻辑"""
        for index, row in pd.DataFrame(results).iterrows():
            self.current_file += 1
            site_name = row["SITE_NAME"]  # 避免文件名含特殊字符
            task_id = row["ID"]
            do_end_time = row["DO_END_TIME"]
            file_name = f"{site_name}_{task_id}_{do_end_time}.zip"
            file_path = os.path.join(save_path, file_name)

            # 显示进度
            print(f"\n正在下载第 {self.current_file}/{self.total_files} 个文件: {file_name}")

            # 跳过已存在的文件（可选，避免重复下载）
            # if os.path.exists(file_path):
            #     print(f"文件已存在，跳过：{file_path}")
            #     continue

            params = {
                "queryType": query_type,
                "orgId": "0098364",
                "BUSI_TYPE": "1",
                "status": status,
                "taskId": task_id,
                "templateName": "联通调度",
                "pageName": "taskListIndex",
                "isWithImage": "withImage"
            }

            success = False
            for attempt in range(3):
                try:
                    print(f"下载 {file_name}（第{attempt + 1}次尝试）...")
                    response = self.session.get(
                        url=self.url2,
                        headers=self.headers,
                        params=params,
                    )
                    response.raise_for_status()
                    with open(file_path, "wb") as f:
                        f.write(response.content)
                    success = True
                    print(f"保存成功：{file_path}")
                    break
                except (requests.exceptions.ChunkedEncodingError,
                        requests.exceptions.ConnectionError,
                        requests.exceptions.Timeout) as e:
                    print(f"第{attempt + 1}次下载失败：{e}")
                    time.sleep(2)
                except Exception as e:
                    print(f"下载异常：{e}")
                    break

            if not success:
                print(f"⚠️ 多次尝试后仍失败，跳过文件：{file_name}")
            time.sleep(1)

    def combine_zip_files(self):
        """合并文件（逻辑不变，增加异常处理）"""
        try:
            with zipfile.ZipFile(self.output_name, 'w', zipfile.ZIP_DEFLATED) as combined_zip:
                for root, _, files in os.walk(self.save_path1):
                    for file in files:
                        combined_zip.write(
                            os.path.join(root, file),
                            arcname=os.path.relpath(os.path.join(root, file), self.INDEX)
                        )
                for root, _, files in os.walk(self.save_path2):
                    for file in files:
                        combined_zip.write(
                            os.path.join(root, file),
                            arcname=os.path.relpath(os.path.join(root, file), self.INDEX)
                        )
            print(f"合并成功：{self.output_name}")
        except Exception as e:
            print(f"合并文件失败：{e}")

    def main(self):
        self.spider1()
        self.spider2()
        self.combine_zip_files()
        print(f"\n全部下载完成！总共处理了 {self.total_files} 个文件")

if __name__ == "__main__":
    Custom_Workorder_yys_photo().main()