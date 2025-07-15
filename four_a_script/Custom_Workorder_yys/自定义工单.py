import requests
import json
import time
import pythoncom
import pandas as pd
import win32com.client as win32
import os

class Custom_Workorder():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.base_url = "http://omms.chinatowercom.cn:9000/portal/SelfTaskController/exportExcelAndImage"
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
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "自定义工单-结果.xlsx")
        self.file_name1 = os.path.join(self.save_path, "当前工单.zip")
        self.file_name2 = os.path.join(self.save_path, "历史工单.zip")

    def spider(self):
        query_types = [1, 2]
        for query_type in query_types:
            if query_type == 1:
                file_name = self.file_name1
            elif query_type == 2:
                file_name = self.file_name2
            params = {
                "queryType": query_type,
                "orgId": "0098364",
                "BUSI_TYPE": "1",
                "status": "8",
                "templateName": "联通调度",
                "pageName": "taskListIndex",
                "isWithImage": "withImage"
            }

            max_retries = 100  # 最大重试次数
            retry_delay = 10  # 重试延迟(秒)

            for attempt in range(max_retries):
                try:
                    print(f"\n尝试第 {attempt + 1} 次下载...")

                    # 使用流式下载
                    with requests.get(
                            url=self.base_url,
                            headers=self.headers,
                            params=params,
                            stream=True,
                    ) as response:
                        response.raise_for_status()

                        # 初始化下载统计
                        downloaded = 0
                        start_time = time.time()
                        last_print = 0
                        # 根据 queryType 保存到不同的文件

                        with open(file_name, "wb") as f:
                            for chunk in response.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                                    downloaded += len(chunk)

                                    # 每5秒打印一次进度
                                    if time.time() - last_print > 5:
                                        speed = downloaded / (time.time() - start_time) / 1024
                                        print(f"\r已下载: {downloaded / 1024 / 1024:.2f} MB | 速度: {speed:.2f} KB/s",
                                              end="")
                                        last_print = time.time()

                        # 下载完成后打印总结
                        total_time = time.time() - start_time
                        speed = downloaded / total_time / 1024
                        print(
                            f"\n下载完成! 总大小: {downloaded / 1024 / 1024:.2f} MB | 平均速度: {speed:.2f} KB/s | 耗时: {total_time:.2f}秒")

                        # 验证文件是否完整（通过解压测试）
                        if self.validate_zip_file(file_name):
                            print("文件验证通过")
                            return True
                        else:
                            print("文件验证失败，将重试...")
                            os.remove(file_name)

                except Exception as e:
                    print(f"下载出错: {str(e)}")
                    if os.path.exists(file_name):
                        os.remove(file_name)

                if attempt < max_retries - 1:
                    print(f"等待 {retry_delay} 秒后重试...")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # 指数退避

            print("达到最大重试次数，下载失败")
            return False

    def validate_zip_file(self, file_name):
        """验证ZIP文件是否包含JPG、TXT和XLS三种格式的文件"""

        import zipfile

        required_extensions = {'.jpg', '.txt', '.xls'}
        found_extensions = set()

        with zipfile.ZipFile(file_name) as zip_ref:
            # 检查ZIP文件是否有效
            if zip_ref.testzip() is not None:
                print("ZIP文件损坏或包含错误")
                return False

            # 检查文件类型
            for file_info in zip_ref.infolist():
                filename = file_info.filename.lower()
                if filename.endswith('.jpg'):
                    found_extensions.add('.jpg')
                elif filename.endswith('.txt'):
                    found_extensions.add('.txt')
                elif filename.endswith('.xls'):
                    found_extensions.add('.xls')

                # 如果已经找到所有需要的文件类型，提前退出循环
                if found_extensions == required_extensions:
                    break

            # 验证是否包含所有必需的文件类型
            missing_extensions = required_extensions - found_extensions
            if missing_extensions:
                print(f"ZIP文件中缺少以下类型的文件: {', '.join(missing_extensions)}")
                return False

            print("ZIP文件验证通过，包含所有必需的文件类型")
            return True
    def main(self):
        start_time = time.time()
        print("开始执行自定义工单下载任务...")
        if self.spider():
            print(f"任务成功完成，总耗时: {time.time() - start_time:.2f}秒")
        else:
            print(f"任务失败，总耗时: {time.time() - start_time:.2f}秒")

if __name__ == "__main__":
    Custom_Workorder().main()
