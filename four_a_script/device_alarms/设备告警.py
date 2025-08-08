import requests
import json
import pandas as pd
import time
import zipfile
import os
import threading
import shutil
from queue import Queue
from datetime import datetime, timedelta


class device_alarms():
    def __init__(self):
        # 获取Cookie
        cookie_url = "http://clound.gxtower.cn:3980/tt/get_aiot_cookie?id=2"
        res = requests.get(cookie_url)
        self.cookie = res.text.strip()
        # 计算日期范围（修改为只爬取今天数据）
        self.calculate_date_range()
        # 基础配置 - 区分活动告警和历史告警URL
        self.active_url = "https://zlzywg.chinatowercom.cn:8070/api/semp/sempAlarm/queryAlarmActive"
        self.history_url = "https://zlzywg.chinatowercom.cn:8070/api/semp/sempAlarm/queryAlarmHistory"
        self.headers = {
            "Accept": "application/json, text/plain, */*",
            "Authorization": f"{self.cookie}",
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
        self.output_path = os.path.join(INDEX, "output")
        self.file_name1 = os.path.join(self.save_path, "活动告警.xlsx")
        self.file_name2 = os.path.join(self.save_path, "关闭告警.xlsx")
        self.file_name3 = os.path.join(self.save_path, "恢复告警.xlsx")
        self.output_name = os.path.join(self.output_path, "设备告警.zip")
        # 进度文件路径
        self.progress_dir = os.path.join(INDEX, "progress")
        self.progress_file = os.path.join(self.progress_dir, "main_progress.json")
        self.data_dir = os.path.join(self.progress_dir, "data")
        # 创建必要的目录
        for dir_path in [self.progress_dir, self.data_dir, self.save_path, self.output_path]:
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)
        # 状态映射配置 - 新增时间字段和表头信息
        self.status_mapping = {
            "0": {
                "url": self.active_url,
                "file": self.file_name1,
                "name": "活动告警",
                "time_field": "updateTime",  # 活动告警使用updateTime
                "time_header": "更新时间",  # 表头为更新时间
                "multi_thread": False
            },
            2: {
                "url": self.history_url,
                "file": self.file_name2,
                "name": "关闭告警",
                "time_field": "alarmCloseTime",  # 关闭告警使用alarmCloseTime
                "time_header": "结束（关闭）时间",  # 表头为结束（关闭）时间
                "multi_thread": False
            },
            1: {
                "url": self.history_url,
                "file": self.file_name3,
                "name": "恢复告警",
                "time_field": "alarmCloseTime",  # 恢复告警使用alarmCloseTime
                "time_header": "结束（恢复）时间",  # 表头为结束（恢复）时间
                "multi_thread": True,
                "thread_count": 5  # 恢复告警使用5个线程
            }
        }
        # 读取主进度
        self.progress = self.load_progress()

    def calculate_date_range(self):
        """计算开始日期和结束日期（只爬取今天的数据）"""
        now = datetime.now()
        # 开始日期：今天0点0分0秒
        self.start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
        # 结束日期：当前时间（确保包含今天截止到现在的所有数据）
        self.end_date = now

        # 转换为字符串格式
        self.start_date_str = self.start_date.strftime("%Y-%m-%d %H:%M:%S")
        self.end_date_str = self.end_date.strftime("%Y-%m-%d %H:%M:%S")

        print(f"开始日期: {self.start_date_str}")
        print(f"结束日期: {self.end_date_str}")

    def load_progress(self, status=None, thread_id=None):
        """加载进度信息"""
        if status is None:
            # 加载主进度
            if os.path.exists(self.progress_file):
                with open(self.progress_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            return {"status": "0", "completed": {}}

        # 加载特定线程的进度
        if thread_id is not None:
            progress_file = os.path.join(self.progress_dir, f"progress_status_{status}_thread_{thread_id}.json")
            if os.path.exists(progress_file):
                with open(progress_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            return {"page": 1, "completed": False, "last_updated": ""}

        # 加载特定状态的进度
        progress_file = os.path.join(self.progress_dir, f"progress_status_{status}.json")
        if os.path.exists(progress_file):
            with open(progress_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {"page": 1, "completed": False, "last_updated": ""}

    def save_progress(self, status, page=None, thread_id=None, completed=False, data=None):
        """保存进度信息和数据"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if thread_id is not None:
            # 保存特定线程的进度
            progress_file = os.path.join(self.progress_dir, f"progress_status_{status}_thread_{thread_id}.json")
            progress = {"page": page, "completed": completed, "last_updated": timestamp}
            with open(progress_file, "w", encoding="utf-8") as f:
                json.dump(progress, f)

            # 如果有数据，保存数据到临时文件
            if data is not None:
                data_file = os.path.join(self.data_dir, f"data_status_{status}_thread_{thread_id}.json")
                with open(data_file, "w", encoding="utf-8") as f:
                    json.dump(data, f)
            return

        if status is None:
            # 保存主进度
            self.progress["last_updated"] = timestamp
            with open(self.progress_file, "w", encoding="utf-8") as f:
                json.dump(self.progress, f)
            return

        # 保存特定状态的进度
        progress_file = os.path.join(self.progress_dir, f"progress_status_{status}.json")
        progress = {"page": page, "completed": completed, "last_updated": timestamp}
        with open(progress_file, "w", encoding="utf-8") as f:
            json.dump(progress, f)

    def load_data(self, status, thread_id=None):
        """加载已保存的数据"""
        if thread_id is not None:
            data_file = os.path.join(self.data_dir, f"data_status_{status}_thread_{thread_id}.json")
            if os.path.exists(data_file):
                with open(data_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            return []

        # 加载特定状态的所有线程数据
        all_data = []
        thread_count = self.status_mapping.get(status, {}).get("thread_count", 1)

        for i in range(thread_count):
            data_file = os.path.join(self.data_dir, f"data_status_{status}_thread_{i}.json")
            if os.path.exists(data_file):
                with open(data_file, "r", encoding="utf-8") as f:
                    all_data.extend(json.load(f))

        return all_data

    def get_total_pages(self, url, status):
        """获取指定状态的总页数"""
        page_size = 100
        params = {
            "admProvinceCode": "450000",
            "alarmStatus": status,
            "alarmTimeStart": self.start_date_str,  # 使用计算的开始日期
            "alarmTimeEnd": self.end_date_str,  # 使用计算的结束日期
            "pageNum": 1,
            "pageSize": page_size
        }

        try:
            response = requests.post(url=url, headers=self.headers, json=params)
            response.raise_for_status()
            data = response.json()

            total = data.get("total") or data.get("data", {}).get("total")
            if total is not None:
                total_pages = (total + page_size - 1) // page_size
                return total_pages
            return None
        except Exception as e:
            print(f"获取总页数失败: {str(e)[:100]}")
            return None

    def download_pages(self, status, url, config, page_range, thread_id, result_queue):
        """下载指定页码范围的数据"""
        page_size = 100
        status_data = []
        max_empty_retry = 3  # 最大连续空页重试次数
        max_request_retry = 3  # 单页请求失败最大重试次数

        # 获取当前状态对应的时间字段
        time_field = config["time_field"]

        start_page, end_page = page_range
        progress = self.load_progress(status, thread_id)

        # 从进度中恢复当前页和数据
        current_page = progress.get("page", start_page)
        if current_page < start_page:
            current_page = start_page

        # 恢复已收集的数据
        status_data = self.load_data(status, thread_id)

        print(f"线程 {thread_id} 恢复进度: 从第 {current_page} 页开始, 已收集 {len(status_data)} 条数据")

        while current_page <= end_page:
            page_url = f"{url}?pageNum={current_page}&pageSize={page_size}"
            params = {
                "admProvinceCode": "450000",
                "alarmStatus": status,
                "alarmTimeStart": self.start_date_str,
                "alarmTimeEnd": self.end_date_str,
            }

            empty_page_count = 0  # 当前连续空页计数
            request_failed_count = 0  # 当前页请求失败计数
            records = None

            # 单页请求重试机制
            while request_failed_count < max_request_retry:
                try:
                    response = requests.post(url=page_url, headers=self.headers, json=params)
                    response.raise_for_status()
                    data = response.json()

                    # 解析records逻辑，兼容多种可能的返回结构
                    if "data" in data:
                        if "records" in data["data"]:
                            records = data["data"]["records"]
                        elif isinstance(data["data"], list):
                            records = data["data"]
                        elif "data" in data["data"] and isinstance(data["data"]["data"], list):
                            records = data["data"]["data"]

                    # 成功获取响应，跳出请求重试循环
                    break

                except requests.exceptions.RequestException as e:
                    request_failed_count += 1
                    print(
                        f"线程 {thread_id}: 第 {current_page} 页请求失败（尝试 {request_failed_count}/{max_request_retry}）: {str(e)[:100]}")
                    time.sleep(2)  # 等待2秒后重试

            # 检查请求是否达到最大重试次数
            if request_failed_count >= max_request_retry:
                print(f"线程 {thread_id}: 第 {current_page} 页请求失败次数过多，保存进度后退出")
                # 保存当前进度和数据
                self.save_progress(status, current_page, thread_id, False, status_data)
                return

            # 处理空页逻辑
            if not records:
                empty_page_count += 1
                if empty_page_count >= max_empty_retry:
                    print(f"线程 {thread_id}: 第 {current_page} 页连续 {max_empty_retry} 次空数据，保存进度后终止下载")
                    self.save_progress(status, current_page, thread_id, False, status_data)
                    break
                else:
                    print(f"线程 {thread_id}: 第 {current_page} 页暂时无数据，等待后重试...")
                    time.sleep(2)  # 等待2秒后继续循环重试当前页
                    continue
            else:
                # 有数据则重置计数器
                empty_page_count = 0

                # 处理数据 - 使用当前状态对应的时间字段
                for item in records:
                    row = [
                        item.get("alarmMsgId", ""),
                        item.get("alarmTitle", ""),
                        item.get("alarmTypeName", ""),
                        item.get("devName", ""),
                        item.get("stationName", ""),
                        item.get("alarmTime", ""),
                        item.get(time_field, ""),  # 使用对应的时间字段
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
                    ]
                    status_data.append(row)

                print(
                    f"线程 {thread_id}: 已获取第 {current_page}/{end_page} 页数据，本页共 {len(records)} 条，累计 {len(status_data)} 条")

                # 每处理5页保存一次进度和数据，或最后一页也保存
                if current_page % 5 == 0 or current_page == end_page:
                    self.save_progress(status, current_page + 1, thread_id, False, status_data)
                current_page += 1

            time.sleep(0.5)  # 减少请求频率

        # 标记线程完成
        self.save_progress(status, None, thread_id, True, status_data)
        result_queue.put(status_data)
        print(f"线程 {thread_id} 下载完成，共获取 {len(status_data)} 条数据")

    def fetch_and_save_alarms(self):
        page_size = 100
        # 基础列名（不含时间相关列）
        base_columns = [
            "告警流水号", "告警名称", "告警类型", "设备名称", "站址名称", "首次时间",
            "设备编码", "站址编码", "管理区域（省）", "管理区域（市）", "管理区域（区）",
            "业务类型", "资源类型", "设备种类", "设备类型", "告警来源"
        ]
        # 遍历不同的状态值
        for status in [1, 2, "0"]:
            try:
                status_int = int(status)
            except ValueError:
                print(f"状态值 {status} 无法转换为整数，跳过")
                continue

            # 检查是否已经完成该状态的下载
            if status in self.progress.get("completed", {}):
                if self.progress["completed"][status]:
                    print(f"\n===== 【{self.status_mapping[status]['name']}】已完成下载，跳过 =====")
                    continue

            config = self.status_mapping[status]
            print(f"\n===== 开始获取【{config['name']}】数据（状态值: {status}）=====")

            # 获取总页数
            total_pages = self.get_total_pages(config["url"], status)
            if total_pages is None:
                print("无法获取总页数，跳过当前状态")
                continue
            print(f"总页数: {total_pages}")

            # 构建当前状态的完整列名（插入时间列）
            time_header = config["time_header"]
            columns = base_columns[:6] + [time_header] + base_columns[6:]

            # 单线程下载
            if not config.get("multi_thread", False):
                status_data = []

                # 获取当前进度
                progress = self.load_progress(status)
                page = progress.get("page", 1)

                # 恢复已收集的数据
                status_data = self.load_data(status)

                while page <= total_pages:
                    url = f"{config['url']}?pageNum={page}&pageSize={page_size}"
                    params = {
                        "admProvinceCode": "450000",
                        "alarmStatus": status,
                        "alarmTimeStart": self.start_date_str,  # 使用计算的开始日期
                        "alarmTimeEnd": self.end_date_str,  # 使用计算的结束日期
                    }

                    try:
                        response = requests.post(url=url, headers=self.headers, json=params)
                        response.raise_for_status()
                        data = response.json()

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
                            break

                        # 处理数据 - 使用当前状态对应的时间字段
                        time_field = config["time_field"]
                        for item in records:
                            row = [
                                item.get("alarmMsgId", ""),
                                item.get("alarmTitle", ""),
                                item.get("alarmTypeName", ""),
                                item.get("devName", ""),
                                item.get("stationName", ""),
                                item.get("alarmTime", ""),
                                item.get(time_field, ""),  # 使用对应的时间字段
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
                            ]
                            status_data.append(row)

                        print(f"已获取第 {page}/{total_pages} 页数据，本页共 {len(records)} 条，累计 {len(status_data)} 条")

                        # 每处理5页保存一次进度和数据，或最后一页也保存
                        if page % 5 == 0 or page == total_pages:
                            self.save_progress(status, page + 1, None, False)
                            self.save_progress(status, page + 1, None, False, status_data)
                        page += 1

                    except requests.exceptions.RequestException as e:
                        print(f"第 {page} 页请求失败: {str(e)[:100]}")
                        time.sleep(5)  # 等待5秒后重试
                        continue

                    time.sleep(0.5)  # 减少请求频率

                # 保存数据
                if status_data:
                    df = pd.DataFrame(status_data, columns=columns)
                    df.to_excel(config["file"], sheet_name=config["name"], index=False)
                    print(f"\n【{config['name']}】爬取完成，共 {len(status_data)} 条，已保存至：{config['file']}")
                else:
                    print(f"\n【{config['name']}】未获取到任何数据\n")

                # 标记状态完成
                self.progress.setdefault("completed", {})[status] = True
                self.save_progress(None)

            # 多线程下载
            else:
                thread_count = config.get("thread_count", 5)
                result_queue = Queue()
                threads = []

                # 计算每个线程负责的页数
                pages_per_thread = (total_pages + thread_count - 1) // thread_count

                # 创建并启动线程
                for i in range(thread_count):
                    start_page = i * pages_per_thread + 1
                    end_page = min((i + 1) * pages_per_thread, total_pages)

                    # 检查该线程是否已经完成
                    thread_progress = self.load_progress(status, i)
                    if thread_progress.get("completed", False):
                        print(f"线程 {i} 已完成下载 {start_page}-{end_page} 页，跳过")
                        continue

                    thread = threading.Thread(
                        target=self.download_pages,
                        args=(status, config["url"], config, (start_page, end_page), i, result_queue)
                    )
                    thread.daemon = True
                    threads.append(thread)
                    thread.start()
                    time.sleep(0.5)  # 线程启动间隔

                # 等待所有线程完成
                for thread in threads:
                    thread.join()

                # 收集所有线程的结果
                all_data = self.load_data(status)

                # 保存数据
                if all_data:
                    df = pd.DataFrame(all_data, columns=columns)
                    df.to_excel(config["file"], sheet_name=config["name"], index=False)
                    print(f"\n【{config['name']}】多线程爬取完成，共 {len(all_data)} 条，已保存至：{config['file']}")
                else:
                    print(f"\n【{config['name']}】未获取到任何数据\n")

                # 标记状态完成
                self.progress.setdefault("completed", {})[status] = True
                self.save_progress(None)

    def merge_excel_files(self):
        files_to_zip = [
            self.file_name1,  # 活动告警
            self.file_name2,  # 关闭告警
            self.file_name3  # 恢复告警
        ]

        # 检查文件是否存在并压缩
        files_added = []
        try:
            with zipfile.ZipFile(self.output_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in files_to_zip:
                    if os.path.exists(file_path):
                        # 获取文件名(不包含路径)
                        file_name = os.path.basename(file_path)
                        # 将文件添加到ZIP文件中
                        zipf.write(file_path, arcname=file_name)
                        files_added.append(file_name)
                        print(f"已添加到压缩包: {file_name}")
                    else:
                        print(f"文件不存在，跳过: {file_path}")

            if files_added:
                print(f"\n成功创建压缩文件: {self.output_name}")
                print(f"包含文件: {', '.join(files_added)}")
            else:
                print("\n没有找到任何文件进行压缩")
                if os.path.exists(self.output_name):
                    os.remove(self.output_name)
                    print("已删除空的压缩文件")

        except Exception as e:
            print(f"压缩文件时出错: {str(e)}")

    def main(self):
        try:
            self.fetch_and_save_alarms()
            self.merge_excel_files()
        finally:
            # 清空progress文件夹
            if os.path.exists(self.progress_dir):
                shutil.rmtree(self.progress_dir)  # 递归删除文件夹及其内容
                os.makedirs(self.progress_dir)  # 重新创建文件夹
                print(f"已清空进度文件夹：{self.progress_dir}")


if __name__ == "__main__":
    device_alarms().main()