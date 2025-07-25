import requests
import json
import pandas as pd
import time
import os
import threading
from queue import Queue


class device_alarms():
    def __init__(self):
        # 获取Cookie
        cookie_url = "http://clound.gxtower.cn:3980/tt/get_aiot_cookie?id=2"
        res = requests.get(cookie_url)
        self.cookie = res.text.strip()

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
        os.makedirs(self.save_path, exist_ok=True)
        self.file_name1 = os.path.join(self.save_path, "活动告警.xlsx")
        self.file_name2 = os.path.join(self.save_path, "关闭告警.xlsx")
        self.file_name3 = os.path.join(self.save_path, "恢复告警.xlsx")

        # 状态映射配置
        self.status_mapping = {
            "0": {
                "url": self.active_url,
                "file": self.file_name1,
                "name": "活动告警",
                "multi_thread": False
            },
            2: {
                "url": self.history_url,
                "file": self.file_name2,
                "name": "关闭告警",
                "multi_thread": False
            },
            1: {
                "url": self.history_url,
                "file": self.file_name3,
                "name": "恢复告警",
                "multi_thread": True,
                "thread_count": 5  # 恢复告警使用5个线程
            }
        }

        # 进度文件路径
        self.progress_dir = os.path.join(INDEX, "progress")
        os.makedirs(self.progress_dir, exist_ok=True)
        self.progress_file = os.path.join(self.progress_dir, "main_progress.json")

        # 读取主进度
        self.progress = self.load_progress()

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
            return {"page": 1, "completed": False}

        # 加载特定状态的进度
        progress_file = os.path.join(self.progress_dir, f"progress_status_{status}.json")
        if os.path.exists(progress_file):
            with open(progress_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return {"page": 1, "completed": False}

    def save_progress(self, status, page=None, thread_id=None, completed=False):
        """保存进度信息"""
        if thread_id is not None:
            # 保存特定线程的进度
            progress_file = os.path.join(self.progress_dir, f"progress_status_{status}_thread_{thread_id}.json")
            progress = {"page": page, "completed": completed}
            with open(progress_file, "w", encoding="utf-8") as f:
                json.dump(progress, f)
            return

        if status is None:
            # 保存主进度
            with open(self.progress_file, "w", encoding="utf-8") as f:
                json.dump(self.progress, f)
            return

        # 保存特定状态的进度
        progress_file = os.path.join(self.progress_dir, f"progress_status_{status}.json")
        progress = {"page": page, "completed": completed}
        with open(progress_file, "w", encoding="utf-8") as f:
            json.dump(progress, f)

    def get_total_pages(self, url, status):
        """获取指定状态的总页数"""
        page_size = 100
        params = {
            "admProvinceCode": "450000",
            "alarmStatus": status,
            "alarmTimeStart": "2025-06-01 00:00:00",
            "alarmTimeEnd": "2025-07-22 00:00:00",
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

        start_page, end_page = page_range
        progress = self.load_progress(status, thread_id)
        current_page = progress.get("page", start_page)

        print(f"线程 {thread_id} 开始下载 {start_page}-{end_page} 页（从第 {current_page} 页开始）")

        while current_page <= end_page:
            page_url = f"{url}?pageNum={current_page}&pageSize={page_size}"
            params = {
                "admProvinceCode": "450000",
                "alarmStatus": status,
                "alarmTimeStart": "2025-06-01 00:00:00",
                "alarmTimeEnd": "2025-07-21 00:00:00",
            }

            try:
                response = requests.post(url=page_url, headers=self.headers, json=params)
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
                    print(f"线程 {thread_id}: 第 {current_page} 页没有数据，终止下载")
                    break

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
                        status
                    ]
                    status_data.append(row)

                print(
                    f"线程 {thread_id}: 已获取第 {current_page}/{end_page} 页数据，本页共 {len(records)} 条，累计 {len(status_data)} 条")
                self.save_progress(status, current_page + 1, thread_id, False)
                current_page += 1

            except requests.exceptions.RequestException as e:
                print(f"线程 {thread_id}: 第 {current_page} 页请求失败: {str(e)[:100]}")
                time.sleep(5)  # 等待5秒后重试
                continue

            time.sleep(0.5)  # 减少请求频率

        # 标记线程完成
        self.save_progress(status, None, thread_id, True)
        result_queue.put(status_data)
        print(f"线程 {thread_id} 下载完成，共获取 {len(status_data)} 条数据")

    def fetch_and_save_alarms(self):
        page_size = 100
        columns = [
            "告警流水号", "告警名称", "告警类型", "设备名称", "站址名称", "首次时间", "更新时间", "设备编码", "站址编码",
            "管理区域（省）", "管理区域（市）", "管理区域（区）", "业务类型", "资源类型", "设备种类", "设备类型", "告警来源", "告警状态"
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

            # 单线程下载
            if not config.get("multi_thread", False):
                status_data = []

                # 获取当前进度
                progress = self.load_progress(status)
                page = progress.get("page", 1)

                while page <= total_pages:
                    url = f"{config['url']}?pageNum={page}&pageSize={page_size}"
                    params = {
                        "admProvinceCode": "450000",
                        "alarmStatus": status,
                        "alarmTimeStart": "2025-06-01 00:00:00",
                        "alarmTimeEnd": "2025-07-21 00:00:00",
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
                                status
                            ]
                            status_data.append(row)

                        print(f"已获取第 {page}/{total_pages} 页数据，本页共 {len(records)} 条，累计 {len(status_data)} 条")
                        self.save_progress(status, page + 1, None, False)
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
                all_data = []
                while not result_queue.empty():
                    all_data.extend(result_queue.get())

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

    def main(self):
        self.fetch_and_save_alarms()


if __name__ == "__main__":
    device_alarms().main()