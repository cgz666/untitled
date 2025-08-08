import requests
import os
import json
import time
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

class interface_result():
    def __init__(self):
        # 基础URL配置
        self.base_url = "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml"

        # 信号量配置
        self.signal_params = [
            {
                "signal_id": "0406111001",
                "signal_name": "直流电压",
                "temp_files": []
            },
            {
                "signal_id": "0406112001",
                "signal_name": "直流负载电流",
                "temp_files": []
            },
            {
                "signal_id": "0406101001",
                "signal_name": "Ua",
                "temp_files": []
            },
            {
                "signal_id": "0406102001",
                "signal_name": "Ub",
                "temp_files": []
            },
            {
                "signal_id": "0406103001",
                "signal_name": "Uc",
                "temp_files": []
            },
            {
                "signal_id": "0406107001",
                "signal_name": "Ia",
                "temp_files": []
            },
            {
                "signal_id": "0406108001",
                "signal_name": "Ib",
                "temp_files": []
            },
            {
                "signal_id": "0406109001",
                "signal_name": "Ic",
                "temp_files": []
            }
        ]

        # 城市列表
        self.city_list = [
            "0099977", "0099978", "0099979", "0099980", "0099981",
            "0099982", "0099983", "0099984", "0099985", "0099986",
            "0099987", "0099988", "0099989", "0099990"
        ]

        # 文件路径配置
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")  # 临时文件路径
        self.output_path = os.path.join(INDEX, "output")  # 合并后文件路径

        # 确保目录存在
        for path in [self.save_path, self.output_path]:
            if not os.path.exists(path):
                os.makedirs(path)

    def calculate_time_range(self):
        """计算时间区间：结束时间为当前时间，开始时间为结束时间减去2小时"""
        end_time = datetime.now()
        start_time = end_time - timedelta(hours=2)
        return start_time.strftime("%Y-%m-%d %H:%M"), end_time.strftime("%Y-%m-%d %H:%M")

    def get_current_timestamp(self):
        """获取精确到毫秒的时间戳字符串"""
        return datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # 保留毫秒前3位

    def generate_output_filename(self, base_name):
        """生成合并后文件的路径（保存到output_path）"""
        return os.path.join(self.output_path, f"{base_name}_{self.get_current_timestamp()}.xlsx")

    def generate_temp_filename(self, signal_id, city_idx):
        """生成临时文件的路径（保存到save_path，不含时间戳）"""
        return os.path.join(self.save_path, f"temp_{signal_id}_{city_idx}.xlsx")

    def get_cookie(self):
        """获取并更新Cookie"""
        try:
            cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
            res = requests.get(cookie_url)
            res.raise_for_status()
            return "; ".join([f"{k}={v}" for k, v in json.loads(res.text.strip()).items()])
        except Exception as e:
            print(f"获取Cookie失败: {e}")
            return None

    def get_headers(self, cookie_header):
        """生成带最新Cookie的请求头"""
        return {
            "Accept": "*/*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": cookie_header,
            "Host": "omms.chinatowercom.cn:9000",
            "Origin": "http://omms.chinatowercom.cn:9000",
            "Referer": self.base_url,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }

    def get_view_state(self, session):
        """获取页面的javax.faces.ViewState值"""
        try:
            res = session.get(url=self.base_url)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, 'html.parser')
            view_state = soup.find('input', id='javax.faces.ViewState')
            return view_state.get('value') if view_state else None
        except Exception as e:
            print(f"获取ViewState失败: {e}")
            return None

    def is_valid_excel(self, file_path):
        """检查文件是否为有效的Excel文件"""
        return os.path.exists(file_path) and os.path.getsize(file_path) > 1024

    def create_data_template(self, signal_id, start_time, end_time):
        """创建基于信号量和时间区间的请求数据模板"""
        # 解析时间参数
        s_date, s_time = start_time.split()
        e_date, e_time = end_time.split()
        s_hour, s_minute = s_time.split(':')
        e_hour, e_minute = e_time.split(':')
        s_month_year = datetime.strptime(s_date, "%Y-%m-%d").strftime("%m/%Y")
        e_month_year = datetime.strptime(e_date, "%Y-%m-%d").strftime("%m/%Y")

        return [
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:aid": "",
                "queryForm:mongositecode": "",
                "queryForm:siteProvinceId": "0098364",
                "queryForm:queryFlag": "",
                "queryForm:unitHidden1": "",
                "queryForm:unitHidden2": "",  # 城市代码将在此处设置
                "queryForm:unitHidden3": "0098364",
                "queryForm:unitTypeHidden": "undefined",
                "queryForm:siteNameId": "",
                "queryForm:deviceName": "",
                "queryForm:did": "",
                "queryForm:midName": "",
                "queryForm:mid": "",
                "queryForm:queryStationId": "",
                "queryForm:queryStationIdShow": "",
                "queryForm:queryFsuId": "",
                "queryForm:midType": "遥测",
                "queryForm:querySpeId": signal_id,
                "queryForm:querySpeIdShow": f"{signal_id}...",
                "queryForm:starttimeInputDate": start_time,
                "queryForm:starttimeInputCurrentDate": s_month_year,
                "queryForm:starttimeTimeHours": s_hour,
                "queryForm:starttimeTimeMinutes": s_minute,
                "queryForm:endtimeInputDate": end_time,
                "queryForm:endtimeInputCurrentDate": e_month_year,
                "queryForm:endtimeTimeHours": e_hour,
                "queryForm:endtimeTimeMinutes": e_minute,
                "queryForm:querySiteSourceCode": "",
                "queryForm:ifRestrict": "true",
                "queryForm:currPageObjId": "0",
                "queryForm:pageSizeText": "35",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "",  # 将动态更新
                "queryForm:j_id52": "queryForm:j_id52"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:aid": "",
                "queryForm:mongositecode": "",
                "queryForm:siteProvinceId": "0098364",
                "queryForm:queryFlag": "",
                "queryForm:unitHidden1": "",
                "queryForm:unitHidden2": "",  # 城市代码将在此处设置
                "queryForm:unitHidden3": "0098364",
                "queryForm:unitTypeHidden": "undefined",
                "queryForm:siteNameId": "",
                "queryForm:deviceName": "",
                "queryForm:did": "",
                "queryForm:midName": "",
                "queryForm:mid": "",
                "queryForm:queryStationId": "",
                "queryForm:queryStationIdShow": "",
                "queryForm:queryFsuId": "",
                "queryForm:midType": "遥测",
                "queryForm:querySpeId": signal_id,
                "queryForm:querySpeIdShow": f"{signal_id}...",
                "queryForm:starttimeInputDate": start_time,
                "queryForm:starttimeInputCurrentDate": s_month_year,
                "queryForm:starttimeTimeHours": s_hour,
                "queryForm:starttimeTimeMinutes": s_minute,
                "queryForm:endtimeInputDate": end_time,
                "queryForm:endtimeInputCurrentDate": e_month_year,
                "queryForm:endtimeTimeHours": e_hour,
                "queryForm:endtimeTimeMinutes": e_minute,
                "queryForm:querySiteSourceCode": "",
                "queryForm:ifRestrict": "true",
                "queryForm:currPageObjId": "0",
                "queryForm:pageSizeText": "35",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "",  # 将动态更新
                "queryForm:j_id53": "queryForm:j_id53",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "j_id421": "j_id421",
                "j_id421:j_id423": "全部",
                "javax.faces.ViewState": ""  # 将动态更新
            }
        ]

    def spider_signal(self, session, signal_config, start_time, end_time):
        """爬取特定信号量的数据（使用传入的时间参数）"""
        signal_id, signal_name = signal_config["signal_id"], signal_config["signal_name"]
        print(f"\n开始爬取 {signal_name} ({signal_id}) | 时间区间: {start_time} 至 {end_time}")

        # 生成临时文件列表（不含时间戳，保存到save_path）
        temp_files = [self.generate_temp_filename(signal_id, i) for i in range(len(self.city_list))]
        success_count = 0

        for city_idx, city_code in enumerate(self.city_list):
            print(f"爬取城市 {city_idx + 1}/{len(self.city_list)}: {city_code}")

            # 获取最新ViewState
            view_state = self.get_view_state(session)
            if not view_state:
                print(f"跳过城市 {city_code}: 无法获取ViewState")
                continue

            # 准备请求数据
            data_templates = self.create_data_template(signal_id, start_time, end_time)
            request_data = [{**d, "javax.faces.ViewState": view_state,
                             "queryForm:unitHidden2": city_code if "queryForm:unitHidden2" in d else ""}
                            for d in data_templates]

            # 发送请求
            for i, data in enumerate(request_data, 1):
                try:
                    response = session.post(url=self.base_url, data=data)
                    response.raise_for_status()

                    # 保存第三个请求的响应（导出文件）
                    if i == 3:
                        temp_file = temp_files[city_idx]
                        with open(temp_file, "wb") as f:
                            f.write(response.content)

                        if self.is_valid_excel(temp_file):
                            print(f"✓ {signal_name} - 城市 {city_code} 保存至: {temp_file}")
                            success_count += 1
                        else:
                            print(f"! {signal_name} - 城市 {city_code} 文件可能损坏")

                    time.sleep(1 + i * 0.5)  # 智能延时

                except Exception as e:
                    print(f"请求失败 ({signal_name} - 城市 {city_code}, 请求 {i}): {e}")
                    time.sleep(5)  # 出错后延长等待时间
                    break

        signal_config["temp_files"] = temp_files
        return success_count, temp_files

    def merge_files(self, signal_config, start_time, end_time):
        """合并同一信号量的所有城市文件（保存到output_path）"""
        signal_name = signal_config["signal_name"]
        output_file = self.generate_output_filename(signal_name)  # 保存到output_path
        temp_files = signal_config.get("temp_files", [])

        print(f"\n开始合并 {signal_name} 的城市文件 | 时间区间: {start_time} 至 {end_time}")

        try:
            import pandas as pd

            valid_files = [f for f in temp_files if self.is_valid_excel(f)]
            if not valid_files:
                print(f"! 没有找到有效的临时文件，跳过合并")
                return

            # 读取并合并所有有效文件
            dfs = [pd.read_excel(f, engine="openpyxl") for f in valid_files]
            combined_df = pd.concat(dfs, ignore_index=True)

            # 保存合并后的文件
            combined_df.to_excel(output_file, index=False, engine="openpyxl")
            print(f"✓ 合并完成: {output_file}")

            # 清理临时文件
            for f in temp_files:
                if os.path.exists(f):
                    os.remove(f)
            print(f"✓ 已清理 {len(temp_files)} 个临时文件")

        except ImportError:
            print(f"! 请安装pandas和openpyxl: pip install pandas openpyxl")
        except Exception as e:
            print(f"! 合并失败: {e}")

    def run_cycle(self):
        """执行一轮完整的爬取循环（更新时间参数）"""
        print(f"\n===== 开始新的爬取循环 =====")

        # 为整个循环更新时间参数
        start_time, end_time = self.calculate_time_range()
        print(f"本轮循环时间区间: {start_time} 至 {end_time}")

        # 获取最新Cookie
        cookie_header = self.get_cookie()
        if not cookie_header:
            print("获取Cookie失败，跳过本轮循环")
            return False

        # 创建会话并设置请求头
        with requests.Session() as session:
            session.headers.update(self.get_headers(cookie_header))

            # 依次爬取所有信号量
            for signal_config in self.signal_params:
                # 爬取信号量数据（使用本轮循环的时间参数）
                success_count, _ = self.spider_signal(session, signal_config, start_time, end_time)
                print(f"{signal_config['signal_name']} 爬取完成: {success_count}/{len(self.city_list)} 城市成功")

                # 合并文件（保存到output_path）
                self.merge_files(signal_config, start_time, end_time)

                # 信号量之间添加延时
                time.sleep(5)

        return True

    def main(self):
        """主循环：无限循环执行爬取任务"""
        cycle_count = 1
        while True:
            try:
                print(f"\n===== 开始第 {cycle_count} 轮循环 =====")
                success = self.run_cycle()
                print(f"{'✓' if success else '!'} 第 {cycle_count} 轮循环完成")

                cycle_count += 1
                # 每轮循环结束后等待一段时间
                wait_time = 20  # 10分钟，可根据需要调整
                print(f"\n等待 {wait_time // 60} 分钟后开始下一轮...")
                time.sleep(wait_time)

            except KeyboardInterrupt:
                print("\n程序被手动终止")
                break
            except Exception as e:
                print(f"! 发生异常: {e}")
                time.sleep(5)  # 发生异常后等待30秒再重试


if __name__ == "__main__":
    interface_result().main()