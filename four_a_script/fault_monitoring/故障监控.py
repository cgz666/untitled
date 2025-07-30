import requests
import json
import time
import random
from datetime import datetime, timedelta
import pandas as pd
import os
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


class fault_monitoring():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.url = "http://omms.chinatowercom.cn:9000/business/resMge/faultAlarmMge/listFaultActive.xhtml"
        self.headers = {
            "Accept": "*/*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Origin": "http://omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/business/resMge/faultAlarmMge/listFaultActive.xhtml",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }
        self.city_list = [
            "0099977", "0099978", "0099979", "0099980", "0099981",
            "0099982", "0099983", "0099984", "0099985", "0099986",
            "0099987", "0099988", "0099989", "0099990",
        ]

        # 日期设置
        self.today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        self.first_day_of_month = self.today.replace(day=1)
        self.days_to_crawl = (self.today - self.first_day_of_month).days

        # 路径设置
        self.INDEX = r'F:\untitled\four_a_script\fault_monitoring'
        self.save_path = os.path.join(self.INDEX, "xls")
        self.output_path = os.path.join(self.INDEX, "output")
        os.makedirs(self.save_path, exist_ok=True)
        os.makedirs(self.output_path, exist_ok=True)

        self.output_name = os.path.join(
            self.output_path,
            f"故障监控_{self.first_day_of_month.strftime('%Y%m%d')}_{(self.today - timedelta(days=1)).strftime('%Y%m%d')}.xlsx"
        )

        self.temp_files = []
        # 创建带重试机制的session
        self.session = self._create_retry_session()

    def _create_retry_session(self, retries=3, backoff_factor=0.3, status_forcelist=(500, 502, 504)):
        """创建带自动重试的session对象"""
        session = requests.Session()
        retry = Retry(
            total=retries,  # 总重试次数
            read=retries,  # 读取超时重试次数
            connect=retries,  # 连接超时重试次数
            backoff_factor=backoff_factor,  # 重试间隔：{backoff_factor} * (2 **({retry - 1}))
            status_forcelist=status_forcelist,  # 需要重试的状态码
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        return session

    def get_view_state(self, url, headers):
        """获取页面的javax.faces.ViewState值，增加超时和重试"""
        try:
            res = self.session.get(url=url, headers=headers)  # 延长超时时间
            res.raise_for_status()
            soup = BeautifulSoup(res.text, 'html.parser')
            view_state_input = soup.find('input', id='javax.faces.ViewState')
            return view_state_input.get('value') if view_state_input else None
        except Exception as e:
            print(f"获取ViewState失败: {e}")
            return None

    def get_base_data(self, start_date, end_date):
        """生成基础数据模板"""
        start_date_str = start_date.strftime('%Y-%m-%d 00:00')
        end_date_str = end_date.strftime('%Y-%m-%d 00:00')
        current_month = start_date.strftime('%m/%Y')

        return [
            {
                "AJAXREQUEST": "_viewRoot",
                "hisQueryForm": "hisQueryForm",
                "hisQueryForm:unitHidden": "",
                "hisQueryForm:unitHid": "",
                "hisQueryForm:queryDay": "30",
                "hisQueryForm:queryFaultMids_hiddenValue": "退服场景",
                "hisQueryForm:queryFaultMids": "退服场景",
                "hisQueryForm:queryFaultDetail": "",
                "hisQueryForm:queryFaultDetailName": "",
                "hisQueryForm:queryLevel_hiddenValue": "",
                "hisQueryForm:j_id201": "",
                "hisQueryForm:j_id205": "",
                "hisQueryForm:j_id209": "",
                "hisQueryForm:j_id213": "",
                "hisQueryForm:j_id217": "",
                "hisQueryForm:j_id221": "",
                "hisQueryForm:firststarttimeInputDate": start_date_str,
                "hisQueryForm:firststarttimeInputCurrentDate": current_month,
                "hisQueryForm:firstendtimeInputDate": end_date_str,
                "hisQueryForm:firstendtimeInputCurrentDate": current_month,
                "hisQueryForm:j_id229": "",
                "hisQueryForm:recoverstarttimeInputDate": "",
                "hisQueryForm:recoverstarttimeInputCurrentDate": current_month,
                "hisQueryForm:recoverendtimeInputDate": "",
                "hisQueryForm:recoverendtimeInputCurrentDate": current_month,
                "hisQueryForm:j_id237": "",
                "hisQueryForm:queryFsuStatus_hiddenValue": "",
                "hisQueryForm:currPageObjId": "1",
                "hisQueryForm:pageSizeText": "35",
                "javax.faces.ViewState": "j_id6",
                "hisQueryForm:j_id245": "hisQueryForm:j_id245",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "hisQueryForm": "hisQueryForm",
                "hisQueryForm:unitHidden": "",
                "hisQueryForm:unitHid": "",
                "hisQueryForm:queryDay": "30",
                "hisQueryForm:queryFaultMids_hiddenValue": "退服场景",
                "hisQueryForm:queryFaultMids": "退服场景",
                "hisQueryForm:queryFaultDetail": "",
                "hisQueryForm:queryFaultDetailName": "",
                "hisQueryForm:queryLevel_hiddenValue": "",
                "hisQueryForm:j_id201": "",
                "hisQueryForm:j_id205": "",
                "hisQueryForm:j_id209": "",
                "hisQueryForm:j_id213": "",
                "hisQueryForm:j_id217": "",
                "hisQueryForm:j_id221": "",
                "hisQueryForm:firststarttimeInputDate": start_date_str,
                "hisQueryForm:firststarttimeInputCurrentDate": current_month,
                "hisQueryForm:firstendtimeInputDate": end_date_str,
                "hisQueryForm:firstendtimeInputCurrentDate": current_month,
                "hisQueryForm:j_id229": "",
                "hisQueryForm:recoverstarttimeInputDate": "",
                "hisQueryForm:recoverstarttimeInputCurrentDate": current_month,
                "hisQueryForm:recoverendtimeInputDate": "",
                "hisQueryForm:recoverendtimeInputCurrentDate": current_month,
                "hisQueryForm:j_id237": "",
                "hisQueryForm:queryFsuStatus_hiddenValue": "",
                "hisQueryForm:currPageObjId": "1",
                "hisQueryForm:pageSizeText": "35",
                "javax.faces.ViewState": "j_id6",
                "hisQueryForm:j_id249": "hisQueryForm:j_id249"
            },
            {
                "j_id407": "j_id407",
                "j_id407:j_id409": "全部",
                "javax.faces.ViewState": "j_id6"
            }
        ]

    def spider(self):
        total_days = self.days_to_crawl
        print(f"开始爬取故障监控工单，共 {total_days} 天")

        for day_idx in range(total_days):
            start_date = self.first_day_of_month + timedelta(days=day_idx)
            end_date = start_date + timedelta(days=1)
            date_str = start_date.strftime('%Y%m%d')

            print(f"\n===== 爬取 {start_date.strftime('%Y-%m-%d')} 数据 =====")

            data_list = self.get_base_data(start_date, end_date)
            view_state = self.get_view_state(self.url, self.headers)
            if not view_state:
                print(f"⚠️ 无法获取ViewState，跳过当天爬取")
                continue

            for city_idx, city_code in enumerate(self.city_list):
                temp_file = os.path.join(self.save_path, f"temp_{city_code}_{date_str}.xls")
                self.temp_files.append(temp_file)

                print(f"  处理城市 {city_code}（{city_idx + 1}/{len(self.city_list)}）")

                # 手动重试机制（针对超时等临时错误）
                max_retries = 3
                retry_count = 0
                success = False

                while retry_count < max_retries and not success:
                    try:
                        for i, data in enumerate(data_list, start=1):
                            if i in [1, 2]:
                                data["hisQueryForm:unitHidden"] = city_code
                                data["hisQueryForm:unitHid"] = city_code
                            data["javax.faces.ViewState"] = view_state

                            # 使用带重试的session发送请求，延长超时到60秒
                            response = self.session.post(
                                url=self.url,
                                data=data,
                                headers=self.headers,
                            )
                            response.raise_for_status()

                            if i == len(data_list):
                                with open(temp_file, "wb") as file:
                                    file.write(response.content)
                                print(f"  ✅ 城市 {city_code} 保存成功")
                                success = True

                        # 随机间隔1-3秒，避免请求过于密集
                        time.sleep(random.uniform(1, 3))

                    except requests.exceptions.ReadTimeout:
                        retry_count += 1
                        wait_time = 2 ** retry_count  # 指数退避等待时间
                        print(f"  ⏳ 城市 {city_code} 读取超时，第 {retry_count}/{max_retries} 次重试（等待 {wait_time} 秒）")
                        time.sleep(wait_time)
                    except Exception as e:
                        print(f"  ❌ 城市 {city_code} 失败: {str(e)}，停止重试")
                        break

                if not success:
                    print(f"  ❌ 城市 {city_code} 达到最大重试次数，跳过")

    def merge_excel_files(self):
        all_data = []
        for file_path in self.temp_files:
            if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                try:
                    df = pd.read_excel(file_path)
                    if not df.empty:
                        all_data.append(df)
                    else:
                        print(f"  文件 {file_path} 为空")
                except Exception as e:
                    print(f"  读取 {file_path} 错误: {e}")

        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)
            columns_to_convert = ['站址资源编码', '站址运维ID']
            for col in columns_to_convert:
                if col in combined_df.columns:
                    combined_df[col] = combined_df[col].astype(str)

            combined_df.to_excel(self.output_name, index=False)
            print(f"\n✅ 合并完成，保存至 {self.output_name}")
        else:
            print("\n❌ 无有效数据可合并")

    def main(self):
        self.spider()
        self.merge_excel_files()


if __name__ == "__main__":
    fault_monitoring().main()