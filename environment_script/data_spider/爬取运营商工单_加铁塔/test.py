import requests
import os
import json
import pythoncom
from bs4 import BeautifulSoup
import pandas as pd
import time
import win32com.client as win32
from datetime import datetime
from tqdm import tqdm  # 导入进度条库


class Charging_current():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.url = "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/realTimePerformanceMge/realTimeperfdata.xhtml"
        self.headers = {
            "Accept": "*/*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Origin": "http://omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/realTimePerformanceMge/realTimeperfdata.xhtml",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }

        # 基本请求数据模板
        self.base_data_template = {
            "AJAXREQUEST": "_viewRoot",
            "queryForm": "queryForm",
            "queryForm:addOrEditAreaNameId": "",
            "queryForm:aid": "",
            "queryForm:fsuid": "",
            "queryForm:deviceName": "",
            "queryForm:mid": "0406301001",
            "queryForm:midName": "电池充电限流设定值",
            "queryForm:panelOpenedState": "",
            "javax.faces.ViewState": "j_id6"
        }

        INDEX = r"F:\untitled\four_a_script\Charging_current"
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "爬取运营商接口工单-结果.xlsx")
        self.file_name = os.path.join(self.save_path, "充电电流设定值.xls")
        self.data = os.path.join(self.save_path, '站址信息.csv')

        # 存储从CSV读取的站址信息
        self.site_info = []

    def read_site_info(self):
        """从CSV文件读取站址信息"""
        try:
            if not os.path.exists(self.data):
                print(f"站址信息文件不存在: {self.data}")
                return False

            # 定义可能的编码列表
            encodings = ['utf-8', 'gbk', 'gb2312', 'iso-8859-1', 'utf-16']
            df = None

            # 尝试不同的编码
            for encoding in encodings:
                try:
                    # 尝试使用当前编码读取文件，并设置low_memory=False
                    df = pd.read_csv(self.data, encoding=encoding, low_memory=False)
                    print(f"成功使用 {encoding} 编码读取文件")
                    break
                except UnicodeDecodeError:
                    print(f"无法使用 {encoding} 编码读取文件")
                    continue
                except Exception as e:
                    print(f"使用 {encoding} 编码读取文件时发生未知错误: {e}")
                    continue

            if df is None:
                # 所有编码尝试均失败，使用更宽容的错误处理
                print("尝试了所有编码格式，但仍无法读取文件，使用错误替换策略")
                for encoding in encodings:
                    try:
                        df = pd.read_csv(self.data, encoding=encoding, errors='replace', low_memory=False)
                        print(f"使用 {encoding} 编码读取文件（替换错误字符）")
                        break
                    except Exception as e:
                        continue

                if df is None:
                    print("无法读取文件，所有方法均失败")
                    return False

            # 定义表头映射关系 - 根据实际CSV文件修改这些映射
            header_mapping = {
                'A': '站址编码',  # 实际CSV中站址编码列的名称
                'B': '名称',  # 实际CSV中名称列的名称
                'C': '运维ID',  # 实际CSV中运维ID列的名称
                'D': '站址名备注'  # 实际CSV中站址名备注列的名称
            }

            # 检查必要的列是否存在
            missing_columns = []
            for key, value in header_mapping.items():
                if value not in df.columns:
                    missing_columns.append(f"{key}({value})")

            if missing_columns:
                print(f"CSV文件缺少必要的列: {', '.join(missing_columns)}")
                # 打印实际列名帮助调试
                print(f"CSV文件的实际列名: {', '.join(df.columns.tolist())}")
                return False

            # 提取所需信息
            for _, row in df.iterrows():
                self.site_info.append({
                    '站址编码': str(row[header_mapping['A']]),  # 确保转为字符串
                    '名称': str(row[header_mapping['B']]),  # 确保转为字符串
                    '运维ID': str(row[header_mapping['C']]),  # 确保转为字符串
                    '站址名备注': str(row[header_mapping['D']])  # 确保转为字符串
                })

            print(f"成功读取 {len(self.site_info)} 条站址信息")
            return True

        except Exception as e:
            print(f"读取站址信息时出错: {e}")
            return False

    def get_view_state(self):
        """获取页面的javax.faces.ViewState值"""
        res = requests.post(url=self.url, headers=self.headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        view_state_input = soup.find('input', id='javax.faces.ViewState')
        if view_state_input:
            return view_state_input.get('value')
        return None

    def process_and_save_data(self):
        """获取数据、处理并保存到Excel文件"""
        try:
            # 读取站址信息
            if not self.read_site_info():
                print("无法读取站址信息，程序终止")
                return

            view_state = self.get_view_state()
            if not view_state:
                print("无法获取javax.faces.ViewState值，请检查URL和headers配置")
                return

            all_data = []
            total_sites = len(self.site_info)
            success_count = 0
            fail_count = 0

            print(f"开始处理 {total_sites} 条站址信息...")

            # 使用tqdm显示进度条
            for i, site in enumerate(tqdm(self.site_info, desc="处理进度", unit="站址")):
                try:
                    # 准备请求数据
                    data_list = []

                    # 第一个请求
                    data1 = self.base_data_template.copy()
                    data1["queryForm:did"] = site['站址编码']
                    data1["queryForm:mid"] = site['运维ID']
                    data1["queryForm:j_id21"] = "queryForm:j_id21"
                    data_list.append(data1)

                    # 第二个请求
                    data2 = self.base_data_template.copy()
                    data2["queryForm:did"] = site['站址编码']
                    data2["queryForm:mid"] = site['运维ID']
                    data2["queryForm:j_id22"] = "queryForm:j_id22"
                    data2["AJAX:EVENTS_COUNT"] = "1"
                    data_list.append(data2)

                    # 更新ViewState
                    for data in data_list:
                        data["javax.faces.ViewState"] = view_state

                    # 发送请求
                    for j, data in enumerate(data_list, start=1):
                        response = requests.post(url=self.url, data=data, headers=self.headers)

                        # 只处理最后一个请求的响应
                        if j == len(data_list):
                            soup = BeautifulSoup(response.text, 'html.parser')
                            table_body = soup.find('tbody', id='listForm:list:tb')  # 根据实际HTML结构调整

                            if not table_body:
                                print(f"警告: 未找到站址 {site['名称']} 的表格数据部分")
                                fail_count += 1
                                continue

                            rows = table_body.find_all('tr')
                            site_data_count = 0

                            for row in rows:
                                cols = row.find_all('td')
                                if len(cols) > 0:
                                    # 使用从CSV读取的信息
                                    site_location = site['名称']
                                    site_note = site['站址名备注']
                                    site_id = site['运维ID']
                                    resource_code = site['站址编码']

                                    # 从表格中提取其他信息
                                    device = cols[5].find('a').text.strip() if cols[5].find('a') else cols[
                                        5].text.strip()
                                    signal_id = cols[6].text.strip() if cols[6] else ""
                                    monitoring_point = cols[7].text.strip() if cols[7] else ""
                                    measured_value = cols[8].text.strip() if cols[8] else ""
                                    unit = cols[9].text.strip() if cols[9] else ""
                                    status = cols[10].find('div').text.strip() if cols[10].find('div') else cols[
                                        10].text.strip()

                                    all_data.append({
                                        "站址": site_location,
                                        "站址名备注": site_note,
                                        "站址运维ID": site_id,
                                        "站址资源编码": resource_code,
                                        "设备": device,
                                        "信号量ID": signal_id,
                                        "监控点": monitoring_point,
                                        "实测值": measured_value,
                                        "单位": unit,
                                        "状态": status
                                    })
                                    site_data_count += 1

                            if site_data_count > 0:
                                success_count += 1
                            else:
                                fail_count += 1

                    # 添加延迟避免请求过于频繁
                    time.sleep(1)

                except Exception as e:
                    fail_count += 1
                    print(f"错误: 处理站址 {site['名称']} 时发生异常: {e}")

            print(f"处理完成: 成功 {success_count} 条, 失败 {fail_count} 条, 总计 {total_sites} 条")

            if not all_data:
                print("没有数据可保存")
                return

            df = pd.DataFrame(all_data)

            # 写入Excel文件
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = os.path.join(self.save_path, f"充电电流设定值_{current_time}.xlsx")

            # 使用ExcelWriter可以设置更多选项
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='充电电流数据', index=False)

                # 获取工作簿和工作表对象以进行格式设置
                workbook = writer.book
                worksheet = writer.sheets['充电电流数据']

                # 设置列宽
                for i, col in enumerate(df.columns):
                    column_width = max(len(str(x)) for x in df[col])
                    column_width = max(column_width, len(col)) + 2
                    column_letter = chr(65 + i)  # A, B, C, ...
                    worksheet.column_dimensions[column_letter].width = column_width

            print(f"数据已成功保存到: {file_path}")

        except Exception as e:
            print(f"处理数据或保存Excel文件时出错: {e}")


if __name__ == "__main__":
    Charging_current().process_and_save_data()