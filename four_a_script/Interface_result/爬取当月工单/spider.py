import requests
import os
import json
import pythoncom
from bs4 import BeautifulSoup
import time
import win32com.client as win32
from datetime import datetime, timedelta

class interface_result():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.url = "http://omms.chinatowercom.cn:9000/billDeal/monitoring/list/billList.xhtml"
        self.headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }
        self.data_list = [
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:msg": "0",
                "queryForm:queryBillId": "",
                "queryForm:queryBillSn": "",
                "queryForm:isQueryHis": "W",
                "queryForm:queryStationId": "",
                "queryForm:deviceidText": "",
                "queryForm:addOrEditAreaNameId": "",
                "queryForm:aid": "",
                "queryForm:queryUnitId": "0099977",
                "queryForm:j_id48": "",
                "queryForm:queryDWCompany": "",
                "queryForm:queryDWCompanyName": "",
                "queryForm:queryAlarmId": "",
                "queryForm:queryAlarmName": "",
                "queryForm:j_id58": "",
                "queryForm:starttimeInputDate": "2025-06-01 00:00",
                "queryForm:starttimeInputCurrentDate": "06/2025",
                "queryForm:starttimeTimeHours": "00",
                "queryForm:starttimeTimeMinutes": "00",
                "queryForm:endtimeInputDate": "2025-06-30 23:59",
                "queryForm:endtimeInputCurrentDate": "06/2025",
                "queryForm:endtimeTimeHours": "23",
                "queryForm:endtimeTimeMinutes": "59",
                "queryForm:revertstarttimeInputDate": "2025-06-01 00:00",
                "queryForm:revertstarttimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeTimeHours": "00",
                "queryForm:revertstarttimeTimeMinutes": "00",
                "queryForm:revertendtimeInputDate": "2025-06-30 23:59",
                "queryForm:revertendtimeInputCurrentDate": "06/2025",
                "queryForm:revertendtimeTimeHours": "23",
                "queryForm:revertendtimeTimeMinutes": "59",
                "queryForm:dealstarttimeInputDate": "",
                "queryForm:dealstarttimeInputCurrentDate": "07/2025",
                "queryForm:dealendtimeInputDate": "",
                "queryForm:dealendtimeInputCurrentDate": "07/2025",
                "queryForm:sitesource_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "",
                "queryForm:billStatus_hiddenValue": "",
                "queryForm:faultSrc_hiddenValue": "铁塔集团动环网管,移动运营商接口,联通运营商接口,电信运营商接口",
                "queryForm:faultSrc": "铁塔集团动环网管",
                "queryForm:isHasten_hiddenValue": "",
                "queryForm:alarmlevel_hiddenValue": "",
                "queryForm:faultDevType_hiddenValue": "",
                "queryForm:isOverTime_hiddenValue": "",
                "queryForm:isReplyOver_hiddenValue": "",
                "queryForm:subOperatorHid_hiddenValue": "",
                "queryForm:operatorLevel_hiddenValue": "",
                "queryForm:turnSend_hiddenValue": "",
                "queryForm:sortSelect_hiddenValue": "",
                "queryForm:faultTypeId_hiddenValue": "",
                "queryForm:queryCrewProvinceId": "",
                "queryForm:queryCrewCityId": "",
                "queryForm:queryCrewAreaId": "",
                "queryForm:queryCrewVillageId": "",
                "queryForm:hideFlag": "",
                "queryForm:queryCrewVillageName": "",
                "queryForm:refreshTime": "",
                "queryForm:isTurnBack_hiddenValue": "",
                "queryForm:deleteproviceIdHidden": "",
                "queryForm:deletecityIdHidden": "",
                "queryForm:deletecountryIdHidden": "",
                "queryForm:queryDeleteCountyName": "",
                "queryForm:isTransitNodeId_hiddenValue": "",
                "queryForm:j_id139": "",
                "queryForm:j_id143": "",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "j_id5",
                "queryForm:btn": "queryForm:btn"
            },

            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:msg": "0",
                "queryForm:queryBillId": "",
                "queryForm:queryBillSn": "",
                "queryForm:isQueryHis": "W",
                "queryForm:queryStationId": "",
                "queryForm:deviceidText": "",
                "queryForm:addOrEditAreaNameId": "",
                "queryForm:aid": "",
                "queryForm:queryUnitId": "0099977",
                "queryForm:j_id48": "",
                "queryForm:queryDWCompany": "",
                "queryForm:queryDWCompanyName": "",
                "queryForm:queryAlarmId": "",
                "queryForm:queryAlarmName": "",
                "queryForm:j_id58": "",
                "queryForm:starttimeInputDate": "2025-06-01 00:00",
                "queryForm:starttimeInputCurrentDate": "06/2025",
                "queryForm:starttimeTimeHours": "00",
                "queryForm:starttimeTimeMinutes": "00",
                "queryForm:endtimeInputDate": "2025-06-30 23:59",
                "queryForm:endtimeInputCurrentDate": "06/2025",
                "queryForm:endtimeTimeHours": "23",
                "queryForm:endtimeTimeMinutes": "59",
                "queryForm:revertstarttimeInputDate": "2025-06-01 00:00",
                "queryForm:revertstarttimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeTimeHours": "00",
                "queryForm:revertstarttimeTimeMinutes": "00",
                "queryForm:revertendtimeInputDate": "2025-06-30 23:59",
                "queryForm:revertendtimeInputCurrentDate": "06/2025",
                "queryForm:revertendtimeTimeHours": "23",
                "queryForm:revertendtimeTimeMinutes": "59",
                "queryForm:dealstarttimeInputDate": "",
                "queryForm:dealstarttimeInputCurrentDate": "07/2025",
                "queryForm:dealendtimeInputDate": "",
                "queryForm:dealendtimeInputCurrentDate": "07/2025",
                "queryForm:sitesource_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "",
                "queryForm:billStatus_hiddenValue": "",
                "queryForm:faultSrc_hiddenValue": "铁塔集团动环网管,移动运营商接口,联通运营商接口,电信运营商接口",
                "queryForm:faultSrc": "铁塔集团动环网管",
                "queryForm:isHasten_hiddenValue": "",
                "queryForm:alarmlevel_hiddenValue": "",
                "queryForm:faultDevType_hiddenValue": "",
                "queryForm:isOverTime_hiddenValue": "",
                "queryForm:isReplyOver_hiddenValue": "",
                "queryForm:subOperatorHid_hiddenValue": "",
                "queryForm:operatorLevel_hiddenValue": "",
                "queryForm:turnSend_hiddenValue": "",
                "queryForm:sortSelect_hiddenValue": "",
                "queryForm:faultTypeId_hiddenValue": "",
                "queryForm:queryCrewProvinceId": "",
                "queryForm:queryCrewCityId": "",
                "queryForm:queryCrewAreaId": "",
                "queryForm:queryCrewVillageId": "",
                "queryForm:hideFlag": "",
                "queryForm:queryCrewVillageName": "",
                "queryForm:refreshTime": "",
                "queryForm:isTurnBack_hiddenValue": "",
                "queryForm:deleteproviceIdHidden": "",
                "queryForm:deletecityIdHidden": "",
                "queryForm:deletecountryIdHidden": "",
                "queryForm:queryDeleteCountyName": "",
                "queryForm:isTransitNodeId_hiddenValue": "",
                "queryForm:j_id139": "",
                "queryForm:j_id143": "",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "j_id5",
                "queryForm:j_id150": "queryForm:j_id150",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "j_id1945": "j_id1945",
                "j_id1945:j_id1947": "N",
                "j_id1945:devExport": "全部",
                "javax.faces.ViewState": "j_id6"
            }

        ]
        self.city_list = [
            "0099977,0108557,0108559,0108560,0108562,0108563,0108564,0108565,0108566,2735327428",
            "0099978,0108569,0108570,0108571,0108572,0108573,0108574,0108575,0108576,0108577,0108578,2788568264,2788570112,2788568312,2788567212,2788568343",
            "0099979",
            "0099980",
            "0099981",
            "0099982",
            "0099983",
            "0099984",
            "0099985",
            "0099986",
            "0099987",
            "0099988",
            "0099989",
            "0099990,0108648,0108649,0108650,0108651",
        ]

        INDEX = os.getcwd()
        print(INDEX)
        self.xls_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "当月归档工单.xlsx")
        self.temp_files = [os.path.join(self.xls_path, f"temp_{i}.xls") for i in range(len(self.city_list))]

        self.current_date = datetime.now()
        self.first_day_of_month = self.current_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        self.last_day_of_month = self.get_last_day_of_month()
        self.first_day_of_month_str = self.first_day_of_month.strftime("%Y-%m-%d %H:%M")
        self.last_day_of_month_str = self.last_day_of_month.strftime("%Y-%m-%d %H:%M")
        self.current_month_str = self.current_date.strftime("%m/%Y")

    def get_last_day_of_month(self):
        """获取当月最后一天"""
        next_month = self.current_date.replace(day=28) + timedelta(days=4)
        return next_month - timedelta(days=next_month.day)

    def get_view_state(self,url, headers):
        """获取页面的javax.faces.ViewState值"""
        res = requests.get(url=url, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        view_state_input = soup.find('input', id='javax.faces.ViewState')
        if view_state_input:
            return view_state_input.get('value')
        return None


    # 爬取一号到当前工单，根据city_list分组爬取
    def process_second_file(self):
        url = self.url
        headers = self.headers
        view_state = self.get_view_state(url, headers)
        for city_idx, city_code in enumerate(self.city_list):
            print(f"正在爬取当天前历史工单 {city_idx + 1}/{len(self.city_list)}")

            # 处理前两个数据项
            for i, data in enumerate(self.data_list, start=1):
                if i in [1, 2]:  # 处理前两个数据项
                    # data["queryForm:isQueryHis"] = "W"
                    # data["queryForm:starttimeInputDate"] = self.first_day_of_month_str
                    # data["queryForm:starttimeInputCurrentDate"] = self.current_month_str
                    # data["queryForm:endtimeInputDate"] = self.last_day_of_month_str
                    # data["queryForm:endtimeInputCurrentDate"] = self.current_month_str
                    # data["queryForm:revertstarttimeInputDate"] = self.first_day_of_month_str
                    # data["queryForm:revertstarttimeInputCurrentDate"] = self.current_month_str
                    # data["queryForm:revertendtimeInputDate"] = self.last_day_of_month_str
                    # data["queryForm:revertendtimeInputCurrentDate"] = self.current_month_str
                    # data["queryForm:dealstarttimeInputCurrentDate"] = self.current_month_str
                    # data["queryForm:dealendtimeInputCurrentDate"] = self.current_month_str
                    data["queryForm:queryUnitId"] = city_code
                data["javax.faces.ViewState"] = view_state
                response = requests.post(url=url, data=data, headers=headers)
                # 保存第三个请求的数据（导出请求）
                if i == len(self.data_list):
                    with open(self.temp_files[city_idx], "wb") as file:
                        file.write(response.content)
                    print(f"城市组 {city_idx + 1} 的临时文件已成功保存到: {self.temp_files[city_idx]}")
            time.sleep(2)

    # 合并Excel文件（保留原格式，仅第一个文件保留表头）
    def merge_excel_files(self):
        print("开始合并Excel文件（保留原格式）...")
        pythoncom.CoInitialize()
        try:
            # 启动Excel应用程序
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False  # 后台运行

            # 创建新的工作簿作为主文件
            wb_master = excel.Workbooks.Add()
            ws_master = wb_master.Sheets(1)
            ws_master.Name = "综合查询信息-导出"  # 设置表名

            first_file_processed = False

            # 遍历所有临时文件
            for file_idx, file in enumerate(self.temp_files):
                if not os.path.exists(file):
                    print(f"警告：文件 {file} 不存在，跳过")
                    continue

                # 打开临时文件
                wb_temp = excel.Workbooks.Open(file)
                ws_temp = wb_temp.Sheets("综合查询信息-导出")  # 假设工作表名固定

                # 获取临时文件的数据范围
                last_row_temp = ws_temp.Cells(ws_temp.Rows.Count, 1).End(win32.constants.xlUp).Row
                last_col_temp = ws_temp.Cells(1, ws_temp.Columns.Count).End(win32.constants.xlToLeft).Column

                # 处理第一个文件：复制表头和所有数据
                if not first_file_processed:
                    if last_row_temp >= 1:
                        # 复制表头
                        header_range = ws_temp.Range(ws_temp.Cells(1, 1), ws_temp.Cells(1, last_col_temp))
                        header_range.Copy()
                        ws_master.Range("A1").PasteSpecial(Paste=win32.constants.xlPasteAll)

                        # 复制数据（如果有）
                        if last_row_temp > 1:
                            data_range = ws_temp.Range(ws_temp.Cells(2, 1), ws_temp.Cells(last_row_temp, last_col_temp))
                            data_range.Copy()
                            ws_master.Range("A2").PasteSpecial(Paste=win32.constants.xlPasteAll)

                            # 更新主表的最后一行
                            last_row_master = last_row_temp - 1  # 减去表头行
                        else:
                            last_row_master = 0  # 只有表头，没有数据

                        first_file_processed = True
                        print(f"已处理第一个文件：{file}")
                    else:
                        print(f"警告：第一个文件 {file} 为空，跳过")
                        wb_temp.Close(SaveChanges=False)
                        continue

                # 处理后续文件：只复制数据（从第二行开始）
                else:
                    if last_row_temp > 1:  # 确保有数据行（至少两行，第一行为表头）
                        # 计算目标粘贴位置（主表的下一行）
                        target_row = last_row_master + 2  # +2 是因为主表的行号从1开始，且要跳过表头

                        # 复制数据（从第二行开始）
                        data_range = ws_temp.Range(ws_temp.Cells(2, 1), ws_temp.Cells(last_row_temp, last_col_temp))
                        data_range.Copy()
                        ws_master.Range(f"A{target_row}").PasteSpecial(Paste=win32.constants.xlPasteAll)

                        # 更新主表的最后一行
                        last_row_master += (last_row_temp - 1)  # 减去表头行，加上数据行数
                        print(f"已合并文件 {file_idx + 1}/{len(self.temp_files)}：{file}")
                    else:
                        print(f"警告：文件 {file} 只有表头，没有数据，跳过")

                wb_temp.Close(SaveChanges=False)  # 关闭临时文件
                excel.CutCopyMode = False
            # 保存合并后的文件
            if first_file_processed:
                wb_master.SaveAs(self.output_name, FileFormat=win32.constants.xlExcel8)  # 保存为.xls格式
                print(f"合并完成，文件保存至: {self.output_name}")
            else:
                print("警告：没有有效数据可合并")

            wb_master.Close()

        except Exception as e:
            print(f"合并失败: {e}")
            raise
        finally:
            excel.Quit()  # 退出Excel
            pythoncom.CoUninitialize()  # 释放COM资源


    def main(self):
        self.process_second_file()
        self.merge_excel_files()
if __name__ == "__main__":
    interface_result().main()