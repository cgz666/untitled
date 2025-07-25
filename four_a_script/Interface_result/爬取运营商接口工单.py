import requests
import os
import json
import pythoncom
from bs4 import BeautifulSoup
import time
import win32com.client as win32
from datetime import datetime

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
                "queryForm:queryUnitId": "",
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
                "queryForm:endtimeInputDate": "2025-06-10 00:00",
                "queryForm:endtimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeInputDate": "2025-06-01 00:00",
                "queryForm:revertstarttimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeTimeHours": "00",
                "queryForm:revertstarttimeTimeMinutes": "00",
                "queryForm:revertendtimeInputDate": "2025-06-10 00:00",
                "queryForm:revertendtimeInputCurrentDate": "06/2025",
                "queryForm:revertendtimeTimeHours": "00",
                "queryForm:revertendtimeTimeMinutes": "00",
                "queryForm:dealstarttimeInputDate": "",
                "queryForm:dealstarttimeInputCurrentDate": "06/2025",
                "queryForm:dealendtimeInputDate": "",
                "queryForm:dealendtimeInputCurrentDate": "06/2025",
                "queryForm:sitesource_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "",
                "queryForm:billStatus_hiddenValue": "",
                "queryForm:faultSrc_hiddenValue": "移动运营商接口,联通运营商接口,电信运营商接口",
                "queryForm:faultSrc": "移动运营商接口",
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
                "javax.faces.ViewState": "j_id6",
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
                "queryForm:queryUnitId": "",
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
                "queryForm:endtimeInputDate": "2025-06-10 00:00",
                "queryForm:endtimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeInputDate": "2025-06-01 00:00",
                "queryForm:revertstarttimeInputCurrentDate": "06/2025",
                "queryForm:revertstarttimeTimeHours": "00",
                "queryForm:revertstarttimeTimeMinutes": "00",
                "queryForm:revertendtimeInputDate": "2025-06-10 00:00",
                "queryForm:revertendtimeInputCurrentDate": "06/2025",
                "queryForm:revertendtimeTimeHours": "00",
                "queryForm:revertendtimeTimeMinutes": "00",
                "queryForm:dealstarttimeInputDate": "",
                "queryForm:dealstarttimeInputCurrentDate": "06/2025",
                "queryForm:dealendtimeInputDate": "",
                "queryForm:dealendtimeInputCurrentDate": "06/2025",
                "queryForm:sitesource_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "",
                "queryForm:billStatus_hiddenValue": "",
                "queryForm:faultSrc_hiddenValue": "移动运营商接口,联通运营商接口,电信运营商接口",
                "queryForm:faultSrc": "移动运营商接口",
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
                "javax.faces.ViewState": "j_id6",
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
            "2710377449,2710377453"
        ]

        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "爬取运营商接口工单-结果.xlsx")
        self.file_name1 = os.path.join(self.save_path, "当前工单.xls")
        self.file_name2 = os.path.join(self.save_path, "当天前历史工单.xls")
        self.model_path =  os.path.join(self.save_path, "运营商接口工单完成情况统计表.xlsx")
        self.temp_files = [os.path.join(self.save_path, f"temp_{i}.xls") for i in range(len(self.city_list))]

        self.current_date = datetime.now()
        self.first_day_of_month = self.current_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        self.first_day_of_month_str = self.first_day_of_month.strftime("%Y-%m-%d %H:%M")
        self.current_day = self.current_date.replace(hour=0, minute=0)
        self.current_day_str = self.current_day.strftime("%Y-%m-%d %H:%M")
        self.current_month_str = self.current_date.strftime("%m/%Y")
    def get_view_state(self,url, headers):
        """获取页面的javax.faces.ViewState值"""
        res = requests.get(url=url, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        view_state_input = soup.find('input', id='javax.faces.ViewState')
        if view_state_input:
            return view_state_input.get('value')
        return None

    # 通用处理数据列表的函数
    def process_data_list(self):
        url = self.url
        headers = self.headers
        view_state = self.get_view_state(url, headers)


        if not view_state:
            print("无法获取javax.faces.ViewState值，请检查URL和headers配置")
            return
        for i, data in enumerate(self.data_list, start=1):
            if i in [1, 2]:  # 处理前两个数据项
                data["queryForm:isQueryHis"] = "N"
                data["queryForm:starttimeInputDate"] = self.first_day_of_month_str
                data["queryForm:starttimeInputCurrentDate"] = self.current_month_str
                data["queryForm:endtimeInputDate"] = self.current_day_str
                data["queryForm:endtimeInputCurrentDate"] = self.current_month_str
                data["queryForm:revertstarttimeInputDate"] = self.first_day_of_month_str
                data["queryForm:revertstarttimeInputCurrentDate"] = self.current_month_str
                data["queryForm:revertendtimeInputDate"] = self.current_day_str
                data["queryForm:revertendtimeInputCurrentDate"] = self.current_month_str
                data["queryForm:dealstarttimeInputCurrentDate"] = self.current_month_str
                data["queryForm:dealendtimeInputCurrentDate"] = self.current_month_str
            data["javax.faces.ViewState"] = view_state
            response = requests.post(url=url, data=data, headers=headers)
            if i == len(self.data_list):
                with open(self.file_name1, "wb") as file:
                    file.write(response.content)
                print("当前工单已保存到:", self.file_name1)

    # 处理第二个文件（当天前历史工单，根据city_list分组爬取）
    def process_second_file(self):
        url = self.url
        headers = self.headers
        view_state = self.get_view_state(url, headers)
        for city_idx, city_code in enumerate(self.city_list):
            print(f"正在爬取当天前历史工单 {city_idx + 1}/{len(self.city_list)}")

            # 处理前两个数据项
            for i, data in enumerate(self.data_list, start=1):
                if i in [1, 2]:  # 处理前两个数据项
                    data["queryForm:isQueryHis"] = "W"
                    data["queryForm:starttimeInputDate"] = self.first_day_of_month_str
                    data["queryForm:starttimeInputCurrentDate"] = self.current_month_str
                    data["queryForm:endtimeInputDate"] = self.current_day_str
                    data["queryForm:endtimeInputCurrentDate"] = self.current_month_str
                    data["queryForm:revertstarttimeInputDate"] = self.first_day_of_month_str
                    data["queryForm:revertstarttimeInputCurrentDate"] = self.current_month_str
                    data["queryForm:revertendtimeInputDate"] = self.current_day_str
                    data["queryForm:revertendtimeInputCurrentDate"] = self.current_month_str
                    data["queryForm:dealstarttimeInputCurrentDate"] = self.current_month_str
                    data["queryForm:dealendtimeInputCurrentDate"] = self.current_month_str
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
                wb_master.SaveAs(self.file_name2, FileFormat=win32.constants.xlExcel8)  # 保存为.xls格式
                print(f"合并完成，文件保存至: {self.file_name2}")
            else:
                print("警告：没有有效数据可合并")

            wb_master.Close()

        except Exception as e:
            print(f"合并失败: {e}")
            raise
        finally:
            excel.Quit()  # 退出Excel
            pythoncom.CoUninitialize()  # 释放COM资源

    def excel_process(self):
        """
        处理Excel文件，将指定文件夹中的数据文件内容复制到主表文件中。

        :param index_path: 文件夹路径
        """
        print('1、把数据文件和通报模板放在同一文件夹下')
        print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
        pythoncom.CoInitialize()
        try:
            # 打开模板文件
            xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
            xl.Visible = True  # 窗口是否可见
            xl.DisplayAlerts = False
            workbook_main = xl.Workbooks.Open(self.model_path)  # 打开上述路径文件

            # 处理第一个数据文件（当前工单）
            workbook_data = xl.Workbooks.Open(self.file_name1)
            sheet_data = workbook_data.Sheets('综合查询信息-导出')
            sheet_main = workbook_main.Sheets('当前工单')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A1:CN{last_row}')  # 从A2开始复制

            # 清空目标表的内容，但保留CO和CP两列的第二行数据公式
            for col in range(1, sheet_main.UsedRange.Columns.Count + 1):
                if col != sheet_main.Columns('CO').Column and col != sheet_main.Columns('CP').Column:
                    sheet_main.Cells(3, col).Resize(sheet_main.UsedRange.Rows.Count - 2).ClearContents()
                else:
                    sheet_main.Cells(3, col).Resize(sheet_main.UsedRange.Rows.Count - 2).ClearContents()
                    # 保留第二行的公式
                    sheet_main.Cells(2, col).Copy()
                    sheet_main.Cells(3, col).PasteSpecial(Paste=win32.constants.xlPasteFormulas)

            # 复制和粘贴
            source_range.Copy()
            sheet_main.Range('A1').PasteSpecial(Paste=win32.constants.xlPasteAll)  # 使用全部的形式粘贴，保留格式
            xl.CutCopyMode = False  # 释放剪切板

            # 使用autofill续写CO、CP数据
            sheet_main.Range('CO2').AutoFill(sheet_main.Range(f'CO2:CO{last_row}'), win32.constants.xlFillDefault)
            sheet_main.Range('CP2').AutoFill(sheet_main.Range(f'CP2:CP{last_row}'), win32.constants.xlFillDefault)

            workbook_data.Close(SaveChanges=False)

            # 处理第二个数据文件（当天前历史工单）
            workbook_data = xl.Workbooks.Open(self.file_name2)
            sheet_data = workbook_data.Sheets('综合查询信息-导出')
            sheet_main = workbook_main.Sheets('当天前历史工单')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A1:CO{last_row}')  # 从A2开始复制

            # 清空目标表的内容，但保留CP和CQ两列的第二行数据公式
            for col in range(1, sheet_main.UsedRange.Columns.Count + 1):
                if col != sheet_main.Columns('CP').Column and col != sheet_main.Columns('CQ').Column:
                    sheet_main.Cells(3, col).Resize(sheet_main.UsedRange.Rows.Count - 2).ClearContents()
                else:
                    sheet_main.Cells(3, col).Resize(sheet_main.UsedRange.Rows.Count - 2).ClearContents()
                    # 保留第二行的公式
                    sheet_main.Cells(2, col).Copy()
                    sheet_main.Cells(3, col).PasteSpecial(Paste=win32.constants.xlPasteFormulas)

            # 复制和粘贴
            source_range.Copy()
            sheet_main.Range('A1').PasteSpecial(Paste=win32.constants.xlPasteAll)  # 使用全部的形式粘贴，保留格式
            xl.CutCopyMode = False  # 释放剪切板

            # 使用autofill续写CP、CQ数据
            sheet_main.Range('CP2').AutoFill(sheet_main.Range(f'CP2:CP{last_row}'), win32.constants.xlFillDefault)
            sheet_main.Range('CQ2').AutoFill(sheet_main.Range(f'CQ2:CQ{last_row}'), win32.constants.xlFillDefault)

            workbook_data.Close(SaveChanges=False)
            workbook_main.SaveAs(self.output_name)
            workbook_main.Close()
            xl.Quit()  # 关闭Excel应用程序
            print('已全部完成')
        except Exception as e:
            raise
        finally:
            # 释放 COM 库
            pythoncom.CoUninitialize()

    def main(self):
        self.process_data_list()
        self.process_second_file()
        self.merge_excel_files()
        self.excel_process()
if __name__ == "__main__":
    interface_result().main()