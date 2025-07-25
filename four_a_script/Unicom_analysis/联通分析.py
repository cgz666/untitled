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

        self.url = "http://omms.chinatowercom.cn:9000/business/resMge/alarmMge/listAlarm.xhtml"
        self.headers = {
            "Accept": "*/*",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Origin": "http://omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/business/resMge/alarmMge/listAlarm.xhtml",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
}

        self.data_list = [
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:querySiteId": "",
                "queryForm:unitHidden": "",
                "queryForm:querySiteNameId": "",
                "queryForm:serialnoText": "",
                "queryForm:hidDeviceId": "",
                "queryForm:objid": "",
                "queryForm:objname": "",
                "queryForm:faultidText": "",
                "queryForm:querystationcode": "",
                "queryForm:fscidText": "",
                "queryForm:selectSignalSize": "0",
                "queryForm:alarmNameMax": "15",
                "queryForm:firststarttimeInputDate": "",
                "queryForm:firststarttimeInputCurrentDate": "07/2025",
                "queryForm:firstendtimeInputDate": "",
                "queryForm:firstendtimeInputCurrentDate": "07/2025",
                "queryForm:j_id71": "",
                "queryForm:j_id75": "",
                "queryForm:j_id79": "",
                "queryForm:queryDWCompany1": "",
                "queryForm:queryDWCompanyName1": "",
                "queryForm:queryDWPersonId": "",
                "queryForm:queryDWPersonName": "",
                "queryForm:queryFactoryName_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "2",
                "queryForm:querystationstatus": "2",
                "queryForm:querySiteSource_hiddenValue": "",
                "queryForm:queryIfConfirm_hiddenValue": "",
                "queryForm:sortSelect_hiddenValue": "",
                "queryForm:isCreateBillSel_hiddenValue": "",
                "queryForm:subOperatorHid_hiddenValue": "",
                "queryForm:refreshTime": "",
                "queryForm:queryCrewAreaIds": "",
                "queryForm:queryCrewAreaName": "",
                "queryForm:queryCrewVillageId": "",
                "queryForm:hideFlag": "",
                "queryForm:queryCrewVillageName": "",
                "queryForm:queryComTypes": "",
                "queryForm:queryComTypeNames": "",
                "queryForm:queryStaTypeSelId_hiddenValue": "",
                "queryForm:isTransitNodeId_hiddenValue": "",
                "queryForm:j_id135": "",
                "queryForm:queryMmanuFactoryNameSelId_hiddenValue": "",
                "queryForm:Type": "",
                "queryForm:j_id146": "",
                "queryForm:signaldevtype1": "",
                "queryForm:liBattery": "",
                "queryForm:currPageObjId": "1",
                "queryForm:pageSizeText": "35",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "j_id12",
                "queryForm:btn1": "queryForm:btn1"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "queryForm": "queryForm",
                "queryForm:querySiteId": "",
                "queryForm:unitHidden": "",
                "queryForm:querySiteNameId": "",
                "queryForm:serialnoText": "",
                "queryForm:hidDeviceId": "",
                "queryForm:objid": "",
                "queryForm:objname": "",
                "queryForm:faultidText": "",
                "queryForm:querystationcode": "",
                "queryForm:fscidText": "",
                "queryForm:selectSignalSize": "0",
                "queryForm:alarmNameMax": "15",
                "queryForm:firststarttimeInputDate": "",
                "queryForm:firststarttimeInputCurrentDate": "07/2025",
                "queryForm:firstendtimeInputDate": "",
                "queryForm:firstendtimeInputCurrentDate": "07/2025",
                "queryForm:j_id71": "",
                "queryForm:j_id75": "",
                "queryForm:j_id79": "",
                "queryForm:queryDWCompany1": "",
                "queryForm:queryDWCompanyName1": "",
                "queryForm:queryDWPersonId": "",
                "queryForm:queryDWPersonName": "",
                "queryForm:queryFactoryName_hiddenValue": "",
                "queryForm:querystationstatus_hiddenValue": "2",
                "queryForm:querystationstatus": "2",
                "queryForm:querySiteSource_hiddenValue": "",
                "queryForm:queryIfConfirm_hiddenValue": "",
                "queryForm:sortSelect_hiddenValue": "",
                "queryForm:isCreateBillSel_hiddenValue": "",
                "queryForm:subOperatorHid_hiddenValue": "",
                "queryForm:refreshTime": "",
                "queryForm:queryCrewAreaIds": "",
                "queryForm:queryCrewAreaName": "",
                "queryForm:queryCrewVillageId": "",
                "queryForm:hideFlag": "",
                "queryForm:queryCrewVillageName": "",
                "queryForm:queryComTypes": "",
                "queryForm:queryComTypeNames": "",
                "queryForm:queryStaTypeSelId_hiddenValue": "",
                "queryForm:isTransitNodeId_hiddenValue": "",
                "queryForm:j_id135": "",
                "queryForm:queryMmanuFactoryNameSelId_hiddenValue": "",
                "queryForm:Type": "",
                "queryForm:j_id146": "",
                "queryForm:signaldevtype1": "",
                "queryForm:liBattery": "",
                "queryForm:currPageObjId": "1",
                "queryForm:pageSizeText": "35",
                "queryForm:panelOpenedState": "",
                "javax.faces.ViewState": "j_id12",
                "queryForm:j_id166": "queryForm:j_id166",
                "AJAX:EVENTS_COUNT": "1"
            },
            {
                "AJAXREQUEST": "_viewRoot",
                "j_id1548": "j_id1548",
                "javax.faces.ViewState": "j_id12",
                "j_id1548:j_id1551": "j_id1548:j_id1551"
            },
            {
                "j_id1548": "j_id1548",
                "j_id1548:j_id1550": "全部",
                "javax.faces.ViewState": "j_id12"
            }
        ]
        self.city_list = [
            "0099977",
            "0099978",
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
            "0099990"
        ]

        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "联通分析-结果.xlsx")
        self.file_name1 = os.path.join(self.save_path, "活动告警_全量.xlsx")
        self.file_name2 = os.path.join(self.save_path, "活动告警_已派.xlsx")
        self.model_path =  os.path.join(self.save_path, "联通分析-模板.xlsx")

    def get_view_state(self):
        """获取页面的javax.faces.ViewState值"""
        res = requests.get(url=self.url, headers=self.headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        view_state_input = soup.find('input', id='javax.faces.ViewState')
        if view_state_input:
            return view_state_input.get('value')
        return None

    # 活动告警-导出（目前告警-全量）
    def process_data_list(self):
        url = self.url
        headers = self.headers
        view_state = self.get_view_state()
        if not view_state:
            print("无法获取javax.faces.ViewState值，请检查URL和headers配置")
            return
        for i, data in enumerate(self.data_list, start=1):
            data["javax.faces.ViewState"] = view_state
            response = requests.post(url=url, data=data, headers=headers)
            if i == len(self.data_list):
                with open(self.file_name1, "wb") as file:
                    file.write(response.content)
                print("当前工单已保存到:", self.file_name1)

    # 活动告警-导出（目前告警-已派）
    def process_second_file(self):
        url = self.url
        headers = self.headers
        view_state = self.get_view_state()

        # 处理前两个数据项
        for i, data in enumerate(self.data_list, start=1):
            if i in [1, 2]:  # 处理前两个数据项
                data["queryForm:isCreateBillSel_hiddenValue"] = 1
                data["queryForm:isCreateBillSel"] = 1
            data["javax.faces.ViewState"] = view_state
            response = requests.post(url=url, data=data, headers=headers)
            # 保存第三个请求的数据（导出请求）
            if i == len(self.data_list):
                with open(self.file_name2, "wb") as file:
                    file.write(response.content)
                print("当前工单已保存到:", self.file_name2)
            time.sleep(2)

    def excel_process(self):
        """
        处理Excel文件，将指定文件夹中的数据文件内容复制到主表文件中。

        :param index_path: 文件夹路径
        """
        print('1、把数据文件和通报模板放在同一文件夹下')
        print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')

        # 初始化 COM 库
        pythoncom.CoInitialize()

        try:
            # 打开模板文件
            xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
            xl.Visible = False  # 窗口是否可见
            xl.DisplayAlerts = False
            workbook_main = xl.Workbooks.Open(self.model_path)  # 打开上述路径文件
            #===================处理第一个子表======================
            # 打开下载文件
            workbook_data = xl.Workbooks.Open(self.file_name1)
            sheet_data = workbook_data.Sheets('sheet0')
            sheet_main = workbook_main.Sheets('活动告警-全量')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:BJ{last_row}')  # 从A2开始复制

            # 1. 只清除B-BK列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"B2:BK{last_clear_row}").ClearContents()

            # 只粘贴值到目标表
            source_range.Copy()
            sheet_main.Range('B2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False  # 释放剪切板

            # 使用autofill续写A列数据
            sheet_main.Range('A2').AutoFill(sheet_main.Range(f'A2:A{last_row}'), win32.constants.xlFillDefault)

            #===================处理第二个子表======================
            # 获取main_sheet的第二个子表
            workbook_data = xl.Workbooks.Open(self.file_name2)
            sheet_data = workbook_data.Sheets('sheet0')
            sheet_main = workbook_main.Sheets('活动告警-已派')

            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:BJ{last_row}')  # 从A2开始复制
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"B2:BK{last_clear_row}").ClearContents()

            # 只粘贴值到目标表
            source_range.Copy()
            sheet_main.Range('B2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False  # 释放剪切板

            # 使用autofill续写第二个子表的A列数据
            sheet_main.Range('A2').AutoFill(sheet_main.Range(f'A2:A{last_row}'), win32.constants.xlFillDefault)

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
        self.excel_process()
if __name__ == "__main__":
    interface_result().main()