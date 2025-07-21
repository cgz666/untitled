import requests
import json
import time
import pythoncom
import pandas as pd
import win32com.client as win32
import os

class custom_workorder_yys():
    def __init__(self):
        cookie_url = "http://10.19.6.250:5000/get_4a_cookie"
        res = requests.get(cookie_url)
        cookie_str = res.text.strip()
        cookie_dict = json.loads(cookie_str)
        cookie_header = "; ".join([f"{k}={v}" for k, v in cookie_dict.items()])

        self.base_url = "http://omms.chinatowercom.cn:9000/portal/SelfTaskController/exportExcel"
        self.headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Cookie": f"{cookie_header}",
            "Host": "omms.chinatowercom.cn:9000",
            "Referer": "http://omms.chinatowercom.cn:9000/portal/iframe.html?modules=selfTask/views/taskListIndex",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
}
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "联通调度工单核实现场设备情况-结果.xlsx")
        self.file_name = os.path.join(self.save_path, "当前工单.xlsx")
        self.model_path = os.path.join(self.save_path, "模板.xls")

    def spider(self):
        params = {
          "queryType": "1",
          "orgId": "0098364",
          "BUSI_TYPE": "1",
          "status": "5,6,7,8,10,9",
          "templateName": "联通调度",
          "yunjianStatus": "5,6,7,8,10,9",
          "pageName": "taskListIndex"
        }
        res = requests.get(url=self.base_url, headers=self.headers, params=params)
        res.raise_for_status()  # 检查请求是否成功
        with open(self.file_name, "wb") as f:
            f.write(res.content)
        print(f"数据已成功保存到 {self.file_name}")

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
            xl.DisplayAlerts = False
            xl.Visible = False  # 窗口是否可见
            workbook_main = xl.Workbooks.Open(self.model_path)  # 打开上述路径文件

            # 打开下载文件
            workbook_data = xl.Workbooks.Open(self.file_name)
            sheet_data = workbook_data.Sheets('工单信息')
            sheet_main = workbook_main.Sheets('目前已派清单（需处理-累计）')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:AE{last_row}')  # 从A2开始复制

            # 1. 只清除A-F列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"A2:AE{last_clear_row}").ClearContents()

            # 只粘贴值到目标表
            source_range.Copy()
            sheet_main.Range('A2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False  # 释放剪切板

            workbook_main.SaveAs(self.output_name)
            workbook_main.Close()
            xl.Quit()
            print('已全部完成')
        except Exception as e:
            raise
        finally:
            # 释放 COM 库
            pythoncom.CoUninitialize()

    def main(self):
        self.spider()
        self.excel_process()
if __name__ == "__main__":
    custom_workorder_yys().main()
