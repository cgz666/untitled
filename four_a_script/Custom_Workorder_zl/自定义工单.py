import requests
import json
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

        self.base_url = "http://omms.chinatowercom.cn:9000/portal/SelfTaskController/exportExcel"
        self.headers = {
          "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
          "Accept-Encoding": "gzip, deflate",
          "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
          "Connection": "keep-alive",
          "Cookie": f"{cookie_header}",
          "Host": "omms.chinatowercom.cn:9000",
          "Referer": "http://omms.chinatowercom.cn:9000/portal/iframe.html?modules/selfTask/views/taskListIndex",
          "Upgrade-Insecure-Requests": "1",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }
        INDEX = os.getcwd()
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name = os.path.join(self.output_path, "自定义工单-结果.xlsx")
        self.file_name1 = os.path.join(self.save_path, "当前工单.xlsx")
        self.file_name2 = os.path.join(self.save_path, "历史工单.xlsx")
        self.file_name3 = os.path.join(self.save_path, "合并工单.xlsx")
        self.model_path =  os.path.join(self.save_path, "模板.xlsx")

    def spider(self):
        query_types = [1, 2]
        for query_type in query_types:
            params = {
                "queryType": query_type,
                "orgId": "0098364",
                "BUSI_TYPE": "2",
                "createTimeStart": "2025-01-01",
                "pageName": "taskListIndex",
            }
            response = requests.get(url=self.base_url, headers=self.headers, params=params)

            # 根据 queryType 保存到不同的文件
            if query_type == 1:
                file_name = self.file_name1
            elif query_type == 2:
                file_name = self.file_name2

            with open(file_name, "wb") as f:
                f.write(response.content)
            print(f"Excel 文件已保存为 {file_name}")

    def merge_excel(self):
        # 读取两个 Excel 文件
        df1 = pd.read_excel(self.file_name1, dtype={"站址编码": str})
        df2 = pd.read_excel(self.file_name2, dtype={"站址编码": str})

        # 合并两个 DataFrame
        merged_df = pd.concat([df1, df2], ignore_index=True)
        # 保存合并后的文件
        merged_df.to_excel(self.file_name3, index=False)
        print(f"合并后的 Excel 文件已保存为 {self.file_name3}")

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
            workbook_data = xl.Workbooks.Open(self.file_name3)
            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('工单信息')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:AH{last_row}')  # 从A2开始复制

            # 1. 只清除A-AH列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"A2:AH{last_clear_row}").ClearContents()

            # 只粘贴值到目标表
            source_range.Copy()
            sheet_main.Range('A2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False  # 释放剪切板

            # 获取模板的格式行（通常是第二行，假设表头下一行是格式样本）
            format_row = 2  # 模板中的格式行
            format_range = sheet_main.Range(f"A{format_row}:F{format_row}")

            # 获取新数据的范围（从第2行到最后一行）
            new_data_range = sheet_main.Range(f"A2:F{last_row}")

            # 应用模板格式到新数据
            format_range.Copy()
            new_data_range.PasteSpecial(Paste=win32.constants.xlPasteFormats)
            xl.CutCopyMode = False  # 释放剪切板

            # 使用autofill续写AI、AJ、AK、AL数据
            sheet_main.Range('AI2').AutoFill(sheet_main.Range(f'AI2:AI{last_row}'), win32.constants.xlFillDefault)
            sheet_main.Range('AJ2').AutoFill(sheet_main.Range(f'AJ2:AJ{last_row}'), win32.constants.xlFillDefault)
            sheet_main.Range('AK2').AutoFill(sheet_main.Range(f'AK2:AK{last_row}'), win32.constants.xlFillDefault)
            sheet_main.Range('AL2').AutoFill(sheet_main.Range(f'AL2:AL{last_row}'), win32.constants.xlFillDefault)

            last_data_row = sheet_main.Cells(sheet_main.Rows.Count, "AD").End(win32.constants.xlUp).Row
            if last_data_row > 1:
                # 选择AD列数据范围
                ad_column_range = sheet_main.Range(f"AD2:AD{last_data_row}")

                # 执行固定宽度分列操作（字符宽度20，实际不分列）
                ad_column_range.TextToColumns(
                    Destination=ad_column_range,  # 目标位置保持原列
                    DataType=win32.constants.xlFixedWidth,  # 固定宽度分列
                    FieldInfo=[(20, 1)],  # 设置宽度为20，列类型为常规(1)
                    DecimalSeparator=".",
                    TrailingMinusNumbers=True
                )
            workbook_data.Close(SaveChanges=False)
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
        self.merge_excel()
        self.excel_process()
if __name__ == "__main__":
    Custom_Workorder().main()