import requests
import sys
from datetime import datetime
import pythoncom
import time
import os
import win32com.client as win32

class battery():
    def __init__(self):
        self.SPIDER_PATH = r'F:\newtowerV2\websource\spider_download'
        self.file_name1 = '均充电压设定值.xlsx'
        self.file_name2 = '浮充电压设定值.xlsx'
        self.file_name3 = '一级低压脱离设定值.xlsx'
        self.file_name4 = '二级低压脱离设定值.xlsx'
        self.down_name_en1 = 'performence_battery'
        self.concat_name = os.path.join(self.SPIDER_PATH, self.down_name_en1)
        self.file_path1 = os.path.join(self.SPIDER_PATH, self.down_name_en1, self.file_name1)
        self.file_path2 = os.path.join(self.SPIDER_PATH, self.down_name_en1, self.file_name2)
        self.file_path3 = os.path.join(self.SPIDER_PATH, self.down_name_en1, self.file_name3)
        self.file_path4 = os.path.join(self.SPIDER_PATH, self.down_name_en1, self.file_name4)
        self.output_name = os.path.join(self.concat_name, 'output', '开关电源电压设置修改情况通报-结果.xlsx')
        self.model_path = os.path.join(self.concat_name, '模板.xlsx')
        self.start_time = datetime.now()
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
            xl = win32.Dispatch('Excel.Application')
            xl.Visible = True
            xl.DisplayAlerts = False  # 禁用警告提示
            # 打开模板文件
            workbook_main = xl.Workbooks.Open(self.model_path)

            # === 处理均充电压设定值 ===
            workbook_data = xl.Workbooks.Open(self.file_path1)
            workbook_data.Application.DisplayAlerts = False  # 单独禁用警告
            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('均充电压清单')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:P{last_row}')

            # 1. 只清除A-P列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"A2:P{last_clear_row}").ClearContents()

            # 2. 复制新数据到 A-P 列
            source_range.Copy()
            sheet_main.Range('A2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False
            xl.CalculateFull()
            workbook_data.Close(SaveChanges=False)
            # 3. 对N列执行分列操作（固定宽度，常规格式，但不实际分列）
            last_data_row = sheet_main.Cells(sheet_main.Rows.Count, "N").End(win32.constants.xlUp).Row
            if last_data_row > 1:
                # 选择N列数据范围
                n_column_range = sheet_main.Range(f"N2:N{last_data_row}")

                # 执行分列操作但不实际分列
                n_column_range.TextToColumns(
                    Destination=n_column_range,
                    DataType=win32.constants.xlFixedWidth,
                    FieldInfo=[(0, 1)],  # 只设置一个字段
                    DecimalSeparator=".",
                    TrailingMinusNumbers=True
                )
                time.sleep(3)

            # === 处理浮充电压设定值 ===
            workbook_data = xl.Workbooks.Open(self.file_path2)
            workbook_data.Application.DisplayAlerts = False  # 单独禁用警告
            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('浮充电压清单')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:P{last_row}')

            # 1. 只清除A-P列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"A2:P{last_clear_row}").ClearContents()

            # 2. 复制新数据到 A-P 列
            source_range.Copy()
            sheet_main.Range('A2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False
            # 重新计算工作表以确保公式更新
            xl.CalculateFull()
            workbook_data.Close(SaveChanges=False)
            # 3. 对N列执行分列操作（固定宽度，常规格式，但不实际分列）
            last_data_row = sheet_main.Cells(sheet_main.Rows.Count, "N").End(win32.constants.xlUp).Row
            if last_data_row > 1:
                # 选择N列数据范围
                n_column_range = sheet_main.Range(f"N2:N{last_data_row}")

                # 执行分列操作但不实际分列
                n_column_range.TextToColumns(
                    Destination=n_column_range,
                    DataType=win32.constants.xlFixedWidth,
                    FieldInfo=[(0, 1)],  # 只设置一个字段
                    DecimalSeparator=".",
                    TrailingMinusNumbers=True
                )
            time.sleep(3)

            # === 处理一级低压脱离设定值 ===
            workbook_data = xl.Workbooks.Open(self.file_path3)
            workbook_data.Application.DisplayAlerts = False

            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('一级低压脱离设定值清单')
            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A1:P{last_row}')  # 从A2开始复制
            # 清空目标表的内容
            sheet_main.Cells.ClearContents()
            # 复制和粘贴
            source_range.Copy()
            sheet_main.Range('A1').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False  # 释放剪切板
            workbook_data.Close(SaveChanges=False)
            # 3. 对N列执行分列操作（固定宽度，常规格式，但不实际分列）
            last_data_row = sheet_main.Cells(sheet_main.Rows.Count, "N").End(win32.constants.xlUp).Row
            if last_data_row > 1:
                # 选择N列数据范围
                n_column_range = sheet_main.Range(f"N2:N{last_data_row}")

                # 执行分列操作但不实际分列
                n_column_range.TextToColumns(
                    Destination=n_column_range,
                    DataType=win32.constants.xlFixedWidth,
                    FieldInfo=[(0, 1)],  # 只设置一个字段
                    DecimalSeparator=".",
                    TrailingMinusNumbers=True
                )
            time.sleep(3)

            # 处理第四个数据文件（二级低压脱离设定值）
            workbook_data = xl.Workbooks.Open(self.file_path4)
            workbook_data.Application.DisplayAlerts = False  # 单独禁用警告

            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('二级低压脱离设定值清单')
            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A1:P{last_row}')  # 从A2开始复制
            # 清空目标表的内容
            sheet_main.Cells.ClearContents()
            # 复制和粘贴
            source_range.Copy()
            sheet_main.Range('A1').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False  # 释放剪切板
            workbook_data.Close(SaveChanges=False)
            # 3. 对N列执行分列操作（固定宽度，常规格式，但不实际分列）
            last_data_row = sheet_main.Cells(sheet_main.Rows.Count, "N").End(win32.constants.xlUp).Row
            if last_data_row > 1:
                # 选择N列数据范围
                n_column_range = sheet_main.Range(f"N2:N{last_data_row}")

                # 执行分列操作但不实际分列
                n_column_range.TextToColumns(
                    Destination=n_column_range,
                    DataType=win32.constants.xlFixedWidth,
                    FieldInfo=[(0, 1)],  # 只设置一个字段
                    DecimalSeparator=".",
                    TrailingMinusNumbers=True
                )
            workbook_main.SaveAs(self.output_name)
            workbook_main.Close()
            xl.Quit()  # 关闭Excel应用程序
            print('已全部完成')
        except Exception as e:
            raise
        finally:
            # 释放 COM 库
            pythoncom.CoUninitialize()

    def check_files_modified_after_time(self):
        """
        检查四个文件是否都在开始时间之后被修改过

        :param start_time: 开始时间(datetime对象)
        :param config: 配置对象
        :return: 如果所有文件都在开始时间之后被修改过则返回True，否则返回False
        """
        files = [
            self.file_name1,
            self.file_name2,
            self.file_name3,
            self.file_name4
        ]

        for file in files:
            if not os.path.exists(file):
                print(f"文件不存在: {file}")
                return False

            mod_time = datetime.fromtimestamp(os.path.getmtime(file))
            if mod_time < self.start_time:
                print(f"文件 {os.path.basename(file)} 修改时间 {mod_time} 早于开始时间 {self.start_time}")
                return False

        print("所有文件都已更新，可以开始处理")
        return True

    def main(self):
        print(f"程序开始时间: {self.start_time}")
        time.sleep(30 * 60)
        while True:
            if self.check_files_modified_after_time():
                self.excel_process()
                break
            else:
                print("等待文件更新...300秒后再次检查")
                time.sleep(300)

if __name__ == "__main__":
    battery().main()
