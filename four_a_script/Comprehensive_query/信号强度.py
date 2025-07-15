import datetime
from datetime import datetime,timedelta
import pythoncom
import win32com.client as win32
import time
import os

class Signal_strength():
    def __init__(self):
        # 路径：运行监控-性能查询-[监控点：信号强度]-查询-导出
        self.SPIDER_PATH = r'F:\newtowerV2\websource\spider_download'
        self.down_name1 = '信号强度.xlsx'
        self.down_name_en1 = 'performence_xinhao'
        self.model_name = '模板.xlsx'
        self.output_name1 = '信号强度_结果.xlsx'
        self.folder_temp1 = os.path.join(self.SPIDER_PATH, self.down_name_en1, 'temp')
        self.concat_name1 = os.path.join(self.SPIDER_PATH, self.down_name_en1)
        self.model_path = os.path.join(self.SPIDER_PATH, self.down_name_en1, self.model_name)
        self.file_name1 = os.path.join(self.SPIDER_PATH, self.down_name_en1, self.down_name1)
        self.output_path = os.path.join(self.SPIDER_PATH, self.down_name_en1, self.output_name1)
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
            xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
            xl.Visible = True
            xl.DisplayAlerts = False

            # 打开模板文件
            workbook_main = xl.Workbooks.Open(self.model_path)

            # === 处理已起租去重(站址编码+运营商) ===
            workbook_data = xl.Workbooks.Open(self.file_name1)
            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('清单')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:Q{last_row}')

            # 1. 只清除E2-U列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"E2:U{last_clear_row}").ClearContents()

            # 2. 复制新数据到 E-U 列
            source_range.Copy()
            sheet_main.Range('E2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False
            xl.CalculateFull()

            workbook_data.Close(SaveChanges=False)
            # 3. 使用AutoFill填充AJ-AN列的数据到last_row
            if last_row > 2:  # 确保有足够的行进行填充
                # 获取A-D列的第二行数据范围
                fill_source = sheet_main.Range('A2:D2')
                # 确定目标范围（A-D列从第2行到last_row行）
                fill_target = sheet_main.Range(f'A2:D{last_row}')
                # 使用AutoFill方法填充数据
                fill_source.AutoFill(Destination=fill_target, Type=win32.constants.xlFillDefault)
            time.sleep(3)

            # 4. 对R列实测值做分列操作，但不实际分列
            # 设置分列的参数（固定宽度和列数据类型为常规）
            last_row_main = sheet_main.Cells(sheet_main.Rows.Count, 1).End(win32.constants.xlUp).Row
            column_r_range = sheet_main.Range(f'R2:R{last_row_main}')
            column_r_range.TextToColumns(Destination=sheet_main.Range(f'R2'), DataType=win32.constants.xlFixedWidth, FieldInfo=[(1, win32.constants.xlGeneralFormat)])

            workbook_main.SaveAs(self.output_path)
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
        files = self.file_name1

        if not os.path.exists(files):
            print(f"文件不存在: {files}")
            return False

        mod_time = datetime.fromtimestamp(os.path.getmtime(files))
        if mod_time < self.start_time:
            print(f"文件 {os.path.basename(files)} 修改时间 {mod_time} 早于开始时间 {self.start_time}")
            return False

        print("所有文件都已更新，可以开始处理")
        return True

    def main(self):
        print(f"程序开始时间: {self.start_time}")
        time.sleep(20 * 60)
        while True:
            if self.check_files_modified_after_time():
                self.excel_process()
                break
            else:
                print("等待文件更新...300秒后再次检查")
                time.sleep(300)

if __name__ == "__main__":
    Signal_strength().main()
    # Signal_strength().excel_process()