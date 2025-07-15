import os
import win32com.client as win32
import datetime

def process_excel_files(index_path):
    """
    处理Excel文件，将指定文件夹中的数据文件内容复制到主表文件中。

    :param index_path: 文件夹路径
    """
    print('1、把数据文件和通报模板放在同一文件夹下')
    print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
    index_path = index_path.replace('\\', '/')  # 确保路径分隔符统一
    print(f"Index path: {index_path}")

    # 打开主表文件
    xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
    xl.Visible = True  # 窗口是否可见

    main_file = os.path.join(index_path, '工单及油机匹配率统计-模板.xlsx') # 要处理的文件路径
    print(f"Main file path: {main_file}")
    workbook_main = xl.Workbooks.Open(main_file)  # 打开上述路径文件

    # 故障工单
    for file_path in os.listdir(index_path):
        if '工单及油机匹配率统计表' in file_path:
            data_file = os.path.join(index_path, file_path)
            print(f"Data file path: {data_file}")
            workbook_data = xl.Workbooks.Open(data_file)
            sheet_data = workbook_data.Sheets('Export')
            sheet_main = workbook_main.Sheets('Export')
            sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
            source_range = sheet_data.Range('A3:J17')
            target_range = sheet_main.Range('A3:J17')
            source_range.Copy()
            target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
            xl.CutCopyMode = False  # 释放剪切板
            workbook_data.Close(SaveChanges=False)
    today = datetime.today().strftime('%m-%d')
    workbook_main.SaveAs(os.path.join(index_path, f'工单及油机匹配率统计({today}).xlsx'))
    workbook_main.Close()
    xl.Quit()  # 关闭Excel应用程序
    print('已全部完成')