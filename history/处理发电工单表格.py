import win32com.client as win32
import os

# 处理Excel文件
def excel_process(index_path):
    """
    处理Excel文件，将指定文件夹中的数据文件内容复制到主表文件中。

    :param index_path: 文件夹路径
    """
    print('1、把数据文件和通报模板放在同一文件夹下')
    print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
    index_path = index_path.replace('\\', '/')  # 确保路径分隔符统一

    # 打开主表文件
    try:
        xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
        xl.Visible = True  # 窗口是否可见
        main_file = os.path.join(index_path, '移动无运营商推送退服类告警统计-模板.xlsx')  # 要处理的文件路径
        if not os.path.exists(main_file):
            raise FileNotFoundError(f"主表文件不存在: {main_file}")
        workbook_main = xl.Workbooks.Open(main_file)  # 打开上述路径文件

        # 故障工单
        for file_path in os.listdir(index_path):
            if '发电工单' in file_path:
                data_file = os.path.join(index_path, file_path)
                workbook_data = xl.Workbooks.Open(data_file)
                sheet_data = workbook_data.Sheets('Export')
                sheet_main = workbook_main.Sheets('Export')

                # 动态获取数据的实际范围
                last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
                source_range = sheet_data.Range(f'A2:DF{last_row}')  # 从A2开始复制

                # 清空目标表的内容
                sheet_main.Cells.ClearContents()

                # 复制和粘贴
                source_range.Copy()
                sheet_main.Range('A1').PasteSpecial(Paste=win32.constants.xlPasteAll)  # 使用全部的形式粘贴，保留格式
                xl.CutCopyMode = False  # 释放剪切板
                workbook_data.Close(SaveChanges=False)

        # 保存并关闭
        aa = datetime.today().strftime("%m-%d")
        workbook_main.SaveAs(os.path.join(index_path, f'移动无运营商推送退服类告警统计({aa}).xlsx'))
        workbook_main.Close()
        xl.Quit()  # 关闭Excel应用程序
        print('已全部完成')
    except Exception as e:
        print(f"处理 Excel 时出错: {e}")
    finally:
        input("程序运行完成，按回车结束")

# 主函数
def main():
    index_path = input('请输入文件夹路径（比如E:/abc）: ')
    excel_process(index_path)
    input('已全部完成，回车退出')

if __name__ == "__main__":
    main()