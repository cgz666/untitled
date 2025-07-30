import os

import win32com.client as win32

print('1、把数据文件和通报模板放在同一文件夹下')
print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
INDEX = input('3、输入该文件夹路径，直接在地址栏复制即可（比如E:/abc）:') + '/'
INDEX = INDEX.replace('\\', '/')
# # 打开主表文件
xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
xl.Visible = True  # 窗口是否可见
main_file = INDEX + r'模板.xlsx'  # 要处理的文件路径
workbook_main = xl.Workbooks.Open(main_file)  # 打开上述路径文件

# # 故障工单
for file_path in os.listdir(INDEX):
   if '故障工单' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('统计表')
        sheet_main = workbook_main.Sheets('1.2.3故障工单急全量工单质检')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('C2:H16')
        target_range = sheet_main.Range('C2:H16')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# # 巡检工单
for file_path in os.listdir(INDEX):
    if '巡检工单' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('统计通报')
        sheet_main = workbook_main.Sheets('4.5巡检工单质检')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('C3:AF17')
        target_range = sheet_main.Range('D3:AG17')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# # 一物两器
for file_path in os.listdir(INDEX):
    if '一物两器' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('通报')
        sheet_main = workbook_main.Sheets('14--24一物两器')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        # # 隐藏列显隐
        # column = sheet_data.Columns('D')
        # column.Hidden = False

        source_range = sheet_data.Range('B1:AL19')
        target_range = sheet_main.Range('A1:AK19')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# # 四率
for file_path in os.listdir(INDEX):
    if '四率' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('通报')
        sheet_main = workbook_main.Sheets('25.运营监控工单通报')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('B1:U22')
        target_range = sheet_main.Range('B1:U22')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)
#
# # 运营商退服
for file_path in os.listdir(INDEX):
    if '电信' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('运行报表_日')
        sheet_main = workbook_main.Sheets('27.运营商退服时长')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('C4:R18')
        target_range = sheet_main.Range('AO5:BD19')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# for file_path in os.listdir(INDEX):
    if '移动' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('运行报表_日')
        sheet_main = workbook_main.Sheets('27.运营商退服时长')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('C4:R18')
        target_range = sheet_main.Range('C5:R19')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# for file_path in os.listdir(INDEX):
    if '联通' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('运行报表_日')
        sheet_main = workbook_main.Sheets('27.运营商退服时长')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('C4:R18')
        target_range = sheet_main.Range('V5:AK19')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值...
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# # 超长超频
for file_path in os.listdir(INDEX):
    if '超长超频' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data_cp = workbook_data.Sheets('超频清单')
        sheet_main = workbook_main.Sheets('26.超频退服')
        sheet_data_cp.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        last_row = sheet_main.UsedRange.Rows.Count
        sheet_main.Range(f'A2:H{last_row}').ClearContents()
        last_row_data = sheet_data_cp.UsedRange.Rows.Count
        source_range = sheet_data_cp.Range(f'A2:H{last_row_data}')
        target_range = sheet_main.Range(f'A2:H{last_row_data}')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)

        sheet_data_cc = workbook_data.Sheets('超长退服明细')
        sheet_main = workbook_main.Sheets('26.超长退服')
        sheet_data_cc.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        last_row = sheet_main.UsedRange.Rows.Count
        # sheet_main.Range(f'A2:AG{last_row}').Copy()
        sheet_main.Range(f'A2:AG{last_row}').ClearContents()
        last_row_data = sheet_data_cc.UsedRange.Rows.Count
        source_range = sheet_data_cc.Range(f'A2:AM{last_row_data}')
        target_range = sheet_main.Range(f'A2:AM{last_row_data}')
        source_range.Copy()
        target_range.PasteSpecial()
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# # 疑似退服
# for file_path in os.listdir(INDEX):
#     if '疑似退服' in file_path:
#         data_file = os.path.join(INDEX, file_path)
#
#         workbook_data_ystf = xl.Workbooks.Open(data_file)
#         sheet_data = workbook_data_ystf.Sheets('退服+疑似退服明细')
#         sheet_main = workbook_main.Sheets('22、23疑似退服指标')
#         last_row = sheet_main.UsedRange.Rows.Count
#         sheet_main.Range(f'A2:AA{last_row}').ClearContents()
#         last_row_data = sheet_data.UsedRange.Rows.Count
#         source_range = sheet_data.Range(f'A2:AA{last_row_data}')
#         target_range = sheet_main.Range(f'A2:AA{last_row_data}')
#         source_range.Copy()
#         target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)
#
#         xl.CutCopyMode = False  # 释放剪切板
#         workbook_data_ystf.Close(SaveChanges=False)
#
# # 能源工单
for file_path in os.listdir(INDEX):
    if '能源维护整体' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('通报数据')
        sheet_main = workbook_main.Sheets('6.7.8能源工单质检和故障处理')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1:J17')
        target_range = sheet_main.Range('A1:J17')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)


# 低速充电桩维护
for file_path in os.listdir(INDEX):
    if '低速充电桩' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('通报')
        sheet_main = workbook_main.Sheets('9.10低速充电桩维护')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        # 粘贴第一张表
        source_range = sheet_data.Range('A1:X20')
        target_range = sheet_main.Range('A1:X20')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板

        # 粘贴第二张表
        source_range = sheet_data.Range('A47:V66')
        target_range = sheet_main.Range('A24:V43')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)


# 能源工单及换电维护指标
for file_path in os.listdir(INDEX):
    if '能源工单及换电维护指标' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('通报数据')
        sheet_main = workbook_main.Sheets('11.能源工单及换电维护指标')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1: AZ19')
        target_range = sheet_main.Range('A1: AZ19')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

# 包站人
for file_path in os.listdir(INDEX):
    if '拓展包站人配置率' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('拓展包站人配置率')
        sheet_main = workbook_main.Sheets('12.13包站人')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        # 粘贴第一张表
        source_range = sheet_data.Range('A1:D18')
        target_range = sheet_main.Range('A1:D18')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板


    if '能源包站人配置率' in file_path:
        # 粘贴第二张表
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('能源包站人配置率')
        sheet_main = workbook_main.Sheets('12.13包站人')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1:T19')
        target_range = sheet_main.Range('G1:Z19')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)



workbook_main.SaveAs(INDEX + r'网络指标_更新后.xlsx')
workbook_main.Close()

input('已全部完成，回车退出')