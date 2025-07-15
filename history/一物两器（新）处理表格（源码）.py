import time
import os
from datetime import datetime
import win32com.client as win32

print('1、把数据文件和通报模板放在同一文件夹下')
print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
INDEX = input('3、输入该文件夹路径，直接在地址栏复制即可（比如E:/abc）:') + '/'
INDEX = INDEX.replace('\\', '/')
# # 打开主表文件
xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
xl.Visible = True  # 窗口是否可见
main_file = INDEX + r'模板.xlsx'  # 要处理的文件路径
workbook_main = xl.Workbooks.Open(main_file)  # 打开上述路径文件、

# 先把通报数据更新到昨日数据
sheet_data = workbook_main.Sheets('通报')
sheet_main = workbook_main.Sheets('昨日数据')
sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
source_range = sheet_data.Range('B1:AL19')
target_range = sheet_main.Range('A1:AK19')
source_range.Copy()
target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴

# # 把各类设备比例更新到昨日数据
sheet_data = workbook_main.Sheets('各类设备比例')
sheet_main = workbook_main.Sheets('昨日数据')
sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
source_range = sheet_data.Range('A1:AF60')
target_range = sheet_main.Range('A23:AF82')
source_range.Copy()
target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴


# # 设备覆盖率统计表+FSU覆盖
# # 设备覆盖
for file_path in os.listdir(INDEX):
   if '设备覆盖统计表' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet1')
        sheet_main = workbook_main.Sheets('设备覆盖率统计表+FSU覆盖+微站监控覆盖')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1:AD18')
        target_range = sheet_main.Range('C1:AF18')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)

#
# # FSU覆盖
for file_path in os.listdir(INDEX):
   if 'FSU覆盖稽核后统计表' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet1')
        sheet_main = workbook_main.Sheets('设备覆盖率统计表+FSU覆盖+微站监控覆盖')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1:AZ17')
        target_range = sheet_main.Range('A24:AZ40')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 微站监控覆盖
for file_path in os.listdir(INDEX):
   if 'FSU覆盖稽核后统计表' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet1')
        sheet_main = workbook_main.Sheets('设备覆盖率统计表+FSU覆盖+微站监控覆盖')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1:AZ17')
        target_range = sheet_main.Range('A24:AZ40')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)



# 设备未覆盖清单+FSU未覆盖清单
# 设备未覆盖
for file_path in os.listdir(INDEX):
    if '设备覆盖明细' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet0')
        sheet_main = workbook_main.Sheets('设备覆盖清单+FSU未覆盖清单')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        column_J = sheet_data.Range('J:J')
        column_J.AutoFilter(1, "否")
        source_range = sheet_data.Range('A:J')
        target_range = sheet_main.Range('A:J')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# FSU未覆盖
for file_path in os.listdir(INDEX):
    if 'FSU覆盖明细表' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet0')
        sheet_main = workbook_main.Sheets('设备覆盖清单+FSU未覆盖清单')

        # 设置筛选条件，筛选列I（第9列）为'否'
        sheet_data.Range('A:K').AutoFilter(Field=9, Criteria1='否')
        # 然后筛选列J（第10列）也为'否'
        sheet_data.Range('A:K').AutoFilter(Field=10, Criteria1='否')

        source_range = sheet_data.Range('A:k')
        target_range = sheet_main.Range('N:X')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 设备离线清单
for file_path in os.listdir(INDEX):
   if '设备离线异常清单' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet0')
        sheet_main = workbook_main.Sheets('设备离线清单')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A:P')
        target_range = sheet_main.Range('A:P')
        source_range.Copy()
        target_range.PasteSpecial()  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 设备不准确清单
for file_path in os.listdir(INDEX):
   if '设备不准确异常清单' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet0')
        sheet_main = workbook_main.Sheets('设备不准确清单')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A:R')
        target_range = sheet_main.Range('A:R')
        source_range.Copy()
        target_range.PasteSpecial()  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 设备离线率统计表
for file_path in os.listdir(INDEX):
   if '综合设备离线率统计表' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet1')
        sheet_main = workbook_main.Sheets('设备离线率统计表')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1:AD17')
        target_range = sheet_main.Range('C1:AF17')
        source_range.Copy()
        target_range.PasteSpecial()  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 设备准确率统计表
for file_path in os.listdir(INDEX):
   if '机构FSU准确率统计表' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('sheet1')
        sheet_main = workbook_main.Sheets('设备准确率统计表')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A1:CQ17')
        target_range = sheet_main.Range('C1:CS17')
        source_range.Copy()
        target_range.PasteSpecial()  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 离线率清单（含超长）
for file_path in os.listdir(INDEX):
    if 'FSU离线率_日' in file_path:
         data_file = os.path.join(INDEX, file_path)
         workbook_data = xl.Workbooks.Open(data_file)
         sheet_data = workbook_data.Sheets('FSU离线率_日')
         sheet_main = workbook_main.Sheets('离线率清单（含超长）')
         sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
         source_range = sheet_data.Range('C3:J17')
         target_range = sheet_main.Range('V4:AC18')
         source_range.Copy()
         target_range.PasteSpecial()  # 使用值的形式粘贴
         xl.CutCopyMode = False  # 释放剪切板
         time.sleep(2)
         workbook_data.Close(SaveChanges=False)

for file_path in os.listdir(INDEX):
    if '超长超频清单' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data_cc = workbook_data.Sheets('超长清单')
        sheet_main = workbook_main.Sheets('离线率清单（含超长）')
        sheet_data_cc.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        today = datetime.now().strftime('%Y-%m-%d')
        column_Q = sheet_data_cc.Range('Q:Q')
        column_Q.AutoFilter(1, today)
        source_range = sheet_data_cc.Range('A:Q')
        target_range = sheet_main.Range('A:Q')
        source_range.Copy()
        target_range.PasteSpecial()

# 超频清单
        sheet_data_cp = workbook_data.Sheets('超频清单')
        sheet_main = workbook_main.Sheets('超频清单')
        sheet_data_cp.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        today = datetime.now().strftime('%Y-%m-%d')
        column_K = sheet_data_cp.Range('K:K')
        column_K.AutoFilter(1, today)
        source_range = sheet_data_cp.Range('A:K')
        target_range = sheet_main.Range('A:K')
        source_range.Copy()
        target_range.PasteSpecial()

        sheet_data_cptj = workbook_data.Sheets('超频统计')
        sheet_main = workbook_main.Sheets('超长超频')
        sheet_data_cptj.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data_cptj.Range('A1:AJ17')
        target_range = sheet_main.Range('A1:AJ17')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴

        sheet_data_cctj = workbook_data.Sheets('超长统计')
        sheet_main = workbook_main.Sheets('超长超频')
        sheet_data_cctj.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data_cctj.Range('A1:AI17')
        target_range = sheet_main.Range('A21:AI37')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 转供电匹配
for file_path in os.listdir(INDEX):
    if '缴纳判断缴纳电费' in file_path:
         data_file = os.path.join(INDEX, file_path)
         workbook_data = xl.Workbooks.Open(data_file)
         sheet_data = workbook_data.Sheets('sheet0')
         sheet_main = workbook_main.Sheets('转供电匹配')
         sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
         source_range = sheet_data.Range('A:O')
         target_range = sheet_main.Range('A:O')
         source_range.Copy()
         target_range.PasteSpecial()  # 使用值的形式粘贴
         xl.CutCopyMode = False  # 释放剪切板
         time.sleep(2)
         workbook_data.Close(SaveChanges=False)

#
#
# # 分路计量及远程抄表明细表
for file_path in os.listdir(INDEX):
    if '分路计量及远程抄表' in file_path:
         data_file = os.path.join(INDEX, file_path)
         workbook_data = xl.Workbooks.Open(data_file)
         sheet_data = workbook_data.Sheets('sheet1')
         sheet_main = workbook_main.Sheets('分路计量及远程抄表明细表')
         sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
         source_range = sheet_data.Range('A:N')
         target_range = sheet_main.Range('A:N')
         source_range.Copy()
         target_range.PasteSpecial()  # 使用值的形式粘贴
         xl.CutCopyMode = False  # 释放剪切板
         time.sleep(2)
         workbook_data.Close(SaveChanges=False)


# 设备管理（前天）
for file_path in os.listdir(INDEX):
    if '前天' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('Sheet0')
        sheet_main = workbook_main.Sheets('设备管理（前天）')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全


        # 取消特定范围内的合并单元格
        range_unmerge_data = sheet_data.Range('A1')  # 根据实际需要调整范围
        range_unmerge_data.MergeCells = False
        range_unmerge_main = sheet_main.Range('A1')  # 根据实际需要调整范围
        range_unmerge_main.MergeCells = False

        # 筛选数据
        column_O = sheet_data.Range('O:O')
        column_O.AutoFilter(Field=1, Criteria1=["缴费电表", "缴费类型"], Operator=win32.constants.xlFilterValues)

        # 复制筛选后的数据
        source_range = sheet_data.Range('A:U')
        target_range = sheet_main.Range('A:U')
        source_range.Copy()
        target_range.PasteSpecial()

        range_to_unmerge = sheet_main.Range('A1:U1')  # 根据实际需要调整范围
        range_to_unmerge.MergeCells = True
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


# 设备管理（昨天）
for file_path in os.listdir(INDEX):
    if '昨天' in file_path:
        data_file = os.path.join(INDEX, file_path)
        workbook_data = xl.Workbooks.Open(data_file)
        sheet_data = workbook_data.Sheets('Sheet0')
        sheet_main = workbook_main.Sheets('设备管理（昨天）')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全

        # 取消特定范围内的合并单元格
        range_unmerge_data = sheet_data.Range('A1')  # 根据实际需要调整范围
        range_unmerge_data.MergeCells = False
        range_unmerge_main = sheet_main.Range('A1')  # 根据实际需要调整范围
        range_unmerge_main.MergeCells = False

        # 筛选数据
        column_O = sheet_data.Range('O:O')
        column_O.AutoFilter(Field=1, Criteria1=["缴费电表", "缴费类型"], Operator=win32.constants.xlFilterValues)

        # 复制筛选后的数据
        source_range = sheet_data.Range('A:U')
        target_range = sheet_main.Range('A:U')
        source_range.Copy()
        target_range.PasteSpecial()

        range_to_unmerge = sheet_main.Range('A1:U1')  # 根据实际需要调整范围
        range_to_unmerge.MergeCells = True
        xl.CutCopyMode = False  # 释放剪切板
        time.sleep(2)
        workbook_data.Close(SaveChanges=False)


workbook_main.SaveAs(INDEX + r'一物两器_更新后.xlsx')
workbook_main.Close()
input("已全部完成，回车退出")
