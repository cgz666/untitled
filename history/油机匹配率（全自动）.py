import os
import requests
import sys
from datetime import datetime, timedelta
import pickle
import win32com.client as win32


# 下载文件
def down_core(begin, end, session):
    url = 'http://clound.gxtower.cn:11080/tower_manage_bms/a/tower/oil/towerOilStatistics/exportOilAndGenerateMatch'
    data = {
            "asOper": "",
            "beginGenerateDate": begin,
            "endGenerateDate": end
    }
    res = session.post(url=url, data=data)
    # 获取pickle文件路径
    pickle_path = os.path.dirname(os.path.abspath('pickle_quxin.pkl'))
    file_path = os.path.join(pickle_path, '油机匹配表.xlsx')
    with open(file_path, "wb") as file:
        file.write(res.content)
    print(f"文件已保存到: {file_path}")
    return pickle_path

def down():
    # 初始化
    pickle_quxin = r'pickle_quxin.pkl'
    # 获取本月一号和今天的前一天的日期
    today = datetime.today()
    first_day_of_month = today.replace(day=1)
    yesterday = today - timedelta(days=1)

    # 格式化日期为字符串
    begin_generate_date = first_day_of_month.strftime("%Y-%m-%d 00:00:00")
    end_generate_date = yesterday.strftime("%Y-%m-%d 23:59:59")

    # 下载pickle
    res = requests.get('http://clound.gxtower.cn:3980/tt/get_session_quxin')
    with open(pickle_quxin, "wb") as file:
        file.write(res.content)
    # 从pickle获取session
    with open(pickle_quxin, 'rb') as f:
        session = pickle.load(f)
    return down_core(begin_generate_date, end_generate_date, session)

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
# 主函数
def main():
    try:
        index_path = down()
        process_excel_files(index_path)
    except Exception as e:
        print(f"程序运行时出错: {e}")
if __name__ == "__main__":
    # 运行主程序
    main()