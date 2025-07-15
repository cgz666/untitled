INDEX=r'F:\untitled'
import sys
sys.path.append(INDEX)
sys.path.append(f'{INDEX}\venv\Lib\site-packages')
import requests
from quxin_script.data_spider.update_pickle import update_pickle
from datetime import datetime, timedelta
import pickle
import win32com.client as win32
import pythoncom
import os

today = datetime.now().strftime('%m-%d')
pickle_path = os.path.join(INDEX,'quxin_script\data_spider\pickle_quxin.pkl')
down_dir=os.path.join(INDEX,'quxin_script\data_spider\gongdan_youji_pipei/xls')
down_path=os.path.join(down_dir,'油机匹配表.xlsx')
model_path=os.path.join(down_dir,'工单及油机匹配率统计-模板.xlsx')
output_dir=os.path.join(INDEX,'quxin_script\data_spider\gongdan_youji_pipei/output')
output_path=os.path.join(output_dir,f'工单及油机匹配率统计({today}).xlsx')

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
    with open(down_path, "wb") as file:
        file.write(res.content)
    return pickle_path
def down():
    # 获取本月一号和今天的前一天的日期
    today = datetime.today()

    if today.day == 1:
        # 如果今天是本月的一号
        begin_generate_date = today.strftime("%Y-%m-%d 00:00:00")
        end_generate_date = today.strftime("%Y-%m-%d 23:59:59")
    else:
        # 如果今天不是一号
        first_day_of_month = today.replace(day=1)
        yesterday = today - timedelta(days=1)

        # 格式化日期为字符串
        begin_generate_date = first_day_of_month.strftime("%Y-%m-%d 00:00:00")
        end_generate_date = yesterday.strftime("%Y-%m-%d 23:59:59")


    with open(pickle_path, 'rb') as f:
        session = pickle.load(f)
    return down_core(begin_generate_date, end_generate_date, session)

def delete_similar_files(output_dir, prefix):
    """
    删除指定目录下所有以指定前缀开头的文件
    """
    if not os.path.exists(output_dir):
        return
    for file_name in os.listdir(output_dir):
        if file_name.startswith(prefix):
            file_path = os.path.join(output_dir, file_name)
            os.remove(file_path)

def process_excel_files(index_path):
    """
    处理Excel文件，将指定文件夹中的数据文件内容复制到主表文件中。

    :param index_path: 文件夹路径
    """
    print('1、把数据文件和通报模板放在同一文件夹下')
    print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
    index_path = index_path.replace('\\', '/')  # 确保路径分隔符统一

    # 初始化 COM 库
    pythoncom.CoInitialize()

    try:
        # 打开主表文件
        xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
        xl.Visible = True  # 窗口是否可见
        workbook_main = xl.Workbooks.Open(model_path)  # 打开上述路径文件

        # 故障工单
        workbook_data = xl.Workbooks.Open(down_path)
        sheet_data = workbook_data.Sheets('Export')
        sheet_main = workbook_main.Sheets('Export')
        sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
        source_range = sheet_data.Range('A3:J17')
        target_range = sheet_main.Range('A3:J17')
        source_range.Copy()
        target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
        xl.CutCopyMode = False  # 释放剪切板
        workbook_data.Close(SaveChanges=False)

        workbook_main.SaveAs(output_path)
        workbook_main.Close()
        xl.Quit()  # 关闭Excel应用程序
        print('已全部完成')
    except Exception as e:
        raise
    finally:
        # 释放 COM 库
        pythoncom.CoUninitialize()

# 主函数
def main():
    try:
        down()
        prefix = '工单及油机匹配率统计'
        # 删除与保存文件名相似的所有文件
        delete_similar_files(output_dir, prefix)
        process_excel_files(down_path)
    except Exception as e:
        print(f"程序运行时出错: {e}")

if __name__ == "__main__":
    # 运行主程序
    main()