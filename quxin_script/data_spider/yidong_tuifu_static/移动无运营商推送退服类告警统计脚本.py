INDEX=r'F:\untitled'
import sys
sys.path.append(INDEX)
sys.path.append(f'{INDEX}\venv\Lib\site-packages')
import os
from quxin_script.data_spider.update_pickle import update_pickle
import requests
import pythoncom
from datetime import datetime, timedelta
import pickle
import win32com.client as win32


today = datetime.now().strftime('%m-%d')
pickle_path = os.path.join(INDEX,'quxin_script\data_spider\pickle_quxin.pkl')
down_dir=os.path.join(INDEX,'quxin_script\data_spider\yidong_tuifu_static/xls')
down_path=os.path.join(down_dir,'发电工单.xlsx')
model_path=os.path.join(down_dir,'移动无运营商推送退服类告警统计-模板.xlsx')
output_dir=os.path.join(INDEX,'quxin_script\data_spider\yidong_tuifu_static/output')
output_path=os.path.join(output_dir,f'移动无运营商推送退服类告警统计({today}).xlsx')

# 下载文件
def down_core(begin, end, session):
    url = 'http://clound.gxtower.cn:11080/tower_manage_bms/a/tower/oil/report/exportOperatorReport'
    data = {
        "pageNo": "1",
        "pageSize": "25",
        "orderBy": "",
        "city.id": "",
        "city.name": "",
        "area.id": "20011335,1129,1130,1131,1132,1133,1134,1135,1136,1137,1138,1139,1140,1828749314,1142,1143,1144,1145,1146,1147,1148,1149,1150,1151,1153,1154,1155,1156,1157,1158,1159,1160,1161,1162,1163,1164,1165,1166,1167,1168,1169,1170,1171,1172,1173,1174,1175,1176,1177,1179,1180,1181,1182,1183,1184,10010113,1185,1186,1187,1188,1189,1190,1191,1192,1193,1196,1197,1198,1199,1200,1201,1202,1203,1204,1205,1206,1207,1208,1209,1210,1211,1212,1213,1214,1215,1216,1217,1219,1220,1221,1222,1223,1224,1225,1226,1227,1228,1229,1230,1231,1232,1233,1234,1235,1236,1237,1238,1239,1240,1241,1242",
        "area.name": "",
        "stationCode": "",
        "stationName": "",
        "asOper": "101",
        "number": "",
        "generatePowerState": "",
        "finishConfigId": "",
        "approvalOfDispatchId": "",
        "generateOfficeName": "",
        "collectorCode": "",
        "beginGenerateDate": begin,
        "endGenerateDate": end,
        "stopBegeinDate": "",
        "stopEndDate": ""
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

    # 从pickle获取session
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

def excel_process(index_path):
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
        # 打开模板文件
        xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
        xl.Visible = True  # 窗口是否可见
        workbook_main = xl.Workbooks.Open(model_path)  # 打开上述路径文件

        # 打开下载文件
        workbook_data = xl.Workbooks.Open(down_path)
        sheet_data = workbook_data.Sheets('0')
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

        # 将Y列的文本日期转换为日期格式
        last_row_main = sheet_main.Cells(sheet_main.Rows.Count, 1).End(win32.constants.xlUp).Row
        date_range = sheet_main.Range(f'Y2:Y{last_row_main}')  # 假设标题行在第1行，数据从第2行开始
        date_range.NumberFormat = 'yyyy-mm-dd'  # 设置日期格式
        date_range.Value = date_range.Value  # 强制Excel重新解析单元格内容

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
        prefix = '移动无运营商推送退服类告警统计'
        # 删除与保存文件名相似的所有文件
        delete_similar_files(output_dir, prefix)
        excel_process(down_path)
    except Exception as e:
        print(f"程序运行时出错: {e}")

if __name__ == "__main__":
    # 运行主程序
    main()