import os
import requests
import sys
from datetime import datetime, timedelta
import pickle
import win32com.client as win32

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
    pickle_path = os.path.dirname(os.path.abspath('pickle_quxin.pkl'))
    file_path = os.path.join(pickle_path, '发电工单.xlsx')
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
    try:
        index_path = down()
        excel_process(index_path)
        print('程序已自动退出')
    except Exception as e:
        print(f"程序运行时出错: {e}")

if __name__ == "__main__":
    # 运行主程序
    main()