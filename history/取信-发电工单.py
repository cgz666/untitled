import os
import pickle
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32

# 配置模块
INDEX = 'F:\\Python313\\untitled2\\test\\取信'
pickle_session = os.path.join(INDEX, 'pickle_quxin.pkl')

# 保存session到pickle文件
def save_session(session):
    with open(pickle_session, 'wb') as f:
        pickle.dump(session, f)
    print("Session已保存到pickle文件中")

# 加载session从pickle文件
def load_session():
    try:
        with open(pickle_session, 'rb') as f:
            session = pickle.load(f)
        print("加载的Session对象：", session)  # 调试信息，确认Session对象
        return session
    except (FileNotFoundError, EOFError) as e:
        print(f"Session文件不存在或无法加载：{e}")
        return None

# 提取HTML表格数据
def extract_table_data(soup):
    table = soup.find('table')  # 假设数据在第一个表格中
    if not table:
        print("未找到表格数据")
        return None, None

    rows = table.find_all('tr')
    table_headers = [th.get_text(strip=True) for th in rows[0].find_all('th')]
    table_data = []
    for row in rows[1:]:
        cells = row.find_all('td')
        row_data = [cell.get_text(strip=True).replace('\n', ' ').replace('\r', ' ') for cell in cells]
        if any(cell.strip() for cell in row_data):  # 检查行是否为空
            table_data.append(row_data)
    return table_headers, table_data


# 处理Excel文件
def process_excel_files(index_path):
    """
    处理Excel文件，将指定文件夹中的数据文件内容按表头匹配并插入到主表文件中。

    :param index_path: 文件夹路径
    """
    print('1、把数据文件和通报模板放在同一文件夹下')
    print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
    index_path = index_path.replace('\\', '/')  # 确保路径分隔符统一
    # 打开主表文件
    xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
    xl.Visible = False  # 窗口是否可见
    main_file = os.path.join(index_path, '移动无运营商推送退服类告警统计（5月7日）.xlsx')  # 要处理的文件路径
    workbook_main = xl.Workbooks.Open(main_file)  # 打开上述路径文件

    # 获取主表文件的表头
    sheet_main = workbook_main.Sheets('Export')
    main_headers = [cell.Value for cell in sheet_main.Range(sheet_main.Cells(1, 1), sheet_main.Cells(1, sheet_main.UsedRange.Columns.Count))]

    # 打开数据文件
    data_file = os.path.join(index_path, 'fadian.xlsx')  # 假设数据文件名为 fadian.xlsx
    workbook_data = xl.Workbooks.Open(data_file)
    sheet_data = workbook_data.Sheets('sheet1')
    sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全

    # 获取数据文件的表头
    data_headers = [cell.Value for cell in sheet_data.Range(sheet_data.Cells(1, 1), sheet_data.Cells(1, sheet_data.UsedRange.Columns.Count))]

    # 创建表头映射
    header_map = {header: i + 1 for i, header in enumerate(main_headers)}

    # 找到主表文件中需要插入数据的起始位置
    start_row = 2  # 假设从第2行开始插入

    # 遍历数据文件的每一行数据
    total_rows = sheet_data.UsedRange.Rows.Count - 1  # 减去表头行
    for row in range(2, sheet_data.UsedRange.Rows.Count + 1):
        print(f"正在处理行 {row - 1}/{total_rows}...")
        # 创建一个空的列表，用于存储按主表表头顺序排列的数据
        row_data = [None] * len(main_headers)
        for col in range(1, sheet_data.UsedRange.Columns.Count + 1):
            data_header = data_headers[col - 1]
            if data_header in header_map:
                target_col = header_map[data_header]
                row_data[target_col - 1] = sheet_data.Cells(row, col).Value

        # 将处理后的数据插入到主表文件中
        for col, value in enumerate(row_data, start=1):
            sheet_main.Cells(start_row, col).Value = value
        start_row += 1

    workbook_data.Close(SaveChanges=False)
    workbook_main.SaveAs(os.path.join(index_path, '网络指标_更新后.xlsx'))
    workbook_main.Close()
    xl.Quit()  # 关闭Excel应用程序
    print('已全部完成')


# 主程序
def main():
    # 加载session
    session = load_session()
    if not session:
        print("未找到Session文件，请先运行save_session()保存Session")
        return

    # 获取本月一号和今天的前一天的日期
    today = datetime.today()
    first_day_of_month = today.replace(day=1)
    yesterday = today - timedelta(days=5)

    # 格式化日期为字符串
    begin_generate_date = first_day_of_month.strftime("%Y-%m-%d 00:00:00")
    end_generate_date = yesterday.strftime("%Y-%m-%d 23:59:59")

    # 初始化一个空的列表，用于存储所有页面的数据
    all_page_data = []

    url = "http://clound.gxtower.cn:11080/tower_manage_bms/a/tower/oil/report/exportOperatorReport"

    # 发送请求获取第一页数据
    first_page_data = {
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
        "beginGenerateDate": begin_generate_date,
        "endGenerateDate": end_generate_date,
        "stopBegeinDate": "",
        "stopEndDate": ""
    }

    # 发送请求获取第一页数据
    response = session.post(url=url, data=first_page_data)
    response.raise_for_status()  # 检查请求是否成功

    # 检查响应内容类型
    if 'text/html' in response.headers.post('Content-Type', ''):
        # 解析 HTML 内容
        soup = BeautifulSoup(response.content, 'html.parser')

        # 提取表格数据
        table = soup.find('table')  # 假设数据在第一个表格中
        if table:
            rows = table.find_all('tr')
            table_headers = [th.get_text(strip=True) for th in rows[0].find_all('th')]
            table_data = []
            for row in rows[1:]:
                cells = row.find_all('td')
                row_data = [cell.get_text(strip=True).replace('\n', ' ').replace('\r', ' ') for cell in cells]  # 替换换行符
                if any(cell.strip() for cell in row_data):  # 检查行是否为空
                    table_data.append(row_data)
            all_page_data.extend(table_data)

            # 提取总页数
            pagination = soup.find('div', class_='pagination')  # 假设分页信息在某个特定的div中
            if pagination:
                page_links = pagination.find_all('a')
                total_pages = 1
                for link in page_links:
                    text = link.get_text().strip()
                    if text.isdigit():
                        total_pages = max(total_pages, int(text))
            else:
                total_pages = 1  # 如果没有分页信息，默认为 1 页

            # 循环遍历每一页
            for page in range(2, total_pages + 1):
                data = first_page_data.copy()
                data['pageNo'] = str(page)
                response = session.post(url=url, data=data)
                response.raise_for_status()  # 检查请求是否成功

                # 解析 HTML 内容
                soup = BeautifulSoup(response.content, 'html.parser')

                # 提取表格数据
                table = soup.find('table')  # 假设数据在第一个表格中
                if table:
                    rows = table.find_all('tr')
                    table_data = []
                    for row in rows[1:]:
                        cells = row.find_all('td')
                        row_data = [cell.get_text(strip=True).replace('\n', ' ').replace('\r', ' ') for cell in cells]  # 替换换行符
                        if any(cell.strip() for cell in row_data):  # 检查行是否为空
                            table_data.append(row_data)
                    all_page_data.extend(table_data)

    # 将所有页面的数据保存到 DataFrame
    if all_page_data:
        df = pd.DataFrame(all_page_data, columns=table_headers)
        # 保存到 Excel 文件
        df.to_excel(r"C:\Users\27569\Desktop\移动运营商发电工单\fadian.xlsx", index=False)
        print("数据已保存到 Excel 文件中")
    else:
        print("没有提取到任何数据")
    #处理表格
    index_path = input('请输入文件夹路径（比如E:/abc）: ')
    process_excel_files(index_path)
    input('已全部完成，回车退出')
if __name__ == "__main__":
    # 运行主程序
    main()