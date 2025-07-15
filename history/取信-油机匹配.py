import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import os
import win32com.client as win32

# 配置模块
url = "http://clound.gxtower.cn:11080/tower_manage_bms/a/tower/oil/towerOilStatistics/countOilAndGenerateMatch?Token=85d083bd536d4da49d568a216b19adfa&APPURI=aHR0cDovL2Nsb3VuZC5neHRvd2VyLmNuOjExMDgwL3Rvd2VyX21hbmFnZV9ibXMvZi8=&MID=fa16591065be47a8a77b9a18b2cba868"
headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Cache-Control": "max-age=0",
    "Connection": "keep-alive",
    "Content-Length": "92",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cookie": "jeesite.session.id=79ff3ee9c5414324896bcf2f9d7592dc; JSESSIONID=17A7F742D0ECD29B94ABA14FE4586E1E; Hm_lvt_82116c626a8d504a5c0675073362ef6f=1746667036,1746751000,1746850695,1747010545; HMACCOUNT=4DF837C069A2D971; Hm_lpvt_82116c626a8d504a5c0675073362ef6f=1747010784",
    "Host": "clound.gxtower.cn:11080",
    "Origin": "http://clound.gxtower.cn:11080",
    "Referer": "http://clound.gxtower.cn:11080/tower_manage_bms/a/tower/oil/towerOilStatistics/countOilAndGenerateMatch",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
}

# 创建一个Session对象
session = requests.Session()
# 设置请求头
session.headers.update(headers)
# 提取Cookie为字典
cookies = {
    "jeesite.session.id": "358dcae3a6b94b82b4fc4261958af459",
    "JSESSIONID": "A30DEF6A8029D03A9B3DB097C96A296D",
    "pageSize": "25",
    "pageNo": "1",
    "Hm_lvt_82116c626a8d504a5c0675073362ef6f": "1746667036,1746751000",
    "HMACCOUNT": "4DF837C069A2D971",
    "Hm_lpvt_82116c626a8d504a5c0675073362ef6f": "1746757446"
}

# 将Cookie添加到Session对象
session.cookies.update(cookies)

# 获取本月一号和今天的前一天的日期
today = datetime.today()
first_day_of_month = today.replace(day=1)
yesterday = today - timedelta(days=1)

# 格式化日期为字符串
begin_generate_date = first_day_of_month.strftime("%Y-%m-%d 00:00:00")
end_generate_date = yesterday.strftime("%Y-%m-%d 23:59:59")

# 数据列表
data_list = [
    {
        "asOper": "",
        "beginGenerateDate": begin_generate_date,
        "endGenerateDate": end_generate_date
    }
]

# 初始化一个空的列表，用于存储最后一个页面的数据
last_page_data = []

# 遍历data_list
for i, data in enumerate(data_list):
    # 发送POST请求
    response = session.post(url, data=data)
    response.raise_for_status()  # 检查请求是否成功

    # 检查响应内容类型
    if 'text/html' in response.headers.get('Content-Type', ''):
        # 解析HTML内容
        soup = BeautifulSoup(response.content, 'html.parser')

        # 提取表格数据
        table = soup.find('table')  # 假设数据在第一个表格中
        if table:
            rows = table.find_all('tr')
            if i == 0:  # 提取表头只在第一次请求时进行
                table_headers = [th.get_text(strip=True) for th in rows[0].find_all('th')]
            table_data = []
            for row in rows[1:]:
                cells = row.find_all('td')
                row_data = [cell.get_text(strip=True).replace('\n', ' ').replace('\r', ' ') for cell in cells]  # 替换换行符
                if any(cell.strip() for cell in row_data):  # 检查行是否为空
                    table_data.append(row_data)
            if i == len(data_list) - 1:  # 如果是最后一个元素，保存数据
                last_page_data.extend(table_data)
    else:
        print("响应内容不是HTML，可能是JSON或其他格式")

# 将最后一个页面的数据保存到DataFrame
if last_page_data:
    df = pd.DataFrame(last_page_data, columns=table_headers)
    # 保存到Excel文件
    df.to_excel(r"C:\Users\27569\Desktop\油机匹配\youdian.xlsx", index=False)
    print("数据已保存到Excel文件中")
else:
    print("没有提取到最后一个页面的数据")

# 处理Excel文件
def process_excel_files(index_path):
    """
    处理Excel文件，将指定文件夹中的数据文件内容复制到主表文件中。

    :param index_path: 文件夹路径
    """
    print('1、把数据文件和通报模板放在同一文件夹下')
    print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
    index_path = index_path.replace('\\', '/')  # 确保路径分隔符统一
    # 打开主表文件
    xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
    xl.Visible = True  # 窗口是否可见
    main_file = os.path.join(index_path, '工单及油机匹配率统计（5月7日)  .xlsx')  # 要处理的文件路径
    workbook_main = xl.Workbooks.Open(main_file)  # 打开上述路径文件

    # 故障工单
    for file_path in os.listdir(index_path):
        if 'youdian' in file_path:
            data_file = os.path.join(index_path, file_path)
            workbook_data = xl.Workbooks.Open(data_file)
            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('Export')
            sheet_data.AutoFilterMode = False  # 全局关闭筛选，保证复制的数据完全
            source_range = sheet_data.Range('A2:J16')
            target_range = sheet_main.Range('A3:J17')
            source_range.Copy()
            target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)  # 使用值的形式粘贴
            xl.CutCopyMode = False  # 释放剪切板
            workbook_data.Close(SaveChanges=False)

    workbook_main.SaveAs(os.path.join(index_path, '网络指标_更新后.xlsx'))
    workbook_main.Close()
    xl.Quit()  # 关闭Excel应用程序
    print('已全部完成')

# 主函数
def main():
    index_path = input('请输入文件夹路径（比如E:/abc）: ')
    process_excel_files(index_path)
    input('已全部完成，回车退出')

if __name__ == "__main__":
    main()