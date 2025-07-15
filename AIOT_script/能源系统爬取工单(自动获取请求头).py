import requests
import csv
import time

# 获取cookie的URL
cookie_url = "http://clound.gxtower.cn:3980/tt/get_aiot_cookie"

# 发送GET请求获取cookie
res = requests.get(cookie_url)

# 检查请求是否成功
if res.status_code == 200:
    cookie = res.text.strip()  # 获取响应内容并去除两端的空白字符
else:
    print(f"Failed to get cookie: {res.status_code}")
    cookie = None

# 更新headers中的Authorization值
headers = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Authorization": f"{cookie}",
    "Connection": "keep-alive",
    "Content-Type": "application/json;charset=UTF-8",
    "Cookie": "HWWAFSESID=473bf768718adcb2dd; HWWAFSESTIME=1747185603088; dc04ed2361044be8a9355f6efb378cf2=WyIyODM4MDM2MDcxIl0",
    "Host": "energy-iot.chinatowercom.cn",
    "Origin": "https://energy-iot.chinatowercom.cn",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0.0 Safari/537.36 Edg/136.0.0.0",
    "sec-ch-ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"'
}

# 目标API URL
url = "https://energy-iot.chinatowercom.cn/api/workorder/workOrder/page"

# 遍历页数范围
total_pages = 10  # 假设要遍历的总页数
info_list = []

for page_num in range(1, total_pages + 1):
    data = {
        "pageNum": page_num,
        "pageSize": 10,
    }

    # 发送POST请求
    response = requests.post(url=url, headers=headers, json=data)
    if response.status_code == 200:
        response_json = response.json()
        if 'rows' in response_json:
            info_list.extend(response_json['rows'])
            print(f"成功获取第 {page_num} 页数据")
        else:
            print(f"第 {page_num} 页响应中没有 'rows' 键")
    else:
        print(f"第 {page_num} 页请求失败，状态码：{response.status_code}")
    time.sleep(1)  # 避免过于频繁的请求

# 打开 CSV 文件并写入数据
with open('output/output.csv', mode='w', newline='', encoding='utf-8') as file:
    csv_writer = csv.DictWriter(file, fieldnames=[
        '管理省份', '管理城市', '管理区县', '工单编码', '工单标题', '告警等级',
        '工单状态', '接单考核', '回单考核', '业务类型', '点位名称', '点位编码',
        '点位业务编码', '设备名称', '建单时间', '接单时间', '工单类型',
        '故障类型', '工单来源', '双打卡状态'
    ])
    csv_writer.writeheader()
    for index in info_list:
        dit = {
            '管理省份': index.get('provinceName', ''),
            '管理城市': index.get('cityName', ''),
            '管理区县': index.get('countyName', ''),
            '工单编码': index.get('workOrderCode', ''),
            '工单标题': index.get('workOrderTitle', ''),
            '告警等级': index.get('alarmLevelName', ''),
            '工单状态': index.get('workOrderStatusName', ''),
            '接单考核': index.get('receiverAssessName', ''),
            '回单考核': index.get('receiptAssessName', ''),
            '业务类型': index.get('businessTypeName', ''),
            '点位名称': index.get('stationName', ''),
            '点位编码': index.get('stationPubCode', ''),
            '点位业务编码': index.get('stationCode', ''),
            '设备名称': index.get('devName', ''),
            '建单时间': index.get('createTime', ''),
            '接单时间': index.get('acceptTime', ''),
            '工单类型': index.get('workOrderTypeName', ''),
            '故障类型': index.get('faultTypeName', ''),
            '工单来源': index.get('faultSourceName', ''),
            '双打卡状态': index.get('doubleCardStatusName', '')
        }
        csv_writer.writerow(dit)  # 写入CSV文件

print(f'文件已保存到{file.name}')