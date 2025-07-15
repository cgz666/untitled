import requests
import csv
import os

def main():
    url = "https://energy-iot.chinatowercom.cn/api/device/station/list"
    headers = {
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Authorization": "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOlsiIl0sInVzZXJfbmFtZSI6IndlYl9tYW5hZ2V8cWprLWNoZW5neiIsInNjb3BlIjpbImFsbCJdLCJleHAiOjE3NDY2MjI4NDYsInVzZXJJZCI6MzQyMDEsImp0aSI6IjU5N2MzNjM3LWMxNWYtNDU3Ny04NDRhLTNkM2RjOTEzNGM2YSIsImNsaWVudF9pZCI6IndlYl9tYW5hZ2UifQ.Wl9K-3SxssVcRzPC_8nzz_pk2u4mJGPwe3gWiauMp3D0aibZAAtPvmjnfCsiZ_RaZLxn6JBbJm9XYLU-D4uBeWKtE02HmPljVQaGviNt85a5qogWw2NIv3UpLxNsWIMribZ90hGiKyFKhBXECam3QPTbeTQTbCFiYP-N4T5vJL_-M49kBpitMV-PCLJcPnVKmFslUPCr9hodVhhtGNJhNiHzGFSonIgYnwM2UbgqVx5iArEvw5F2hfoI_IPBl-kl9uWMHngEXSRVu_y7Ep36O6i6ewMSnNIEHNfZEukkOhd-6BNgYqX5HFsbBsOoEb768pZV6a4P54R4EtmVkeZEPQ",
        "Connection": "keep-alive",
        "Content-Type": "application/json;charset=UTF-8",
        "Cookie": "HWWAFSESID=dbfbd79d7a5e765d7d; HWWAFSESTIME=1746579613727; dc04ed2361044be8a9355f6efb378cf2=WyIzNTI0NjE3OTgzIl0",
        "Host": "energy-iot.chinatowercom.cn",
        "Origin": "https://energy-iot.chinatowercom.cn",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0",
        "sec-ch-ua": "\"Chromium\";v=\"130\", \"Microsoft Edge\";v=\"130\", \"Not?A_Brand\";v=\"99\"",
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": "\"Windows\""
    }

    # 创建一个文件夹来保存CSV文件
    output_folder = "output_data"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 定义业务类型
    business_types = [
        {"businessType": "2", "businessTypeName": "换电"},
        {"businessType": "1", "businessTypeName": "备电"},
        {"businessType": "4", "businessTypeName": "低速充电"}
    ]

    # 打开CSV文件
    with open(os.path.join(output_folder, "output.csv"), mode='w', newline='', encoding='utf-8') as file:
        csv_writer = csv.DictWriter(file, fieldnames=[
            '点位编码', '业务类型', '安全生产第一责任人', '监管责任人(安全员)', '责任领导', '站长', '包机人'
        ])
        csv_writer.writeheader()

        for business_type in business_types:
            # 循环遍历到第十页
            for page in range(1, 11):
                data = {
                    "pageNum": page,
                    "pageSize": 10,
                    "provinceId": "",
                    "cityId": "",
                    "countId": "",
                    "name": "",
                    "code": "",
                    "rsSource": "",
                    "businessType": business_type["businessType"],
                    "status": "",
                    "onlineStatus": "",
                    "maintenancePerson": "",
                    "secondBusinessType": "",
                    "deptIds": []
                }
                try:
                    response = requests.post(url=url, headers=headers, json=data)
                    response.raise_for_status()  # 检查HTTP状态码是否为200
                    json_data = response.json()

                    # 检查返回的数据结构
                    if 'rows' in json_data:
                        info_list = json_data['rows']
                        print(info_list)
                    else:
                        print(f"第 {page} 页返回的数据中没有找到 'rows' 键")
                        continue

                    # 遍历列表，提取每个元素的信息
                    for index in info_list:
                        dit = {
                            '点位编码': index.get('pubCode', ''),
                            '业务类型': business_type["businessTypeName"],
                            '安全生产第一责任人': index.get('firstDutyName', ''),
                            '监管责任人(安全员)': index.get('safeDutyName', ''),
                            '责任领导': index.get('dutyLeadName', ''),
                            '站长': index.get('stationMasterName', ''),
                            '包机人': index.get('charteredAirplaneName', ''),
                        }

                        csv_writer.writerow(dit)  # 写入CSV文件
                except requests.exceptions.RequestException as e:
                    print(f"请求第 {page} 页时出错：{e}")
                    break

if __name__ == "__main__":
    main()