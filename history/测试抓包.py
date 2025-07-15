import os
import requests
from datetime import datetime, timedelta

def get_date_range():
    """获取本月1号和昨天的日期"""
    today = datetime.now()
    first_day_of_month = today.replace(day=1).strftime("%Y-%m-%d 00:00:00")
    yesterday = (today - timedelta(days=1)).strftime("%Y-%m-%d 23:59:59")
    return first_day_of_month, yesterday

def send_request(url, headers, data):
    """发送POST请求并处理响应"""
    try:
        session = requests.Session()
        response = session.post(url, headers=headers, data=data)
        response.raise_for_status()  # 检查请求是否成功
        return response
    except requests.RequestException as e:
        print(f"请求发生错误: {e}")
        return None

def save_response_as_excel(response, output_path):
    """将响应内容保存为Excel文件"""
    try:
        with open(output_path, "wb") as file:
            file.write(response.content)
        print(f"Excel 文件已保存为 {output_path}")
    except IOError as e:
        print(f"文件保存失败: {e}")

def main():
    # 配置参数
    url = "http://clound.gxtower.cn:11080/tower_manage_bms/a/tower/oil/report/exportOperatorReport"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Cache-Control": "max-age=0",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded",
        "Cookie": "jeesite.session.id=79ff3ee9c5414324896bcf2f9d7592dc; JSESSIONID=17A7F742D0ECD29B94ABA14FE4586E1E; Hm_lvt_82116c626a8d504a5c0675073362ef6f=1746667036,1746751000,1746850695,1747010545; HMACCOUNT=4DF837C069A2D971; pageNo=1; pageSize=25; Hm_lpvt_82116c626a8d504a5c0675073362ef6f=1747015641",
        "Host": "clound.gxtower.cn:11080",
        "Origin": "http://clound.gxtower.cn:11080",
        "Referer": "http://clound.gxtower.cn:11080/tower_manage_bms/a/tower/oil/report/operatorReport",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
    }
    folder_path = r"E:\陈桂志\普通文件"
    output_path = os.path.join(folder_path, "response.xlsx")

    # 获取日期范围
    first_day_of_month, yesterday = get_date_range()

    # 准备请求数据
    data = {
        "asOper": "",
        "beginGenerateDate": first_day_of_month,
        "endGenerateDate": yesterday,
        "pageNo": "1",
        "pageSize": "25",
        "orderBy": "",
        "city.id": "",
        "city.name": "",
        "area.id": "20011335,1129,1130,1131,1132,1133,1134,1135,1136,1137,1138,1139,1140,1828749314,1142,1143,1144,1145,1146,1147,1148,1149,1150,1151,1153,1154,1155,1156,1157,1158,1159,1160,1161,1162,1163,1164,1165,1166,1167,1168,1169,1170,1171,1172,1173,1174,1175,1176,1177,1179,1180,1181,1182,1183,1184,10010113,1185,1186,1187,1188,1189,1190,1191,1192,1193,1196,1197,1198,1199,1200,1201,1202,1203,1204,1205,1206,1207,1208,1209,1210,1211,1212,1213,1214,1215,1216,1217,1219,1220,1221,1222,1223,1224,1225,1226,1227,1228,1229,1230,1231,1232,1233,1234,1235,1236,1237,1238,1239,1240,1241,1242",
        "area.name": "",
        "stationCode": "",
        "stationName": "",
        "number": "",
        "generatePowerState": "",
        "finishConfigId": "",
        "approvalOfDispatchId": "",
        "generateOfficeName": "",
        "stopBegeinDate": "",
        "stopEndDate": ""
    }

    # 发送请求
    response = send_request(url, headers, data)
    if response:
        # 保存响应为Excel文件
        save_response_as_excel(response, output_path)

if __name__ == "__main__":
    main()