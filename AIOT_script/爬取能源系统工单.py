import requests
import hmac
import hashlib
import base64
import time
from urllib.parse import quote

# 配置你的华为云 OBS 凭证
ACCESS_KEY_ID = "XNSIE6TG10ZDMCRR6QOQ"
SECRET_KEY = "your_secret_key_here"  # 替换为你的 SecretKey

# 目标文件路径
BUCKET_NAME = "sc-aiot-obs"
OBJECT_NAME = "Users/export/iot/datafile/exportjob/result/%E5%B7%A5%E5%8D%95%E7%9B%91%E6%8E%A7-20250526102603284150.xlsx"

# 生成签名的函数
def generate_signature(access_key, secret_key, http_method, content_md5, content_type, expires, bucket_name, object_name):
    # 对象路径需要进行 URL 编码
    encoded_object_name = quote(object_name, safe='')
    string_to_sign = f"{http_method}\n{content_md5}\n{content_type}\n{expires}\n/{bucket_name}/{encoded_object_name}"
    signature = hmac.new(secret_key.encode('utf-8'), string_to_sign.encode('utf-8'), hashlib.sha1).digest()
    return base64.b64encode(signature).decode('utf-8')

# 设置请求参数
http_method = "GET"
content_md5 = ""  # 如果没有内容MD5，可以留空
content_type = ""  # 如果没有内容类型，可以留空
expires = int(time.time()) + 3600  # 当前时间戳 + 1小时

# 生成签名
signature = generate_signature(ACCESS_KEY_ID, SECRET_KEY, http_method, content_md5, content_type, expires, BUCKET_NAME, OBJECT_NAME)

# 请求参数
params = {
    "AccessKeyId": ACCESS_KEY_ID,
    "Expires": expires,
    "Signature": signature
}

# 请求头
headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Connection": "keep-alive",
    "Host": "sc-aiot-obs.obs.cn-north-4.myhuaweicloud.com",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "cross-site",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
    "sec-ch-ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"'
}

# 目标 URL
url = f"https://sc-aiot-obs.obs.cn-north-4.myhuaweicloud.com/{BUCKET_NAME}/{OBJECT_NAME}"

# 发送请求
response = requests.get(url=url, headers=headers, params=params)

# 打印响应内容
print(response.content)