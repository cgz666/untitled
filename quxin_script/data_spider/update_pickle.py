import requests
from config import INDEX
import os

# 确保路径中的反斜杠被正确处理
pickle_path = f'{INDEX}/quxin_script/data_spider/pickle_quxin.pkl'

def update_pickle():
    url = 'http://clound.gxtower.cn:3980/tt/get_session_quxin'
    try:
        # 下载pickle
        res = requests.get(url)
        res.raise_for_status()  # 检查请求是否成功
        with open(pickle_path, "wb") as file:
            file.write(res.content)
        print("Pickle文件下载成功")
    except requests.exceptions.RequestException as e:
        print(f"下载Pickle文件时出错: {e}")
        print("请检查链接的合法性或网络连接，并重试。")

# 调用函数
update_pickle()