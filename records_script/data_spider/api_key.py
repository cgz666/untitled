import requests

def get_response(page):
    api_key = "2f44c011a9d9427a9733d870657bf5d2"
    api_url = "https://beian.china-eia.com/a/registrationform/tBasRegistrationForm/formIndex"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Host": "beian.china-eia.com",
        "Cookie": f"sysweb.session.id={api_key}",
        "Referer": "https://beian.china-eia.com/a/registrationform/tBasRegistrationForm/formIndex",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
        "sec-ch-ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"'
    }
    data = {
        'orderBy': '',
        'pageNo': str(page + 1),
        'pageSize': '10'
    }
    response = requests.post(url = api_url, headers=headers, json=data)
    if response.status_code == 200:
        result = response.content
        print(result)
        return (result['choices'][0]['message']['content'])
    else:
        print("请求失败，状态码：", response.status_code)
        print("错误信息：", response.text)

get_response(0)


