import requests
from bs4 import BeautifulSoup

data_list=[{
  "AJAXREQUEST": "_viewRoot",
  "queryForm:unitHidden": "0099977",
  "queryForm:queryFlag": "queryFlag",
  "queryForm:queryStaStatusSelId_hiddenValue": "2",
  "queryForm:queryStaStatusSelId": "2",
  "queryForm:currPageObjId": "0",
  "queryForm:pageSizeText": "35",
  "queryForm": "queryForm",
  "javax.faces.ViewState": "j_id17",
  "queryForm:j_id155": "queryForm:j_id155"
},
  {
  "AJAXREQUEST": "_viewRoot",
  "queryForm:unitHidden": "0099977",
  "queryForm:queryFlag": "queryFlag",
  "queryForm:queryStaStatusSelId_hiddenValue": "2",
  "queryForm:queryStaStatusSelId": "2",
  "queryForm:currPageObjId": "0",
  "queryForm:pageSizeText": "35",
  "queryForm": "queryForm",
  "javax.faces.ViewState": "j_id17",
  "queryForm:j_id156": "queryForm:j_id156",
  "AJAX:EVENTS_COUNT": "1"
},{
  "AJAXREQUEST": "_viewRoot",
  "j_id814": "j_id814",
  "javax.faces.ViewState": "j_id17",
  "j_id814:j_id817": "j_id814:j_id817"
},{
  "j_id814": "j_id814",
  "j_id814:j_id816": "全部",
  "javax.faces.ViewState": "j_id17"
}]
headers={
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
  "Accept-Encoding": "gzip, deflate",
  "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
  "Connection": "keep-alive",
  "Cookie": "Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1745809498,1745939912; route=7634b5a86f92649382e66b47cf771393; ULTRA_U_K=; JSESSIONID=E612AD76AACBF01B21DE44713E66F5EF; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzQ1OTc3MDQzLCJOQU5PU0VDT05EIjozNDg1NDA5MDQwNTI0MzQwM30.oF79MiwfAl0lM0PUhQ9D5S8MnSL8OUaNV4caTp0xg48; acctId=101029143; uid=dw-wangcx9; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.105.4:all8380; BIGipServerywjk_new_pool1=42016172.10275.0000",
  "Host": "omms.chinatowercom.cn:9000",
  "Referer": "http://omms.chinatowercom.cn:9000/layout/index.xhtml",
  "Upgrade-Insecure-Requests": "1",
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
}
url='http://omms.chinatowercom.cn:9000/business/resMge/pwMge/fsuMge/listFsu.xhtml'
file_name='E:/todesk/fsu离线.xlsx'
cookies = {}
cookie_str = headers["Cookie"]
for cookie in cookie_str.split("; "):  # 按照分号和空格分割
    key, value = cookie.split("=", 1)  # 按照等号分割，只分割一次
    cookies[key] = value
res=requests.get(url=url,headers=headers,cookies=cookies)
soup = BeautifulSoup(res.text, 'html.parser')
view_state_input = soup.find('input', id='javax.faces.ViewState')
if view_state_input:
    javax = view_state_input.get('value')

######################
headers={
  "Accept": "*/*",
  "Accept-Encoding": "gzip, deflate",
  "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
  "Connection": "keep-alive",
  "Content-Length": "1418",
  "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
  "Cookie": "Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1745809498,1745939912; route=7634b5a86f92649382e66b47cf771393; ULTRA_U_K=; JSESSIONID=E612AD76AACBF01B21DE44713E66F5EF; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzQ1OTc3MDQzLCJOQU5PU0VDT05EIjozNDg1NDA5MDQwNTI0MzQwM30.oF79MiwfAl0lM0PUhQ9D5S8MnSL8OUaNV4caTp0xg48; acctId=101029143; uid=dw-wangcx9; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.105.4:all8380; BIGipServerywjk_new_pool1=42016172.10275.0000",
  "Host": "omms.chinatowercom.cn:9000",
  "Origin": "http://omms.chinatowercom.cn:9000",
  "Referer":"http://omms.chinatowercom.cn:9000/business/resMge/pwMge/fsuMge/listFsu.xhtml",
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0"
}
i=0
for data in data_list:
    i+=1
    data["javax.faces.ViewState"]=javax
    res=requests.post(url=url,data=data,headers=headers)
    if i==4:
        with open(file_name, "wb") as file:
            # 将响应的二进制内容写入文件
            file.write(res.content)




