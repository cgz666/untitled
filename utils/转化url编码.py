import json
from urllib.parse import parse_qs, unquote

def parse_query_string(query_string):
    # 使用 parse_qs 解析查询字符串
    parsed_dict = parse_qs(query_string)

    # 解码 URL 编码的键和值
    decoded_query_string = unquote(query_string)
    parsed_dict = parse_qs(decoded_query_string, keep_blank_values=True)
    decoded_dict = {k: v[0] if v else '' for k, v in parsed_dict.items()}
    # 如果每个键只有一个值，可以将值从列表中提取出来
    # 保留值为空的情况
    final_dict = {k: v[0] if len(v) == 1 else v for k, v in decoded_dict.items()}

    return final_dict

# 示例：多个查询字符串
query_strings = [
     'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977%2C0099978%2C0099979&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=&queryForm%3Amid=&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=0406111001&queryForm%3AquerySpeIdShow=0406111001...&queryForm%3AstarttimeInputDate=2025-07-29%2008%3A16&queryForm%3AstarttimeInputCurrentDate=07%2F2025&queryForm%3AstarttimeTimeHours=08&queryForm%3AstarttimeTimeMinutes=16&queryForm%3AendtimeInputDate=2025-07-29%2010%3A16&queryForm%3AendtimeInputCurrentDate=07%2F2025&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id4&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977%2C0099978%2C0099979&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=&queryForm%3Amid=&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=0406111001&queryForm%3AquerySpeIdShow=0406111001...&queryForm%3AstarttimeInputDate=2025-07-29%2008%3A16&queryForm%3AstarttimeInputCurrentDate=07%2F2025&queryForm%3AstarttimeTimeHours=08&queryForm%3AstarttimeTimeMinutes=16&queryForm%3AendtimeInputDate=2025-07-29%2010%3A16&queryForm%3AendtimeInputCurrentDate=07%2F2025&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id4&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id4'
]
# 处理每个查询字符串并打印结果
for i, query_string in enumerate(query_strings, start=1):
    result_dict = parse_query_string(query_string)
    # 使用 json.dumps 格式化字典为 JSON 格式的字符串，并保留方括号和双引号
    formatted_json = json.dumps(result_dict, ensure_ascii=False, indent=2)
    print(formatted_json+',')
    print()  # 添加空行分隔不同字典的输出