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
     'AJAXREQUEST=_viewRoot&hisQueryForm=hisQueryForm&hisQueryForm%3AunitHidden=&hisQueryForm%3AunitHid=&hisQueryForm%3AqueryDay=30&hisQueryForm%3AqueryFaultMids_hiddenValue=%E9%80%80%E6%9C%8D%E5%9C%BA%E6%99%AF&hisQueryForm%3AqueryFaultMids=%E9%80%80%E6%9C%8D%E5%9C%BA%E6%99%AF&hisQueryForm%3AqueryFaultDetail=&hisQueryForm%3AqueryFaultDetailName=&hisQueryForm%3AqueryLevel_hiddenValue=&hisQueryForm%3Aj_id201=&hisQueryForm%3Aj_id205=&hisQueryForm%3Aj_id209=&hisQueryForm%3Aj_id213=&hisQueryForm%3Aj_id217=&hisQueryForm%3Aj_id221=&hisQueryForm%3AfirststarttimeInputDate=2025-07-01%2000%3A00&hisQueryForm%3AfirststarttimeInputCurrentDate=07%2F2025&hisQueryForm%3AfirstendtimeInputDate=2025-07-24%2000%3A00&hisQueryForm%3AfirstendtimeInputCurrentDate=07%2F2025&hisQueryForm%3Aj_id229=&hisQueryForm%3ArecoverstarttimeInputDate=&hisQueryForm%3ArecoverstarttimeInputCurrentDate=07%2F2025&hisQueryForm%3ArecoverendtimeInputDate=&hisQueryForm%3ArecoverendtimeInputCurrentDate=07%2F2025&hisQueryForm%3Aj_id237=&hisQueryForm%3AqueryFsuStatus_hiddenValue=&hisQueryForm%3AcurrPageObjId=0&hisQueryForm%3ApageSizeText=35&javax.faces.ViewState=j_id6&hisQueryForm%3Aj_id244=hisQueryForm%3Aj_id244&',
    'AJAXREQUEST=_viewRoot&hisQueryForm=hisQueryForm&hisQueryForm%3AunitHidden=&hisQueryForm%3AunitHid=&hisQueryForm%3AqueryDay=30&hisQueryForm%3AqueryFaultMids_hiddenValue=%E9%80%80%E6%9C%8D%E5%9C%BA%E6%99%AF&hisQueryForm%3AqueryFaultMids=%E9%80%80%E6%9C%8D%E5%9C%BA%E6%99%AF&hisQueryForm%3AqueryFaultDetail=&hisQueryForm%3AqueryFaultDetailName=&hisQueryForm%3AqueryLevel_hiddenValue=&hisQueryForm%3Aj_id201=&hisQueryForm%3Aj_id205=&hisQueryForm%3Aj_id209=&hisQueryForm%3Aj_id213=&hisQueryForm%3Aj_id217=&hisQueryForm%3Aj_id221=&hisQueryForm%3AfirststarttimeInputDate=2025-07-01%2000%3A00&hisQueryForm%3AfirststarttimeInputCurrentDate=07%2F2025&hisQueryForm%3AfirstendtimeInputDate=2025-07-24%2000%3A00&hisQueryForm%3AfirstendtimeInputCurrentDate=07%2F2025&hisQueryForm%3Aj_id229=&hisQueryForm%3ArecoverstarttimeInputDate=&hisQueryForm%3ArecoverstarttimeInputCurrentDate=07%2F2025&hisQueryForm%3ArecoverendtimeInputDate=&hisQueryForm%3ArecoverendtimeInputCurrentDate=07%2F2025&hisQueryForm%3Aj_id237=&hisQueryForm%3AqueryFsuStatus_hiddenValue=&hisQueryForm%3AcurrPageObjId=1&hisQueryForm%3ApageSizeText=35&javax.faces.ViewState=j_id6&hisQueryForm%3Aj_id245=hisQueryForm%3Aj_id245&AJAX%3AEVENTS_COUNT=1&',
    'AJAXREQUEST=_viewRoot&hisQueryForm=hisQueryForm&hisQueryForm%3AunitHidden=&hisQueryForm%3AunitHid=&hisQueryForm%3AqueryDay=30&hisQueryForm%3AqueryFaultMids_hiddenValue=%E9%80%80%E6%9C%8D%E5%9C%BA%E6%99%AF&hisQueryForm%3AqueryFaultMids=%E9%80%80%E6%9C%8D%E5%9C%BA%E6%99%AF&hisQueryForm%3AqueryFaultDetail=&hisQueryForm%3AqueryFaultDetailName=&hisQueryForm%3AqueryLevel_hiddenValue=&hisQueryForm%3Aj_id201=&hisQueryForm%3Aj_id205=&hisQueryForm%3Aj_id209=&hisQueryForm%3Aj_id213=&hisQueryForm%3Aj_id217=&hisQueryForm%3Aj_id221=&hisQueryForm%3AfirststarttimeInputDate=2025-07-01%2000%3A00&hisQueryForm%3AfirststarttimeInputCurrentDate=07%2F2025&hisQueryForm%3AfirstendtimeInputDate=2025-07-24%2000%3A00&hisQueryForm%3AfirstendtimeInputCurrentDate=07%2F2025&hisQueryForm%3Aj_id229=&hisQueryForm%3ArecoverstarttimeInputDate=&hisQueryForm%3ArecoverstarttimeInputCurrentDate=07%2F2025&hisQueryForm%3ArecoverendtimeInputDate=&hisQueryForm%3ArecoverendtimeInputCurrentDate=07%2F2025&hisQueryForm%3Aj_id237=&hisQueryForm%3AqueryFsuStatus_hiddenValue=&hisQueryForm%3AcurrPageObjId=1&hisQueryForm%3ApageSizeText=35&javax.faces.ViewState=j_id6&hisQueryForm%3Aj_id249=hisQueryForm%3Aj_id249&',
    'AJAXREQUEST=_viewRoot&j_id407=j_id407&javax.faces.ViewState=j_id6&j_id407%3Aj_id410=j_id407%3Aj_id410&',
    'j_id407=j_id407&j_id407%3Aj_id409=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id6'
]
# 处理每个查询字符串并打印结果
for i, query_string in enumerate(query_strings, start=1):
    result_dict = parse_query_string(query_string)
    # 使用 json.dumps 格式化字典为 JSON 格式的字符串，并保留方括号和双引号
    formatted_json = json.dumps(result_dict, ensure_ascii=False, indent=2)
    print(formatted_json+',')
    print()  # 添加空行分隔不同字典的输出