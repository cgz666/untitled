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
     'AJAXREQUEST=_viewRoot&stationListForm%3AnameText=%E5%8D%97%E5%AE%81%E9%A9%AC%E5%B1%B1%E5%8E%BF%E9%87%8C%E5%BD%93%E4%B9%A1%E6%A3%8B%E7%9B%98%E5%9F%BA%E7%AB%99%E6%97%A0%E7%BA%BF%E6%9C%BA%E6%88%BF&stationListForm%3AstationidText=&stationListForm%3AqueryStatusId=&stationListForm%3Astationcode=&stationListForm%3AcurrPageObjId=0&stationListForm=stationListForm&autoScroll=&javax.faces.ViewState=j_id6&stationListForm%3Aj_id283=stationListForm%3Aj_id283&',
    'AJAXREQUEST=_viewRoot&stationListForm%3AnameText=%E5%8D%97%E5%AE%81%E9%A9%AC%E5%B1%B1%E5%8E%BF%E9%87%8C%E5%BD%93%E4%B9%A1%E6%A3%8B%E7%9B%98%E5%9F%BA%E7%AB%99%E6%97%A0%E7%BA%BF%E6%9C%BA%E6%88%BF&stationListForm%3AstationidText=&stationListForm%3AqueryStatusId=&stationListForm%3Astationcode=&stationListForm%3AcurrPageObjId=0&stationListForm=stationListForm&autoScroll=&javax.faces.ViewState=j_id6&stationListForm%3Aj_id284=stationListForm%3Aj_id284&AJAX%3AEVENTS_COUNT=1&',
    'AJAXREQUEST=_viewRoot&stationListForm%3AnameText=%E5%8D%97%E5%AE%81%E9%A9%AC%E5%B1%B1%E5%8E%BF%E9%87%8C%E5%BD%93%E4%B9%A1%E6%A3%8B%E7%9B%98%E5%9F%BA%E7%AB%99%E6%97%A0%E7%BA%BF%E6%9C%BA%E6%88%BF&stationListForm%3AstationidText=&stationListForm%3AqueryStatusId=&stationListForm%3Astationcode=&stationListForm%3AcurrPageObjId=0&selectFlag=%E5%8D%97%E5%AE%81%E9%A9%AC%E5%B1%B1%E5%8E%BF%E9%87%8C%E5%BD%93%E4%B9%A1%E6%A3%8B%E7%9B%98%E5%9F%BA%E7%AB%99%E6%97%A0%E7%BA%BF%E6%9C%BA%E6%88%BF&stationListForm=stationListForm&autoScroll=&javax.faces.ViewState=j_id6&staId=9838307&j_id328=j_id328&AJAX%3AEVENTS_COUNT=1&',
    'AJAXREQUEST=_viewRoot&queryForm2=queryForm2&queryForm2%3Aj_id123=&queryForm2%3Aaid=9838307&queryForm2%3Apanel2OpenedState=&javax.faces.ViewState=j_id6&queryForm2%3AqueryDevice=queryForm2%3AqueryDevice&',
    'AJAXREQUEST=_viewRoot&queryForm2=queryForm2&queryForm2%3Aj_id123=&queryForm2%3Aaid=9838307&queryForm2%3Apanel2OpenedState=&javax.faces.ViewState=j_id6&queryForm2%3Aj_id125=queryForm2%3Aj_id125&AJAX%3AEVENTS_COUNT=1&',
    'AJAXREQUEST=_viewRoot&queryForm3=queryForm3&queryForm3%3Amname=&queryForm3%3Adid=45012440601028&queryForm3%3Apanel3OpenedState=&javax.faces.ViewState=j_id6&queryForm3%3AqueryMid=queryForm3%3AqueryMid&',
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3AaddOrEditAreaNameId=%E5%8D%97%E5%AE%81%E9%A9%AC%E5%B1%B1%E5%8E%BF%E9%87%8C%E5%BD%93%E4%B9%A1%E6%A3%8B%E7%9B%98%E5%9F%BA%E7%AB%99%E6%97%A0%E7%BA%BF%E6%9C%BA%E6%88%BF&queryForm%3Aaid=9838307&queryForm%3Afsuid=&queryForm%3AdeviceName=%E5%8D%97%E5%AE%81%E9%A9%AC%E5%B1%B1%E5%8E%BF%E9%87%8C%E5%BD%93%E4%B9%A1%E6%A3%8B%E7%9B%98%E5%9F%BA%E7%AB%99%E6%97%A0%E7%BA%BF%E6%9C%BA%E6%88%BF%2F%E5%BC%80%E5%85%B3%E7%94%B5%E6%BA%9001&queryForm%3Adid=45012440601028&queryForm%3AmidName=&queryForm%3Amid=&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id6&queryForm%3Aj_id17=queryForm%3Aj_id17&',
    'AJAXREQUEST=_viewRoot&queryForm3=queryForm3&queryForm3%3Amname=&queryForm3%3Adid=45012440601028&queryForm3%3Apanel3OpenedState=&javax.faces.ViewState=j_id6&queryForm3%3Aj_id157=queryForm3%3Aj_id157&AJAX%3AEVENTS_COUNT=1&',
    'AJAXREQUEST=_viewRoot&queryForm3=queryForm3&queryForm3%3Amname=%E7%94%B5%E6%B1%A0%E5%85%85%E7%94%B5%E9%99%90%E6%B5%81%E8%AE%BE%E5%AE%9A%E5%80%BC&queryForm3%3Adid=45012440601028&queryForm3%3Apanel3OpenedState=&javax.faces.ViewState=j_id6&queryForm3%3AqueryMid=queryForm3%3AqueryMid&',
    'AJAXREQUEST=_viewRoot&queryForm3=queryForm3&queryForm3%3Amname=%E7%94%B5%E6%B1%A0%E5%85%85%E7%94%B5%E9%99%90%E6%B5%81%E8%AE%BE%E5%AE%9A%E5%80%BC&queryForm3%3Adid=45012440601028&queryForm3%3Apanel3OpenedState=&javax.faces.ViewState=j_id6&queryForm3%3Aj_id157=queryForm3%3Aj_id157&AJAX%3AEVENTS_COUNT=1&'
 ]
# 处理每个查询字符串并打印结果
for i, query_string in enumerate(query_strings, start=1):
    result_dict = parse_query_string(query_string)
    # 使用 json.dumps 格式化字典为 JSON 格式的字符串，并保留方括号和双引号
    formatted_json = json.dumps(result_dict, ensure_ascii=False, indent=2)
    print(formatted_json+',')
    print()  # 添加空行分隔不同字典的输出