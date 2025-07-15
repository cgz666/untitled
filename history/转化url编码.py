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
            'asOper=&beginGenerateDate=2025-05-12+00%3A00%3A00&endGenerateDate=2025-05-12+23%3A59%3A59'
 ]
# 处理每个查询字符串并打印结果
for i, query_string in enumerate(query_strings, start=1):
    result_dict = parse_query_string(query_string)
    # 使用 json.dumps 格式化字典为 JSON 格式的字符串，并保留方括号和双引号
    formatted_json = json.dumps(result_dict, ensure_ascii=False, indent=2)
    print(formatted_json+',')
    print()  # 添加空行分隔不同字典的输出