import requests

def get_response(prompt):
    api_key = "sk-J4ndZW4RLXwrd32D40289e4dAdAc4306B4634cBbBcD5BdE7"
    model_name = "Qwen/Qwen2.5-32B-Instruct"
    # 移除了URL末尾的多余空格
    api_url = "https://quchi-llm-oneapi.runjian.com/v1/chat/completions"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }

    data = {
        "model": model_name,
        "messages": [{"role": "user", "content": prompt}]
    }

    try:
        # 设置超时参数，避免程序长时间等待无响应的请求
        response = requests.post(api_url, headers=headers, json=data, timeout=30)
        # 使用raise_for_status自动处理HTTP错误
        response.raise_for_status()
        result = response.json()
        # 添加类型检查，确保返回格式符合预期
        if (
                isinstance(result, dict) and
                'choices' in result and
                len(result['choices']) > 0 and
                'message' in result['choices'][0] and
                'content' in result['choices'][0]['message']
        ):
            return result['choices'][0]['message']['content']
        else:
            print("API返回格式异常:", result)
            return None
    except requests.exceptions.Timeout:
        print("请求超时，请检查网络连接或API服务器状态")
        return None
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP错误发生: {http_err}")
        print("错误详情:", response.text)
        return None
    except requests.exceptions.RequestException as req_err:
        print(f"请求发生异常: {req_err}")
        return None
    except ValueError as json_err:
        print(f"JSON解析错误: {json_err}")
        print("响应内容:", response.text)
        return None


# 调用示例
if __name__ == "__main__":
    prompt ='''
'''

    response = get_response(prompt)
    if response:
        print("AI回复:", response)



