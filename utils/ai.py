import openai
import requests
import json
def ai_yzm(src):
    client = openai.OpenAI(
        api_key="sk-J4ndZW4RLXwrd32D40289e4dAdAc4306B4634cBbBcD5BdE7",
        # base_url="https://quchi-llm-oneapi.runjian.com/v1"  # 公网地址
        base_url="https://llm-oneapi.bytebroad.com.cn/v1"  # 或内网地址：
    )
    response = client.chat.completions.create(
        model="Qwen/Qwen2.5-VL-72B-Instruct",  # 当前私有化部署的多模态模型
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "识别图片中的验证码，结果只用输出验证码不要包含别的部分"},
                    {"type": "image_url", "image_url": {
                        "url": src
                    }}
                ]
            }
        ],
        max_tokens=2000
    )
    return response.choices[0].message.content

def ai_text(text):
    client = openai.OpenAI(
        api_key="sk-J4ndZW4RLXwrd32D40289e4dAdAc4306B4634cBbBcD5BdE7",
        # base_url="https://quchi-llm-oneapi.runjian.com/v1"  # 公网地址
        base_url="https://llm-oneapi.bytebroad.com.cn/v1"  # 或内网地址：
    )
    response = client.chat.completions.create(
        model="Qwen/Qwen2.5-VL-72B-Instruct",  # 当前私有化部署的多模态模型
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": text}
                ]
            }
        ],
        max_tokens=2000
    )
    return response.choices[0].message.content

print(ai_yzm(""))