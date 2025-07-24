from bs4 import BeautifulSoup
import os
from datetime import datetime
import ast
import pandas as pd
import re
import time
import base64
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
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
                    {"type": "text", "text": "识别图片中的4位验证码，结果只用输出验证码不要包含别的部分"},
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

class main():
    def __init__(self):
        INDEX = os.getcwd()
        self.result_df = None  # 初始化为空
        self.driver_path = os.path.join(INDEX, "chromedriver.exe")
        self.chrome_path = os.path.join(INDEX, "chrome-win64/chrome.exe")
        self.headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br, zstd',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Host': 'beian.china-eia.com',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Google Chrome";v="135", "Not-A.Brand";v="8", "Chromium";v="135"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"'
        }

        self.xls_path = os.path.join(INDEX, "xls")
        self.temp_path = os.path.join(INDEX, "temp")
        self.output_path = os.path.join(INDEX, "output")
        self.account_path = os.path.join(self.xls_path, "环保厅账密.xlsx")  # 确保文件路径正确
        self.white_list_path = os.path.join(self.xls_path, "白名单.xlsx")  # 确保文件路径正确
        # 用户手动输入开始日期和结束日期
        self.start_date = self.get_user_input_date("请输入开始日期 (格式: YYYY-MM-DD): ")
        self.end_date = self.get_user_input_date("请输入结束日期 (格式: YYYY-MM-DD): ")

    def get_user_input_date(self, prompt):
        """获取用户输入的日期，并验证格式"""
        while True:
            date_str = input(prompt)
            try:
                return datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                print("日期格式不正确，请重新输入 (格式: YYYY-MM-DD)。")

    def is_date_in_range(self, date_str):
        """检查日期是否在指定范围内，若早于开始日期返回-1，在范围内返回1，无效或晚于结束日期返回0"""
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        if date_obj < self.start_date:
            return -1  # 早于开始日期
        elif date_obj > self.end_date:
            return 0  # 晚于结束日期
        else:
            return 1  # 在范围内

    def login(self, username, password):
        """自动打开浏览器，使用指定的账号密码登录"""
        service = Service(executable_path=self.driver_path)
        option = webdriver.ChromeOptions()
        # option.add_argument('--headless')  # 如果需要无头模式，取消注释
        option.add_argument("--disable-gpu")
        option.add_argument("--no-sandbox")
        option.binary_location = self.chrome_path
        driver = webdriver.Chrome(service=service, options=option)
        max_attempts = 5  # 设置最大重试次数
        attempts = 0

        while attempts < max_attempts:
            try:
                # 打开登录页面
                login_url = "https://beian.china-eia.com/a/login"
                driver.get(login_url)
                driver.save_screenshot("full_page.png")
                with open("full_page.png", "rb") as image_file:
                    encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                src = f"data:image/png;base64,{encoded_string}"
                res = ai_yzm(src)
                print(f"验证码识别结果: {res}")

                # 选择用户名登录方式
                login_method_menu = driver.find_element(By.CSS_SELECTOR, ".select2-choice")
                login_method_menu.click()
                username_login_option = driver.find_element(By.XPATH,
                                                            "//div[@class='select2-result-label' and contains(text(), '用户名登录')]")
                username_login_option.click()

                # 输入账号和密码
                username_field = driver.find_element(By.ID, "username")
                password_field = driver.find_element(By.ID, "password")
                vcode_field = driver.find_element(By.ID, "validateCode")
                username_field.send_keys(username)
                password_field.send_keys(password)
                vcode_field.send_keys(res)

                # 点击登录按钮
                login_button = driver.find_element(By.ID, "denglu")
                login_button.click()

                # 等待登录完成
                time.sleep(5)

                # 检查是否登录成功
                if driver.current_url == "https://beian.china-eia.com/a?login":
                    print("登录成功")
                    cookies = driver.get_cookies()
                    session = requests.Session()
                    for cookie in cookies:
                        session.cookies.set(cookie["name"], cookie["value"])
                    self.session = session
                    driver.quit()
                    return True
            except Exception:
                print("验证码错误")
                attempts += 1

        print("多次登录失败，退出程序")
        driver.quit()
        return False

    def gen_text(self, html_content, max_retries=3):
        soup = BeautifulSoup(html_content, 'html.parser')

        # 初始化结果字典
        result = {
            "填表日期": None,
            "项目名称": None,
            "建设地点": None,
            "建筑面积（平方米）": None,
            "建设单位": None,
            "法定代表人": None,
            "联系人": None,
            "联系电话": None,
            "项目投资（万元）": None,
            "环保投资（万元）": None,
            "拟投入生产运营日期": None,
            "建设性质": None,
            "备案依据": None,
            "建设内容及规模": {
                "备案说明": None,
                "备案详情": []
            },
            "主要环境影响": None,
            "采取的环保措施及排放去向": None,
            "备案回执": None
        }

        # 提取并验证填表日期（关键优化点）
        fill_date = soup.find('p', string=lambda t: t and '填表日期' in t)
        if not fill_date:
            print("警告：未找到填表日期，跳过该项目")
            return result, 0  # 视为日期无效，返回状态0

        fill_date_str = fill_date.get_text(strip=True).replace('填表日期：', '')
        result['填表日期'] = fill_date_str

        # 立即检查日期范围并决定是否继续处理
        date_status = self.is_date_in_range(fill_date_str)
        if date_status == 0:  # 不在范围内（-1:早于开始日期, 0:晚于结束日期）
            return {}, 0  # 直接返回，不执行后续解析
        elif date_status == -1:
            return {}, -1
        # 日期在范围内，继续提取其他数据
        retries = 0
        while retries < max_retries:
            try:
                # 提取项目名称
                project_name = soup.find('td', string=lambda t: t and '项目名称' in t)
                result['项目名称'] = project_name.find_next_sibling('td').get_text(strip=True)

                # 提取建设地点
                location = soup.find('td', string=lambda t: t and '建设地点' in t)
                result['建设地点'] = location.find_next_sibling('td').get_text(strip=True)
                result['建筑面积（平方米）'] = location.find_next_sibling('td').find_next_sibling('td').find_next_sibling(
                    'td').get_text(strip=True)

                # 提取建设单位
                company = soup.find('td', string=lambda t: t and '建设单位' in t)
                result['建设单位'] = company.find_next_sibling('td').get_text(strip=True)

                # 提取法定代表人
                representative = soup.find('td', string=lambda t: t and '法定代表人' in t)
                result['法定代表人'] = representative.find_next_sibling('td').get_text(strip=True)

                # 提取联系人
                contact_person = soup.find('td', string=lambda t: t and '联系人' in t)
                result['联系人'] = contact_person.find_next_sibling('td').get_text(strip=True)

                # 提取联系电话
                contact_phone = soup.find('td', string=lambda t: t and '联系电话' in t)
                result['联系电话'] = contact_phone.find_next_sibling('td').get_text(strip=True)

                # 提取项目投资（万元）
                investment = soup.find('td', string=lambda t: t and '项目投资' in t)
                result['项目投资（万元）'] = investment.find_next_sibling('td').get_text(strip=True)

                # 提取环保投资（万元）
                env_investment = soup.find('td', string=lambda t: t and '环保投资' in t)
                result['环保投资（万元）'] = env_investment.find_next_sibling('td').get_text(strip=True)

                # 提取拟投入生产运营日期
                operation_date = soup.find('td', string=lambda t: t and '拟投入生产运营日期' in t)
                result['拟投入生产运营日期'] = operation_date.find_next_sibling('td').get_text(strip=True)

                # 提取建设性质
                nature = soup.find('td', string=lambda t: t and '建设性质' in t)
                result['建设性质'] = nature.find_next_sibling('td').get_text(strip=True)

                # 提取备案依据
                basis = soup.find('td', string=lambda t: t and '备案依据' in t)
                result['备案依据'] = basis.find_next_sibling('td').get_text(strip=True)

                # 提取主要环境影响
                script_tag = soup.find('script', string=re.compile(r'var value="'))
                match = re.search(r'var value="([^"]+)"', script_tag.string)
                value = match.group(1)

                def get_text_from_value(value):
                    num = int(value)  # 确保输入是整数
                    if 8 <= num <= 11:
                        return "固废"
                    elif 12 <= num <= 15:
                        return "噪声"
                    elif 16 <= num <= 19:
                        return "生态影响"
                    elif 20 <= num <= 23:
                        return "电磁辐射"
                    elif 24 <= num <= 27:
                        return "辐射环境影响"
                    else:
                        return "未知"

                result['主要环境影响'] = get_text_from_value(value)

                # 提取采取的环保措施及排放去向
                env_measures = soup.find('input', id='d20val')
                result['采取的环保措施及排放去向'] = env_measures.get('value').replace('。', '').replace(',', '')

                # 提取备案回执
                record_number_strong = soup.find('strong', string=lambda t: t and t.startswith('20'))
                result['备案回执'] = record_number_strong.text if record_number_strong else None

                # 提取建设内容及规模
                content_scale = soup.find('td', string=lambda t: t and '建设内容及规模' in t)
                content_text = content_scale.find_next_sibling('td').get_text(strip=True)
                if '备案详情：' in content_text:
                    # 如果存在“备案详情：”，按此分隔
                    content_parts = content_text.split('备案详情：', 1)
                    result['建设内容及规模']['备案说明'] = content_parts[0].strip() + '。'
                    content_details = content_parts[1].strip()  # 获取备案详情部分
                else:
                    # 如果不存在“备案详情：”，假设整个内容都是备案详情
                    result['建设内容及规模']['备案说明'] = ''  # 备案说明为空
                    content_details = content_text.strip()

                # 按回车符分割备案详情为列表
                content_parts = content_details.split('\n')

                # 分批处理备案详情
                all_details = []
                batch_size = 15  # 每批处理30行数据

                for i in range(0, len(content_parts), batch_size):
                    # 获取当前批次的数据
                    batch = content_parts[i:i + batch_size]
                    batch_text = '\n'.join(batch)

                    # 处理当前批次，最多重试3次
                    max_ai_retries = 3
                    for attempt in range(max_ai_retries):
                        try:

                            # 调用AI接口
                            ai_res = ai_text(
                                batch_text + """处理上述文本,其内部有不同站址的信息，结果以[{'铁塔站址名称':'','铁塔站址编码':''注意：铁塔站址编码必须是45开头的18位纯数字或空值,'位置':'','经度':'','纬度':''}]列表格式返回，且每一条记录必须完整以{}返回,重复的也要保留重复记录，结果只返回json格式列表[]，[]外面不需要其他文字和多余字符""")

                            # 清理AI返回内容
                            cleaned_res = ai_res.strip()
                            # 移除可能的markdown代码块标记
                            if cleaned_res.startswith('```json'):
                                cleaned_res = cleaned_res[len('```json'):].lstrip()
                            elif cleaned_res.startswith('```'):
                                cleaned_res = cleaned_res[len('```'):].lstrip()

                            if cleaned_res.endswith('```'):
                                cleaned_res = cleaned_res[:-len('```')].rstrip()

                            # 确保以[]包裹
                            if not cleaned_res.startswith('['):
                                cleaned_res = '[' + cleaned_res
                            if not cleaned_res.endswith(']'):
                                cleaned_res = cleaned_res + ']'
                            print(cleaned_res)
                            # 解析JSON
                            batch_details = ast.literal_eval(cleaned_res)

                            if isinstance(batch_details, list) and len(batch_details) > 0:
                                all_details.extend(batch_details)
                                print(f"\n第{i // batch_size + 1}批处理成功，获取{len(batch_details)}条记录")
                                break  # 成功则跳出重试循环
                            else:
                                print(f"\n第{i // batch_size + 1}批AI返回无效数据，尝试重试 ({attempt + 1}/{max_ai_retries})")
                                print(
                                    f"解析结果类型: {type(batch_details)}, 长度: {len(batch_details) if isinstance(batch_details, list) else 0}")

                        except Exception as e:
                            print(f"\n第{i // batch_size + 1}批AI处理出错: {e}，尝试重试 ({attempt + 1}/{max_ai_retries})")

                        # 重试前等待1秒
                        time.sleep(1)
                    else:
                        # 所有重试都失败
                        print(f"\n第{i // batch_size + 1}批AI处理失败，已达到最大重试次数")

                result['建设内容及规模']['备案详情'] = all_details
                print(f"\n解析完成，共获取 {len(all_details)} 条备案详情")
                return result, date_status  # 返回完整数据和日期状态
            except Exception as e:
                print(f"解析HTML内容时出错: {e}")
                retries += 1
                if retries < max_retries:
                    print(f"尝试第 {retries + 1} 次解析...")
                else:
                    raise

    def clean_ai_response(self, response):
        """清理AI返回的结果，去除多余的标记和字符"""
        # 移除开始的```json和结尾的```标记
        response = response.strip()
        if response.startswith('```json'):
            response = response[len('```json'):].lstrip()
        elif response.startswith('```'):
            response = response[len('```'):].lstrip()

        if response.endswith('```'):
            response = response[:-len('```')].rstrip()

        # 移除可能存在的其他多余字符或前缀
        response = response.strip()
        return response

    def get_id_list(self, page, max_retries=3):
        for attempt in range(max_retries):
            try:
                print(f"正在爬取第 {page} 页(尝试 {attempt + 1}/{max_retries})...")
                data = {'orderBy': '', 'pageNo': str(page), 'pageSize': '10'}
                res = self.session.post(
                    url="https://beian.china-eia.com/a/registrationform/tBasRegistrationForm/formIndex",
                    headers=self.headers,
                    data=data
                )
                result = {}
                soup = BeautifulSoup(res.text, 'html.parser')
                rows = soup.find_all('tr')
                for row in rows:
                    tds = row.find_all('td')
                    if len(tds) > 6:  # 确保有足够的<td>标签
                        record_number_td = tds[6]
                        # 找到<input>标签并提取备案号
                        input_tag = record_number_td.find('input', class_='recordNumbers')
                        if input_tag and 'value' in input_tag.attrs:
                            record_number = input_tag['value']
                        else:
                            continue  # 如果<input>标签不存在或没有value属性，跳过当前行

                        # 找到查看链接并提取id
                        view_links = row.find_all('a', string='查看')
                        if view_links:
                            view_link = view_links[0]
                            view_id = view_link['href'].split('id=')[-1]
                            result[record_number] = view_id  # 将结果存储在字典中
                        else:
                            continue  # 如果没有找到查看链接，跳过当前行
                return result, 1  # 确保返回两个值：字典和状态码
            except Exception as e:
                print(f"获取第 {page} 页列表时出错: {e}")
                return {}, 0  # 出错时返回空字典和错误状态

    def get_account(self):
        # 读取Excel文件
        df = pd.read_excel(self.account_path)
        df.columns = ['地市', '铁塔_账号', '铁塔_密码', '电信_账号', '电信_密码', '移动_账号', '移动_密码', '联通_账号', '联通_密码']

        # 读取白名单，提取已存在的地市+运营商组合
        if os.path.exists(self.white_list_path):
            white_list_df = pd.read_excel(self.white_list_path, usecols=['地市', '申报方'])
            existing_combinations = set(zip(white_list_df['地市'], white_list_df['申报方']))
        else:
            existing_combinations = set()

        data_list = []

        # 遍历每一行数据
        for index, row in df.iterrows():
            if pd.isna(row['地市']):
                print(f"跳过第 {index + 1} 行，因为地市列为空")
                continue

            for operator in ['铁塔', '电信', '移动', '联通']:
                city = row['地市']
                account = row[f'{operator}_账号']
                password = row[f'{operator}_密码']

                # 如果账号为"暂无"或空值，跳过
                if pd.isna(account) or account == "暂无":
                    continue

                # 检查白名单是否已有该地市+运营商组合
                if (city, operator) in existing_combinations:
                    print(f"跳过 {city}-{operator}，因为白名单中已存在")
                    continue

                # 添加到待处理列表
                new_index = f"{city}-{operator}"
                data_list.append({'索引': new_index, '账号': account, '密码': password})

        result_df = pd.DataFrame(data_list)
        result_df.set_index('索引', inplace=True)
        self.result_df = result_df

    def process_account(self, account_info):
        """处理单个账号的爬取任务"""
        index_name = account_info.name
        username = account_info['账号']
        password = account_info['密码']

        print(f"\n开始处理账号: {index_name} ({username})")

        # 登录
        if not self.login(username, password):
            print(f"账号 {index_name} 登录失败，跳过")
            return

        all_projects_data = []
        continue_flag = True
        page = 1

        # 读取白名单中的备案号
        df = pd.read_excel(self.white_list_path, usecols=['备案号'], dtype=str)
        record_numbers = df['备案号'].dropna().tolist()

        while continue_flag:
            try:
                id_dict, status = self.get_id_list(page)  # 现在只接收一个返回值
                if not id_dict:  # 如果没有获取到数据，可能是最后一页
                    break

                record_count = 0  # 初始化当前页记录计数器
                for record_number_td, obj_id in id_dict.items():
                    if record_number_td in record_numbers:
                        continue

                    record_count += 1  # 记录计数器递增
                    res = self.session.get(
                        url=f"https://beian.china-eia.com/a/registrationform/tBasRegistrationForm/viewfront?id={obj_id}",
                        headers=self.headers,
                    )
                    df_data, date_flag = self.gen_text(res.text)

                    # 修改后的输出语句，显示当前页的第几条记录
                    print(f"第{page}页的第{record_count}条记录读取成功")

                    if date_flag == -1:
                        print(f"项目 {obj_id} 的填表日期早于开始日期，停止爬取")
                        continue_flag = False
                        break
                    elif date_flag == 0:
                        print(f"项目 {obj_id} 的填表日期晚于结束日期，跳过此项目")
                        continue

                    all_projects_data.append(df_data)
                page += 1
            except Exception as e:
                print(f"处理第 {page} 页时出错: {e}")

        if all_projects_data:
            # 创建一个空列表来存储处理后的所有行数据
            expanded_rows = []

            for project in all_projects_data:
                # 提取共同数据（除了备案详情）
                common_data = {k: v for k, v in project.items() if k != '建设内容及规模'}

                # 提取备案说明
                if '建设内容及规模' in project and '备案说明' in project['建设内容及规模']:
                    common_data['备案说明'] = project['建设内容及规模']['备案说明']
                else:
                    common_data['备案说明'] = None

                # 处理备案详情中的每个字典
                if '建设内容及规模' in project and '备案详情' in project['建设内容及规模']:
                    for detail in project['建设内容及规模']['备案详情']:
                        # 合并共同数据和详情数据
                        row = {**common_data, **detail}
                        expanded_rows.append(row)
                else:
                    # 如果没有备案详情，只添加共同数据
                    expanded_rows.append(common_data)

            # 创建DataFrame
            df = pd.DataFrame(expanded_rows)

            # 重新排列列顺序（可选）
            columns_order = ['填表日期', '项目名称', '建设地点', '建筑面积（平方米）',
                             '建设单位', '法定代表人', '联系人', '联系电话',
                             '项目投资（万元）', '环保投资（万元）', '拟投入生产运营日期',
                             '建设性质', '备案依据',
                             '铁塔站址名称', '铁塔站址编码', '位置', '经度', '纬度',
                             '主要环境影响', '采取的环保措施及排放去向', '备案回执']

            # 只保留实际存在的列
            columns_order = [col for col in columns_order if col in df.columns]
            df = df[columns_order]

            output_filename = os.path.join(self.output_path, f"{index_name}.xlsx")
            df.to_excel(output_filename, index=False)
            print(f"账号 {index_name} 处理完成，结果保存到 {output_filename}")
        else:
            print(f"账号 {index_name} 没有获取到有效数据")

    def run_thread(self):
        self.get_account()
        # 遍历所有账号
        for index, row in self.result_df.iterrows():
            self.process_account(row)
            # 每次处理完一个账号后等待一段时间
            time.sleep(5)

if __name__ == "__main__":
    main().run_thread()