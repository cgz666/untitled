import os
import time
from bs4 import BeautifulSoup
import requests
from selenium import webdriver
import re
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pickle

class WeatherSpider:
    def __init__(self):
        INDEX = os.getcwd()
        self.data_text = None
        self.base_url = "http://gx.weather.com.cn"
        self.weather_url = f"{self.base_url}/weather.shtml"
        self.session = requests.Session()
        self.options = Options()
        self.options.add_argument('--headless')  # 无头模式
        self.options.add_argument('--disable-gpu')
        self.options.add_argument('--no-sandbox')
        self.chrome_path = os.path.join(INDEX, "chrome-win64/chrome.exe")
        self.options.binary_location = self.chrome_path  # 设置浏览器路径
        self.driver_path = os.path.join(INDEX, "chromedriver.exe")
        self.driver = webdriver.Chrome(service=Service(self.driver_path), options=self.options)
        self.headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0"
        }
        self.city_list = ["南宁市", "柳州市", "桂林市", "梧州市", "北海市", "防城港市", "钦州市", "贵港市", "玉林市", "百色市", "贺州市", "河池市", "来宾市", "崇左市"]

    def parse_land_forecast(self, html_content):
        if not html_content:
            print("没有获取到页面内容")
            return None

        soup = BeautifulSoup(html_content, 'html.parser')
        dl_tag = soup.find('dl', id='mainContent')
        h3_tag = soup.find('h3')
        if h3_tag:
            date_text = h3_tag.get_text(strip=True)
            date_text = date_text.replace('【字体：大中小】', '').strip()
            date_text = date_text.split('来源：')[0].strip()
            self.data_text = date_text
        if dl_tag:
            p_tags = dl_tag.find_all('p')
            guangxi_cities = ["南宁", "柳州", "桂林", "梧州", "北海", "防城港", "钦州", "贵港", "玉林", "百色", "贺州", "河池", "来宾", "崇左"]
            for p in p_tags:
                text = p.get_text(strip=True)
                # 去除换行符
                text = text.replace('\n', '').replace('\r', '')
                if text.startswith("今天") and any(city in text for city in guangxi_cities):
                    return text
            print("未找到目标段落")
        else:
            print("未找到广西陆地天气预报信息")
        return None

    def parse_weather_alerts(self, url):
        if not url.startswith('http'):
            url = f"{self.base_url}/{url}"
        self.driver.get(url)
        time.sleep(5)  # 增加等待时间确保动态内容加载

        try:
            alert_list = self.driver.find_element(By.CLASS_NAME, "alarml")
            alert_items = alert_list.find_elements(By.TAG_NAME, "li")

            alerts_data = []

            for item in alert_items:
                try:
                    alert_text = item.text.strip()
                    if alert_text:
                        datetime_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}', alert_text)
                        if datetime_match:
                            datetime_str = datetime_match.group()
                        else:
                            datetime_str = "未知时间"
                        location_match = re.search(r'广西壮族自治区[^发布]+', alert_text)
                        if location_match:
                            location_str = location_match.group()
                        else:
                            location_str = "未知地区"

                        alert_type_match = re.search(r'发布[^预警]+预警', alert_text)
                        if alert_type_match:
                            alert_type_str = alert_type_match.group()
                        else:
                            alert_type_str = "未知预警"
                        final_alert = f"{datetime_str}，{location_str}发布{alert_type_str}"
                        alerts_data.append(final_alert)
                except Exception:
                    raise
            return alerts_data
        except Exception:
            raise

    def run(self):
        try:
            response = self.session.get(url=self.weather_url, headers=self.headers)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            target_div1 = soup.find('div', class_='indextopnews')
            link1 = target_div1.find('a') if target_div1 else None

            land_forecast = None
            if link1:
                link1_href = link1.get('href')
                details_url = f"{self.base_url}/{link1_href}"
                response1 = self.session.get(url=details_url, headers=self.headers)
                response1.raise_for_status()
                res_content1 = response1.content.decode('UTF-8', errors='ignore')
                land_forecast = self.parse_land_forecast(res_content1)
                if land_forecast:
                    print("陆地天气预报内容:", self.data_text+','+land_forecast)
                else:
                    print("未找到陆地天气预报内容")
            target_div2 = soup.find('div', class_='bottom indexbottom')
            link2 = target_div2.find('a') if target_div2 else None
            alerts = []
            if link2:
                link2_href = link2.get('href')
                alerts = self.parse_weather_alerts(link2_href)

            data = {
                "广西": self.data_text + ',' + land_forecast,
                "其他地区": []  # 初始化 "其他地区" 键
            }
            # 为self.city_list中的每个城市初始化一个空列表
            for city in self.city_list:
                data[city] = []

            # 处理预警数据
            for alert in alerts:
                matched = False
                for city in self.city_list:
                    if city in alert:
                        data[city].append(alert)
                        matched = True
                        break
                if not matched:
                    data["其他地区"].append(alert)

            if alerts:
                for alert in alerts:
                    print("天气预警内容:", alert)
            else:
                print("未找到天气预警内容")

            self.save_to_pickle(data)

        finally:
            if hasattr(self, 'driver'):
                self.driver.quit()
                print("\n浏览器已关闭")

    def save_to_pickle(self, data, filename='weather_data.pkl'):
        with open(filename, 'wb') as f:
            pickle.dump(data, f)
        print(f"数据已保存到 {filename}")

if __name__ == "__main__":
    spider = WeatherSpider()
    spider.run()