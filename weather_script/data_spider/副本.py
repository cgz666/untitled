import os
import time
from bs4 import BeautifulSoup
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pickle


class WeatherSpider:
    def __init__(self):
        self.base_url = "http://gx.weather.com.cn"
        self.weather_url = f"{self.base_url}/weather.shtml"
        self.session = requests.Session()
        self.options = Options()
        self.options.add_argument('--headless')  # 无头模式
        self.options.add_argument('--disable-gpu')
        self.options.add_argument('--no-sandbox')

        # 设置浏览器可执行文件路径
        INDEX = os.getcwd()
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

    def parse_land_forecast(self, html_content):
        if not html_content:
            print("没有获取到页面内容")
            return None

        soup = BeautifulSoup(html_content, 'html.parser')
        dl_tag = soup.find('dl', id='mainContent')

        if dl_tag:
            p_tags = dl_tag.find_all('p')
            for p in p_tags:
                text = p.get_text(strip=True)
                if text.startswith("今天"):
                    return text
        else:
            print("未找到id为mainContent的dl标签")

        return None

    def parse_weather_alerts(self, url):
        if not url.startswith('http'):
            url = f"{self.base_url}/{url}"
        self.driver.get(url)
        time.sleep(3)
        self.driver.implicitly_wait(5)
        alert_elements = self.driver.find_elements(By.XPATH,
                                                   "//*[contains(text(), '广西壮族自治区') and contains(text(), '预警')]")
        alerts = [element.text for element in alert_elements if element.text.strip()]
        return alerts

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
                    print("陆地天气预报内容:", land_forecast)
                else:
                    print("未找到陆地天气预报内容")

            target_div2 = soup.find('div', class_='bottom indexbottom')
            link2 = target_div2.find('a') if target_div2 else None
            alerts = []
            if link2:
                link2_href = link2.get('href')
                alerts = self.parse_weather_alerts(link2_href)

            data = {
                "广西": land_forecast,
            }
            # 提取城市名称并添加到字典
            for alert in alerts:
                if "广西壮族自治区" in alert:
                    city = alert.split("广西壮族自治区")[1].split("发布")[0].strip()
                    data[city] = alert
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