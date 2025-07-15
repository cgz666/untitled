import requests
import win32com.client as win32
import os
import shutil
import psutil
import pythoncom
import time
from urllib.parse import unquote
from datetime import datetime, timedelta
from contextlib import contextmanager
from typing import Dict, List, Optional
from seleniumwire import webdriver
from selenium.webdriver.chrome.service import Service
from urllib3.exceptions import ReadTimeoutError
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from utils.ai import ai_yzm

class ExcelDataProcessor():
    """Excel数据处理与爬虫程序"""
    def __init__(self):
        """初始化处理器"""
        INDEX = os.getcwd()
        self.output_path = os.path.join(INDEX,"output")
        self.driver_path =os.path.join(INDEX,"chromedriver.exe")
        self.chrome_path=os.path.join(INDEX,"chrome-win64/chrome.exe")
        # 文件路径配置
        self.save_path = os.path.join(INDEX, "xls")
        self.output_path = os.path.join(INDEX, "output")
        self.output_name1 = os.path.join(self.output_path, "委托备案数据与环保厅备案数据匹配情况统计-结果.xlsx")
        self.output_name2 = os.path.join(self.output_path, "委托备案数据与环保厅备案数据匹配情况统计-结果_筛选后.xlsx")
        self.file_name1 = os.path.join(self.save_path, "output1.xlsx")
        self.file_name2 = os.path.join(self.save_path, "output2.xlsx")
        self.file_name3 = os.path.join(self.save_path, "output3.xlsx")
        self.model_path = os.path.join(self.save_path, "模板.xlsx")

    def init_driver(self):
        """初始化浏览器驱动"""
        service = Service(executable_path=self.driver_path)
        option = webdriver.ChromeOptions()
        option.add_argument("--disable-gpu")
        option.add_argument("--no-sandbox")
        option.add_argument("--disable-blink-features=Automation")  # 禁用自动化检测
        option.add_argument("--disable-infobars")  # 禁用信息栏
        option.add_experimental_option("excludeSwitches", ["enable-automation"])  # 禁用自动化
        option.add_argument("--ignore-certificate-errors")  # 禁用证书错误检查
        option.add_argument('--allow-running-insecure-content')  # 允许不安全内容
        option.add_argument('--allow-insecure-localhost')
        option.add_argument(
            r'--user-data-dir=C:\Users\27569\AppData\Local\Google\Chrome for Testing\User Data')
        option.add_experimental_option("prefs", {
            "download.default_directory": self.output_path,  # 设置下载路径为目标路径
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
        option.binary_location = self.chrome_path
        return webdriver.Chrome(service=service, options=option)

    def login(self):
        driver = self.init_driver()
        try:
            login_url = "https://cloud.gxtower.cn:8012/fronts/tower_fe/login"
            driver.get(login_url)
            WebDriverWait(driver, 3).until(EC.title_contains("基站环境保护全流程系统"))
            max_attempts = 10
            attempts = 0
            while attempts < max_attempts:
                try:
                    src = driver.find_element(By.CLASS_NAME, "identifying-code-img").get_attribute("src")
                    src = unquote(src)
                    res = ai_yzm(src)
                    print(f"验证码识别结果: {res}")
                    username_input = driver.find_element(By.ID, "userName")
                    password_input = driver.find_element(By.ID, "password")
                    captcha_input = driver.find_element(By.ID, "vcode")
                    login_button = driver.find_element(By.CLASS_NAME, "ant-btn-primary")
                    username_input.send_keys("gxtt")
                    password_input.send_keys("Fhf6Ea")
                    captcha_input.send_keys(res)
                    login_button.click()
                    time.sleep(3)  # 等待5秒
                    if driver.current_url == login_url:
                        print("登录失败")
                        driver.refresh()
                        WebDriverWait(driver, 5).until(EC.title_contains("基站环境保护全流程系统"))
                        time.sleep(2)
                        attempts += 1
                    else:
                        print("登录成功！")
                        break
                except:
                    time.sleep(10)
            for req in driver.requests:
                if 'https://cloud.gxtower.cn:8012/towerGateway/towerSystem/sysResource/listMenuTreesByCurrentUser' in req.url:
                    self.token = req.headers.get('token')
                    break
        except Exception as e:
            raise
        finally:
            driver.quit()

    def clear_pywin32_cache(self):
        """清除pywin32缓存"""
        cache_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", "gen_py")
        if os.path.exists(cache_dir):
            print(f"清除pywin32缓存: {cache_dir}")
            try:
                shutil.rmtree(cache_dir)
                print("缓存已清除")
            except Exception as e:
                print(f"清除缓存失败: {e}")
        else:
            print("未找到pywin32缓存目录")
    @contextmanager
    def excel_application(self) -> win32.CDispatch:
        """安全管理Excel应用程序对象的上下文管理器"""
        excel = None
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.Interactive = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            yield excel
        finally:
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
                del excel
            self.kill_excel_processes()
    def kill_excel_processes(self) -> None:
        """仅终止由当前脚本创建的Excel进程"""
        current_pid = os.getpid()
        try:
            for proc in psutil.process_iter(['name', 'ppid']):
                if proc.info['name'] == 'EXCEL.EXE' and proc.info['ppid'] == current_pid:
                    proc.kill()
        except Exception as e:
            print(f"清理Excel进程时出错: {e}")
    def fetch_and_save(self, url: str, headers: Dict[str, str], method: str = 'GET',
                       payload: Optional[Dict] = None, output_file: Optional[str] = None) -> Optional[
        requests.Response]:
        """
        通用的网络请求并保存响应内容的函数

        Args:
            url: 请求URL
            headers: 请求头
            method: 请求方法，默认为GET
            payload: 请求载荷，默认为None
            output_file: 保存的文件路径
        """
        try:
            # 创建会话对象
            with requests.Session() as session:
                session.headers.update(headers)

                print(f"开始请求: {url}")
                start_time = time.time()

                # 根据请求方法发送请求
                if method.upper() == 'GET':
                    response = session.get(url)
                elif method.upper() == 'POST':
                    response = session.post(url, json=payload)
                else:
                    raise ValueError(f"不支持的请求方法: {method}")

                # 检查响应状态码
                response.raise_for_status()

                # 保存响应内容
                if output_file:
                    with open(output_file, 'wb') as file:
                        file.write(response.content)
                    print(f"成功保存文件到: {output_file}")
                else:
                    print("请求成功，但未指定输出文件")

                end_time = time.time()
                print(f"请求完成，耗时: {end_time - start_time:.2f}秒")

                return response

        except:
            raise
        return None

    def spider1(self) -> None:
        if not self.token:
            self.token = self.login()
            if not self.token:
                print("未获取到token，爬虫任务取消")
                return
        """爬取第一张表"""
        url = "https://cloud.gxtower.cn:8012/towerGateway/towerBiz/dataList/exportData?args=%7B%22cfgName%22%3A%22custInfoComfirm%22%2C%22equal%22%3A%7B%7D%2C%22like%22%3A%7B%7D%2C%22range%22%3A%7B%7D%7D"
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Cookie": "sssct=00; sssty=4",
            "Host": "cloud.gxtower.cn:8012",
            "Referer": "https://cloud.gxtower.cn:8012/fronts/tower_fe/standingBook?cfgName=custInfoComfirm",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
            "sec-ch-ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": '"Windows"',
            "token": f"{self.token}"
        }

        self.fetch_and_save(url, headers, method='GET', output_file=self.file_name1)

    def spider2(self) -> None:
        if not self.token:
            if not self.token:
                print("未获取到token，爬虫任务取消")
                return
        url = "https://cloud.gxtower.cn:8012/towerGateway/towerBiz/tower/hpMonitorManage/exportBeiAnQuery"
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br, zstd",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Connection": "keep-alive",
            "Content-Type": "application/json;charset=UTF-8",
            "Cookie": "sssct=00; sssty=4",
            "Host": "cloud.gxtower.cn:8012",
            "Origin": "https://cloud.gxtower.cn:8012",
            "Referer": "https://cloud.gxtower.cn:8012/fronts/tower_fe/qurey/ectrust",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
            "sec-ch-ua": "\"Chromium\";v=\"136\", \"Microsoft Edge\";v=\"136\", \"Not.A/Brand\";v=\"99\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\"",
            "token": f"{self.token}"
        }

        self.fetch_and_save(url, headers, method='POST', payload={}, output_file=self.file_name2)

    def spider3(self):
        driver = self.init_driver()
        try:
            login_url = "http://116.11.253.94:8180/tower/sys/datacenter/Login"
            driver.get(login_url)
            username_input = driver.find_element(By.NAME, "account")
            password_input = driver.find_element(By.NAME, "pwd")
            login_button = driver.find_element(By.CLASS_NAME, "submit")
            wait = WebDriverWait(driver, 10)
            username_input.send_keys("qujkzx")
            password_input.send_keys("qujkzx123456")
            login_button.click()

            maintenance_menu = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//li[contains(@class, "ant-menu-submenu") and .//span[text()="维护台账"]]')))
            actions = ActionChains(driver)
            actions.move_to_element(maintenance_menu).perform()
            target_menu_item = wait.until(
                EC.element_to_be_clickable((By.XPATH,'//li[contains(@class, "ant-menu-item") and text()="环保厅备案单站信息"]')))
            target_menu_item.click()
            time.sleep(2)

            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if iframes:
                driver.switch_to.frame(iframes[0])

            try:
                max_retries = 3
                retries = 0
                while retries < max_retries:
                    try:
                        export_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable(
                                (By.XPATH,
                                 '//button[.//span[text()="导出"]] | //button[.//i[contains(@class,"anticon-download")]]')
                            )
                        )
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_button)
                        driver.execute_script("arguments[0].click();", export_button)
                        print("点击导出成功")
                        break
                    except (ReadTimeoutError, WebDriverException) as e:
                        retries += 1
                        if retries >= max_retries:
                            raise
                        time.sleep(10 * retries)

                target_file_path = self.file_name3
                target_keyword = "环保厅备案单站信息"  # 文件名关键词
                timeout = 600
                start_time = time.time()

                while True:
                    if time.time() - start_time > timeout:
                        raise Exception(f"下载超时，未找到包含'{target_keyword}'的文件")

                    all_files = os.listdir(self.output_path)
                    matched_files = [f for f in all_files if target_keyword in f and f.endswith('.xlsx')]

                    if matched_files:
                        source_file = matched_files[0]
                        source_path = os.path.join(self.output_path, source_file)

                        print(f"找到匹配文件: {source_file}")

                        previous_size = -1
                        for _ in range(5):
                            current_size = os.path.getsize(source_path)
                            if current_size == previous_size:
                                break
                            previous_size = current_size
                            time.sleep(1)

                        shutil.move(source_path, target_file_path)
                        break
                    else:
                        time.sleep(2)
            except Exception as e:
                print(f"导出操作失败: {e}")
                raise
        finally:
            driver.quit()

    def process_excel_file(self) -> None:
        """处理Excel文件，将公告日期转换为短日期格式(不带前导零)，并保存到原文件"""
        self.clear_pywin32_cache()
        self.kill_excel_processes()

        # 只使用原文件路径，不创建新文件
        files_to_process = {
            self.file_name1: '公告日期',
            self.file_name2: '公告日期',
            self.file_name3: '公告日期'
        }

        # 工作表名及对应列映射
        sheet_column_mapping = {
            self.file_name1: ("export", "Y"),
            self.file_name2: ("export", "R"),
            self.file_name3: ("sheet1", "P")
        }

        # columns_to_delete = ['起租操作时间', '协议变更时间', '是否总部管控（人工填写）','监测日期']

        for input_file, date_column in files_to_process.items():
            if not os.path.exists(input_file):
                print(f"输入文件 {input_file} 不存在，跳过处理")
                continue

            print(f"\n开始处理文件: {os.path.basename(input_file)}")
            print(f"日期列: {date_column}")
            print("⚠️ 注意：将直接修改原文件并覆盖保存")

            max_retries = 3
            for attempt in range(max_retries):
                try:
                    with self.excel_application() as excel:
                        wb = excel.Workbooks.Open(input_file, ReadOnly=False)

                        # 获取对应工作表和列
                        sheet_name, column_letter = sheet_column_mapping.get(input_file, ("Sheet1", "A"))
                        try:
                            ws = wb.Worksheets(sheet_name)
                        except:
                            print(f"无法找到工作表 '{sheet_name}'，使用活动工作表")
                            ws = wb.ActiveSheet

                        ws.DisplayPageBreaks = False

                        # 删除指定列
                        # if input_file == self.file_name1 and columns_to_delete:
                        #     print(f"正在删除列: {', '.join(columns_to_delete)}")
                        #     header_row = ws.Rows(1)
                        #     col_indices = {
                        #         header_row.Cells(col).Value: col
                        #         for col in range(1, ws.UsedRange.Columns.Count + 1)
                        #     }
                        #     cols_to_delete = [
                        #         col_indices[col_name] for col_name in columns_to_delete
                        #         if col_name in col_indices
                        #     ]
                        #     for col_idx in sorted(cols_to_delete, reverse=True):
                        #         ws.Columns(col_idx).Delete()
                        #     print(f"已成功删除 {len(cols_to_delete)} 列")

                        # 定位日期列
                        header_values = [ws.Cells(1, col).Value for col in range(1, ws.UsedRange.Columns.Count + 1)]
                        if date_column not in header_values:
                            print(f"未找到列 '{date_column}'，跳过处理")
                            break

                        date_col_idx = header_values.index(date_column) + 1
                        last_row = ws.UsedRange.Rows.Count
                        if last_row <= 1:
                            print("没有数据可处理")
                            break

                        # 处理日期列数据（不再分块）
                        date_range = ws.Range(
                            f"{ws.Cells(2, date_col_idx).Address}:{ws.Cells(last_row, date_col_idx).Address}")
                        date_values = date_range.Value

                        processed_values = []
                        for value in date_values:
                            val = value[0]
                            # 跳过空值
                            if val is None or val == '':
                                processed_values.append('')
                                continue

                            date_str = None

                            if isinstance(val, (float, int)):
                                date_obj = datetime(1899, 12, 30) + timedelta(days=val)
                                date_str = f"{date_obj.year}/{date_obj.month}/{date_obj.day}"
                            elif isinstance(val, str):
                                try:
                                    if '-' in val:
                                        date_obj = datetime.strptime(val.split()[0], '%Y-%m-%d')
                                    elif '/' in val:
                                        date_obj = datetime.strptime(val.split()[0], '%Y/%m/%d')
                                    else:
                                        date_str = val
                                        continue
                                    date_str = f"{date_obj.year}/{date_obj.month}/{date_obj.day}"
                                except:
                                    date_str = val
                                    continue

                            if date_str:
                                processed_values.append(f"'{date_str}")
                            else:
                                processed_values.append(val)

                        # 一次性更新所有值
                        if processed_values:
                            date_range.Value = [[v] for v in processed_values]

                        # 显示进度
                        print(f"  进度: {last_row - 1}/{last_row - 1} 行 (100.0%)")

                        # 完成日期处理后执行分列操作
                        last_data_row = ws.Cells(ws.Rows.Count, column_letter).End(win32.constants.xlUp).Row
                        if last_data_row > 1:
                            # 选择对应列数据范围
                            column_range = ws.Range(f"{column_letter}2:{column_letter}{last_data_row}")

                            print(f"对 {sheet_name} 工作表的 {column_letter} 列执行分列操作")
                            column_range.TextToColumns(
                                Destination=column_range,
                                DataType=win32.constants.xlFixedWidth,
                                FieldInfo=[(0, 1)],  # 只设置一个字段
                                DecimalSeparator=".",
                                TrailingMinusNumbers=True
                            )

                        # 保存到原文件路径
                        wb.Save()
                        print(f"✔️ 文件处理完成并已保存到原文件: {input_file}")
                        break

                except Exception as e:
                    print(f"尝试 {attempt + 1}/{max_retries} 失败: {e}")
                    if attempt == max_retries - 1:
                        print(f"❌ 处理文件失败: {input_file}")
                        raise
                finally:
                    time.sleep(0.5)

        print("\n所有文件处理完成!")

    def excel_process(self):
        """
        处理Excel文件，将指定文件夹中的数据文件内容复制到主表文件中。

        :param index_path: 文件夹路径
        """
        print('1、把数据文件和通报模板放在同一文件夹下')
        print('2、打开上述文件，如果提示保护视图则取消（报错大概率是这个问题），如果提示别的东西请点击掉，保证程序能够编辑文档')
        # 初始化 COM 库
        pythoncom.CoInitialize()
        try:
            xl = win32.gencache.EnsureDispatch('Excel.Application')  # 开启excel软件
            xl.Visible = True
            xl.DisplayAlerts = False

            # 打开模板文件
            workbook_main = xl.Workbooks.Open(self.model_path)

            # === 处理已起租去重(站址编码+运营商) ===
            workbook_data = xl.Workbooks.Open(self.file_name1)
            sheet_data = workbook_data.Sheets('export')
            sheet_main = workbook_main.Sheets('已起租去重(站址编码+运营商)')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:AI{last_row}')

            # 1. 只清除A-AI列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"A2:AI{last_clear_row}").ClearContents()

            # 2. 复制新数据到 A-AI 列
            source_range.Copy()
            sheet_main.Range('A2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False
            xl.CalculateFull()

            # 新增：设置Y列（公告日期）的格式
            if last_row > 1:
                date_range = sheet_main.Range(f'Y2:Y{last_row}')
                date_range.NumberFormat = 'yyyy/m/d'

            workbook_data.Close(SaveChanges=False)
            # 3. 使用AutoFill填充AJ-AN列的数据到last_row
            if last_row > 2:  # 确保有足够的行进行填充
                # 获取AJ-AN列的第二行数据范围
                fill_source = sheet_main.Range('AJ2:AN2')
                # 确定目标范围（AJ-AN列从第2行到last_row行）
                fill_target = sheet_main.Range(f'AJ2:AN{last_row}')
                # 使用AutoFill方法填充数据
                fill_source.AutoFill(Destination=fill_target, Type=win32.constants.xlFillDefault)
            time.sleep(3)

            # === 委托备案（公告日期剔除1990年的信息） ===
            workbook_data = xl.Workbooks.Open(self.file_name2)
            sheet_data = workbook_data.Sheets('export')
            sheet_main = workbook_main.Sheets('委托备案（公告日期剔除1990年的信息）')

            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:CM{last_row}')

            # 1. 只清除A-CM列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"A2:CM{last_clear_row}").ClearContents()

            # 2. 复制新数据到 A-P 列
            source_range.Copy()
            sheet_main.Range('A2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False
            # 重新计算工作表以确保公式更新
            xl.CalculateFull()
            workbook_data.Close(SaveChanges=False)
            # 3. 使用AutoFill填充CN-CP列的数据到last_row
            if last_row > 2:  # 确保有足够的行进行填充
                # 获取AF-AK列的第二行数据范围
                fill_source = sheet_main.Range('CN2:CP2')
                # 确定目标范围（AF-AK列从第2行到last_row行）
                fill_target = sheet_main.Range(f'CN2:CP{last_row}')
                # 使用AutoFill方法填充数据
                fill_source.AutoFill(Destination=fill_target, Type=win32.constants.xlFillDefault)
            time.sleep(3)

            # === 环保厅 ===
            workbook_data = xl.Workbooks.Open(self.file_name3)
            sheet_data = workbook_data.Sheets('sheet1')
            sheet_main = workbook_main.Sheets('环保厅')
            # 动态获取数据的实际范围
            last_row = sheet_data.Cells(sheet_data.Rows.Count, 1).End(win32.constants.xlUp).Row
            source_range = sheet_data.Range(f'A2:AG{last_row}')  # 从A2开始复制
            # 1. 只清除A-AG列的数据（不包含第一行表头）
            last_clear_row = sheet_main.UsedRange.Rows.Count
            if last_clear_row > 1:
                sheet_main.Range(f"A2:AG{last_clear_row}").ClearContents()
            # 复制和粘贴
            source_range.Copy()
            sheet_main.Range('A2').PasteSpecial(Paste=win32.constants.xlPasteValues)
            xl.CutCopyMode = False  # 释放剪切板
            workbook_data.Close(SaveChanges=False)
            # 3. 使用AutoFill填充AH-AJ列的数据到last_row
            if last_row > 2:  # 确保有足够的行进行填充
                # 获取AF-AK列的第二行数据范围
                fill_source = sheet_main.Range('AH2:AJ2')
                # 确定目标范围（AF-AK列从第2行到last_row行）
                fill_target = sheet_main.Range(f'AH2:AJ{last_row}')
                # 使用AutoFill方法填充数据
                fill_source.AutoFill(Destination=fill_target, Type=win32.constants.xlFillDefault)
            time.sleep(3)

            workbook_main.SaveAs(self.output_name1)
            workbook_main.Close()
            xl.Quit()  # 关闭Excel应用程序
            print('已全部完成')
        except Exception as e:
            raise
        finally:
            # 释放 COM 库
            pythoncom.CoUninitialize()

    def filter_and_statistics(self):

        print("开始筛选数据并生成最终结果文件...")

        # 初始化Excel应用
        pythoncom.CoInitialize()
        excel = None

        try:
            # 创建Excel应用程序
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.DisplayAlerts = False  # 不显示警告

            print(f"正在打开源文件: {self.output_name1}")
            workbook = excel.Workbooks.Open(self.output_name1)

            # 创建新工作簿用于保存结果
            new_workbook = excel.Workbooks.Add()

            # ===== 处理'委托备案'工作表 =====
            print("开始处理'委托备案'工作表...")
            try:
                source_sheet = workbook.Sheets('委托备案（公告日期剔除1990年的信息）')
                new_sheet = new_workbook.Sheets.Add(Before=new_workbook.Sheets(1))
                new_sheet.Name = '委托备案筛选'

                # 复制表头（第一行）
                source_sheet.Rows(1).Copy()
                new_sheet.Rows(1).PasteSpecial(Paste=win32.constants.xlPasteAll)

                # 查找"是否匹配"列(CO列)
                header_row = source_sheet.Rows(1)
                co_col = None
                for col in range(1, source_sheet.UsedRange.Columns.Count + 1):
                    if header_row.Cells(1, col).Value == "是否匹配":
                        co_col = col
                        break

                if co_col is None:
                    raise Exception("错误：在工作表中找不到'是否匹配'列")

                # 筛选并复制"不匹配"的行
                row_count = 0
                total_rows = source_sheet.UsedRange.Rows.Count
                print(f"共发现 {total_rows} 行数据，开始筛选...")

                for row in range(2, total_rows + 1):
                    # 显示进度
                    if row % 100 == 0:
                        print(f"正在处理第 {row}/{total_rows} 行...")

                    cell_value = source_sheet.Cells(row, co_col).Value
                    if cell_value == "不匹配":
                        row_count += 1
                        # 复制整行数据和格式
                        source_sheet.Rows(row).Copy()
                        # 粘贴到新工作表（只粘贴值和格式）
                        new_sheet.Rows(row_count + 1).PasteSpecial(Paste=win32.constants.xlPasteValuesAndNumberFormats)
                        # 粘贴列宽和格式
                        new_sheet.Rows(row_count + 1).PasteSpecial(Paste=win32.constants.xlPasteFormats)

                print(f"从'委托备案'工作表中筛选出 {row_count} 条'不匹配'记录")

                # 调整列宽匹配原表
                for col in range(1, source_sheet.UsedRange.Columns.Count + 1):
                    new_sheet.Columns(col).ColumnWidth = source_sheet.Columns(col).ColumnWidth

            except Exception as e:
                print(f"处理'委托备案'工作表时出错: {e}")
                raise

            # ===== 处理'统计'工作表 =====
            try:
                source_sheet = workbook.Sheets('统计')
                new_sheet = new_workbook.Sheets.Add(Before=new_workbook.Sheets(1))
                new_sheet.Name = '统计'

                # 复制整个工作表（带格式）
                source_sheet.UsedRange.Copy()

                # 粘贴值和数字格式
                new_sheet.Range("A1").PasteSpecial(Paste=win32.constants.xlPasteValuesAndNumberFormats)
                # 粘贴格式（包括字体、颜色等）
                new_sheet.Range("A1").PasteSpecial(Paste=win32.constants.xlPasteFormats)

                # 调整列宽匹配原表
                for col in range(1, source_sheet.UsedRange.Columns.Count + 1):
                    new_sheet.Columns(col).ColumnWidth = source_sheet.Columns(col).ColumnWidth

            except Exception as e:
                print(f"处理'统计'工作表时出错: {e}")
                raise

            # 删除默认的Sheet1
            try:
                for sheet in new_workbook.Sheets:
                    if sheet.Name == "Sheet1":
                        sheet.Delete()
                        break
            except:
                pass

            # 保存新工作簿
            print(f"\n正在保存结果到: {self.output_name2}")
            new_workbook.SaveAs(self.output_name2)
            print("文件保存成功！")

        except Exception as e:
            print(f"\n错误：筛选数据时发生异常 - {e}")
            raise
        finally:
            # 清理资源
            if 'workbook' in locals():
                workbook.Close(False)
            if 'new_workbook' in locals():
                new_workbook.Close(True)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()
            self.kill_excel_processes()
            print("所有操作已完成！")


    def main(self):
        print("开始执行爬虫程序...")
        self.login()
        self.spider1()
        # time.sleep(2)
        # self.spider2()
        # time.sleep(2)
        # self.spider3()
        # self.process_excel_file()

        # 合并到模板文件
        self.excel_process()

        # # 导出需要的数据
        self.filter_and_statistics()

        print("程序执行完毕!")

if __name__ == "__main__":
    ExcelDataProcessor().main()