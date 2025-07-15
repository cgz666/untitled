import datetime
from config import SPIDER_PATH
from websource.spider.down_foura import down_model
from datetime import datetime,timedelta
from websource.universal_config import config_serch_performence
from websource.spider.down_foura import down_model
import os
import pandas as pd
from datetime import datetime,timedelta
from websource.universal_config import config_serch_performence_his
from websource.spider.down_foura import down_model
import requests
import os
from config import INDEX
from bs4 import BeautifulSoup

class down_performence():
    # 路径：运行监控-性能查询-查询-导出
    def __init__(self):
        CONFIG = config_serch_performence
        self.data = {
            '1': CONFIG.INTO_DATA1,
            '2': CONFIG.INTO_DATA2,
            'FINAL': CONFIG.INTO_DATA_FINAL,
        }
        self.URL = CONFIG.URL

    def down(self,cities,search_id,folder_temp):
        for city in cities:
            for key in ['1','2']:
                # 当前时间
                now = datetime.now()
                # 一天前的时间
                one_day_ago = now - timedelta(days=1)

                # 格式时间化
                end_time_input_date = now.strftime("%Y-%m-%d %H:%M")
                start_time_input_date = one_day_ago.strftime("%Y-%m-%d %H:%M")
                end_time_input_current_date = now.strftime("%m/%Y")
                start_time_input_current_date = one_day_ago.strftime("%m/%Y")

                self.data[key]['queryForm:endtimeInputCurrentDate'] = end_time_input_current_date
                self.data[key]['queryForm:endtimeInputDate'] = end_time_input_date
                self.data[key]['queryForm:starttimeInputCurrentDate'] = start_time_input_current_date
                self.data[key]['queryForm:starttimeInputDate'] = start_time_input_date
                self.data[key]['queryForm:querySpeId'] = search_id
                self.data[key]['queryForm:querySpeIdShow'] = search_id + "..."
                self.data[key]['queryForm:unitHidden2'] = city

            path = os.path.join(folder_temp, f"{city}_{search_id}.xlsx")
            # down_model.down_post(self.URL, self.data, path)
            try:
                response = down_model.down_post(self.URL, self.data, path)
                if not os.path.exists(path) or os.path.getsize(path) == 0:
                    print(f"警告: 文件可能未正确下载 - 城市: {city}, ID: {search_id}")
            except Exception as e:
                print(f"下载失败 - 城市: {city}, ID: {search_id}, 错误: {str(e)}")
# class down_performence_his():
#     # 路径：运行监控-性能查询(历史)-查询-导出
#     def __init__(self):
#         self.data = {
#             '1': config_serch_performence_his.INTO_DATA1,
#             '2': config_serch_performence_his.INTO_DATA2,
#             'FINAL': config_serch_performence_his.INTO_DATA_FINAL,
#         }


class Cascade_battery():
    def __init__(self):
        self.down_name1='电池单体电压'
        self.down_name_en1='performence_dantidianya'
        self.folder_temp1 = os.path.join(SPIDER_PATH, self.down_name_en1, 'temp')
        self.concat_name1 = os.path.join(SPIDER_PATH, self.down_name_en1)

        self.down_name2='剩余容量'
        self.down_name_en2='performence_shengyuronliang'
        self.folder_temp2 = os.path.join(SPIDER_PATH, self.down_name_en2, 'temp')
        self.concat_name2 = os.path.join(SPIDER_PATH, self.down_name_en2)

        self.down_name3='电池组总电压'
        self.down_name_en3='performence_dianchizuzongdianya'
        self.folder_temp3 = os.path.join(SPIDER_PATH, self.down_name_en3, 'temp')
        self.concat_name3 = os.path.join(SPIDER_PATH, self.down_name_en3)

        self.cities=['0099977','0099978','0099979', '0099980', '0099981', '0099982', '0099983', '0099984',
                '0099985', '0099986', '0099987', '0099988', '0099989', '0099990']

    def down1(self):
        # 路径：运行监控-性能查询-[监控点：电池单体电压]-查询-导出
        down_model.clear_folder(self.folder_temp1)
        for city in self.cities:
            try:
                # 处理每个城市前清空临时文件夹
                # 下载电流数据
                for serch_id in ['0447104001', '0447104002', '0447104003', '0447104004', '0447104005', '0447104006',
                                 '0447104007', '0447104008', '0447104009', '0447104010', '0447104011', '0447104012',
                                 '0447104013', '0447104014', '0447104015', '0447104016']:
                    down_performence().down([city], serch_id, self.folder_temp1)
                # 合并电流数据
                print(f"城市 {city} 下载完成")
            except Exception as e:
                print(f"城市 {city} 下载失败: {str(e)}")
                raise
        down_model.concat_df(self.folder_temp1, os.path.join(self.concat_name1, f"电池单体电压.xlsx"), csv=True)
        print("所有城市下载完成")

    def down2(self):
        # 路径：运行监控-性能查询-[监控点：剩余容量]-查询-导出
        down_model.clear_folder(self.folder_temp2)
        for city in self.cities:
            try:
                # 处理每个城市前清空临时文件夹

                down_performence().down([city], '0447105001', self.folder_temp2)
                # 合并电流数据
                print(f"城市 {city} 下载完成")
            except Exception as e:
                print(f"城市 {city} 下载失败: {str(e)}")
                raise
        down_model.concat_df(self.folder_temp2, os.path.join(self.concat_name2, f"剩余容量.xlsx"), csv=True)
        print("所有城市下载完成")

    def down3(self):
        # 路径：运行监控-性能查询-[监控点：电池组总电压]-查询-导出
        down_model.clear_folder(self.folder_temp3)
        for city in self.cities:
            try:
                # 处理每个城市前清空临时文件夹
                down_performence().down([city], '0447103001', self.folder_temp3)
                # 合并电流数据
                print(f"城市 {city} 下载完成")
            except Exception as e:
                print(f"城市 {city} 下载失败: {str(e)}")
                raise
        down_model.concat_df(self.folder_temp3, os.path.join(self.concat_name3, f"电池组总电压.xlsx"), csv=True)
        print("所有城市下载完成")
class performence_zhengliu():
    def __init__(self):
        # 路径：运行监控-性能查询-[监控点：信号强度]-查询-导出
        self.down_name3='整流模块电流'
        self.down_name_en3='performence_zhengliu_dianliu'
        self.folder_temp3 = os.path.join(SPIDER_PATH, self.down_name_en3, 'temp')
        self.concat_name3 = os.path.join(SPIDER_PATH, self.down_name_en3)
        self.down_name4='整流模块温度'
        self.down_name_en4='performence_zhengliu_wendu'
        self.folder_temp4 = os.path.join(SPIDER_PATH, self.down_name_en4, 'temp')
        self.concat_name4 = os.path.join(SPIDER_PATH, self.down_name_en4)
        self.cities=['0099979', '0099980']
    def down(self):
        # 处理每个城市前清空临时文件夹
        down_model.clear_folder(self.folder_temp3)
        down_model.clear_folder(self.folder_temp4)
        for city in self.cities:
            try:
                # 下载电流数据
                # for serch_id in ['0406113001', '0406113002', '0406113003', '0406113004', '0406113005', '0406113006',
                #                  '0406113007', '0406113008', '0406113009', '0406113010', '0406113011', '0406113012']:
                for serch_id in ['0406113001', '0406113002', '0406113003']:
                    down_performence().down([city], serch_id, self.folder_temp3)

                # 下载温度数据
                # for serch_id in ['0406114001', '0406114002', '0406114003', '0406114004', '0406114005', '0406114006',
                #                  '0406114007', '0406114008', '0406114009', '0406114010', '0406114011', '0406114012']:
                for serch_id in ['0406114001', '0406114002', '0406114003']:
                    down_performence().down([city], serch_id, self.folder_temp4)

                print(f"城市 {city} 下载完成")
            except Exception as e:
                print(f"城市 {city} 下载失败: {str(e)}")
                # 可以选择继续执行下一个城市
                continue

        print("所有城市下载完成")
        # 合并电流数据
        down_model.concat_df(self.folder_temp3, os.path.join(self.concat_name3, "整流模块电流.xlsx"), csv=True)
        down_model.concat_df(self.folder_temp4, os.path.join(self.concat_name4, "整流模块温度.xlsx"), csv=True)

Cascade_battery().down1()
# performence_zhengliu().down()
