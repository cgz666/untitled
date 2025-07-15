import os
IP_SERVICE='10.19.6.250'
BROWSER_DOWN='C:/Users/Administrator/Downloads/'
INDEX = os.path.dirname(os.path.abspath(__file__))+'/'
SPIDER_PATH=f'{INDEX}websource/spider_download/'
TEMP_PATH_ONE_MONTH=f'{INDEX}websource/temp_folder_one_month/'
TEMP_PATH_ONE_DAY=f'{INDEX}websource/temp_folder_one_day/'
LOG_PATH=f'{INDEX}log/'

FILES = {
    'quxin_order_youji': {'path': r'F:\newtower3.8\project\quxin_script\工单及油机匹配率统计\output\工单及油机匹配率统计-结果.xlsx'},
    'Interface_order': {'path': r'F:\untitled\four_a_script\Interface_result\output\结果.xls'}
}