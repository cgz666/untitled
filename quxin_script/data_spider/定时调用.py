import sys
sys.path.append(r'F:\untitled')
sys.path.append(r'F:\untitled\venv\Lib\site-packages')
import schedule
import time
import os
from quxin_script.data_spider.gongdan_youji_pipei.工单及油机匹配率脚本 import main as gongdan_main
from quxin_script.data_spider.yidong_tuifu_static.移动无运营商推送退服类告警统计脚本 import main as yidong_main

import logging
# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 获取当前程序的目录
current_script_dir = os.path.dirname(os.path.abspath(__file__))
logging.info(f"当前程序目录: {current_script_dir}")

# 定义运行脚本的函数
def run_scripts():
    try:
        logging.info("开始运行工单及油机匹配率脚本")
        gongdan_main()
        logging.info("工单及油机匹配率脚本运行完成")
    except Exception as e:
        logging.error(f"工单及油机匹配率脚本运行出错: {e}")

    try:
        logging.info("开始运行移动无运营商推送退服类告警统计脚本")
        yidong_main()
        logging.info("移动无运营商推送退服类告警统计脚本运行完成")
    except Exception as e:
        logging.error(f"移动无运营商推送退服类告警统计脚本运行出错: {e}")

    logging.info("所有脚本运行完成")

# 定义设置定时任务的函数

def schedule_function():
    # 设置每个工作日上午9点运行
    schedule.every().monday.at("09:00").do(run_scripts)
    schedule.every().tuesday.at("09:00").do(run_scripts)
    schedule.every().wednesday.at("09:00").do(run_scripts)
    schedule.every().thursday.at("09:00").do(run_scripts)
    schedule.every().friday.at("09:00").do(run_scripts)
    logging.info("定时任务已设置，将在每个工作日上午9点运行")

if __name__ == '__main__':
    # 设置定时任务
    schedule_function()
    logging.info("开始运行定时任务")

    try:
        # 保持主程序运行
        while True:
            schedule.run_pending()
            time.sleep(1)
    except (KeyboardInterrupt, SystemExit):
        logging.info("定时任务已停止")