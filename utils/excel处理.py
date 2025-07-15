import os
import pandas as pd
from datetime import datetime

# 定义文件路径
FOLDER_PATH = r"C:\Users\27569\Desktop\录入数据库"
INPUT_FILE = os.path.join(FOLDER_PATH, 'input.xlsx')  # 输入文件路径
OUTPUT_FILE = os.path.join(FOLDER_PATH, 'output.xlsx')  # 输出文件路径
LOG_FILE = os.path.join(FOLDER_PATH, 'deleted_records.csv')  # 删除记录日志文件

def convert_to_float(value):
    """将值转换为浮点数，空值或无效值返回0"""
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0

def remove_duplicates(df, key_column, date_column):
    """
    删除重复项，只保留指定日期最新的记录
    若有最新的时间重复的也保留
    """
    if key_column not in df.columns or date_column not in df.columns:
        print(f"未找到'{key_column}'或'{date_column}'列，无法删除重复项")
        return df

    # 临时将日期转换为datetime进行比较
    df[date_column + '_temp'] = pd.to_datetime(df[date_column], errors='coerce')

    # 按关键列分组，找到每个关键列的最大日期
    max_dates = df.groupby(key_column)[date_column + '_temp'].transform('max')

    # 保留日期等于最大日期的记录
    df_filtered = df[df[date_column + '_temp'] == max_dates].copy()

    # 记录删除的重复项
    df_deleted = df[df[date_column + '_temp'] < max_dates].copy()

    # 保存删除的记录到日志文件
    if not df_deleted.empty:
        print(f"正在保存{len(df_deleted)}条删除记录到: {LOG_FILE}")
        df_deleted.to_csv(LOG_FILE, index=False)

    # 删除临时列，不保留在最终结果中
    df_filtered = df_filtered.drop(columns=[date_column + '_temp'])

    return df_filtered

def process_files():
    try:
        if not os.path.exists(INPUT_FILE):
            print(f"输入文件 {INPUT_FILE} 不存在")
            return

        # 以字符串形式读取Excel文件的所有数据
        xls = pd.ExcelFile(INPUT_FILE)

        # 检查是否包含 '巡检数据' 和 '输出表格情况' 两个工作表
        if '巡检数据' not in xls.sheet_names or '输出表格情况' not in xls.sheet_names:
            print("工作簿中必须包含 '巡检数据' 和 '输出表格情况' 两个工作表")
            return

        # 以字符串形式读取两个工作表
        df_inspect = xls.parse('巡检数据', dtype=str)
        df_output = xls.parse('输出表格情况', dtype=str)

        # 删除 '巡检数据' 表中的重复项
        print("正在处理 '巡检数据' 表中的重复项...")
        df_inspect = remove_duplicates(df_inspect, '站址编码', '巡检日期')

        # 删除 '输出表格情况' 表中的重复项
        print("正在处理 '输出表格情况' 表中的重复项...")
        df_output = remove_duplicates(df_output, '站址名称', '巡检日期')

        # 保存到输出文件
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            df_inspect.to_excel(writer, sheet_name='巡检数据', index=False)
            df_output.to_excel(writer, sheet_name='输出表格情况', index=False)

        print(f"数据处理完成，结果已保存到: {OUTPUT_FILE}")
        if os.path.exists(LOG_FILE):
            print(f"删除记录已保存到: {LOG_FILE}")

    except Exception as e:
        print(f"处理过程中出错: {str(e)}")
        raise

if __name__ == "__main__":
    process_files()