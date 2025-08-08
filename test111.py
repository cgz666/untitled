import pandas as pd

# 定义文件路径
input_file = r'C:\Users\27569\Desktop\1.xlsx'  # 替换为你的文件名
output_file = r'C:\Users\27569\Desktop\24年7-25年7停电数量（联通）.xlsx'  # 输出文件名

# 定义需要处理的工作表名称
sheets = ['24年7-12月', '25年1-7月']

# 定义需要转换格式的列
columns_to_convert = [
    '告警发生时间', '告警清除时间', '告警入库时间',
    '告警开始时间（FSU）', '告警结束时间（FSU）'
]

# 读取 Excel 文件
with pd.ExcelFile(input_file) as xls:
    # 创建一个字典来存储处理后的数据
    processed_sheets = {}

    # 遍历每个工作表
    for sheet in sheets:
        # 读取当前工作表
        df = pd.read_excel(xls, sheet_name=sheet)

        # 遍历需要转换的列
        for col in columns_to_convert:
            if col in df.columns:
                # 转换日期格式
                df[col] = pd.to_datetime(df[col], format='%d/%m/%Y %H:%M:%S').dt.strftime('%Y/%m/%d %H:%M:%S')

        # 将处理后的数据存储到字典中
        processed_sheets[sheet] = df

# 将处理后的数据写入新的 Excel 文件
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, df in processed_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"处理完成，结果已保存到 {output_file}")