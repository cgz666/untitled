import os
import pandas as pd

# 文件夹路径
folder_path = r'F:\untitled\experimental_data\xls'
output_folder = r'F:\untitled\experimental_data\output'  # 输出文件夹路径

# 确保输出文件夹存在
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 获取文件夹内所有PDF文件的名称
file_names = [file for file in os.listdir(folder_path) if file.endswith('.pdf')]

# 存储提取出的编号
ids = []

# 遍历文件名列表，提取前18个字符作为编号
for file_name in file_names:
    # 提取前18个字符
    extracted_id = file_name[:18]
    # 检查提取的编号是否包含至少两个短横线
    if extracted_id.count('-') >= 2:
        ids.append(extracted_id)

# 创建一个DataFrame
df = pd.DataFrame(ids, columns=['ID'])

# 将DataFrame保存到CSV文件
output_file_path = os.path.join(output_folder, 'extracted_ids.csv')
df.to_csv(output_file_path, index=False)

print(f"提取的编号已保存到 {output_file_path}")