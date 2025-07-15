import os

# 假设你的文件位于当前目录下的'files'文件夹中
folder_path = r'F:\untitled\experimental_data\xls'

# 获取文件夹内所有文件的名称
file_names = os.listdir(folder_path)

# 存储提取出的编号
ids = []

# 遍历文件名列表，提取编号
for file_name in file_names:
    if file_name.endswith('.pdf'):  # 确保只处理PDF文件
        # 提取编号，假设编号格式为HPJC-240320-898704，位于文件名开头
        parts = file_name.split('-')
        if len(parts) >= 3:
            id = parts[0] + '-' + parts[1] + '-' + parts[2]
            ids.append(id)

# 打印提取出的编号
for id in ids:
    print(id)