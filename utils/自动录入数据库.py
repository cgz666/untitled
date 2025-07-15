from utils.sql import sql_orm
import os
from xlsx2csv import Xlsx2csv
import pandas as pd
from sqlalchemy import inspect


def insert_into_sql(df, table_name):
    with sql_orm(database='test').session_scope() as temp:
        session, Base = temp
        df = df.where(pd.notnull(df), None)
        inspector = inspect(session.bind)
        if not inspector.has_table(table_name):
            raise ValueError(f"表 {table_name} 不存在于数据库中")

        print("Base.classes 中的表：", Base.classes.keys())  # 调试信息

        if hasattr(Base.classes, table_name):
            pojo = getattr(Base.classes, table_name)
        else:
            raise ValueError(f"表 {table_name} 不存在于数据库模型中")

        rows = []
        for index, row in df.iterrows():
            temp = pojo(**row.to_dict())
            rows.append(temp)
        session.bulk_save_objects(rows)
        session.commit()


def xlsxtocsv(file_path):
    """
    将.xlsx文件转换为.csv文件
    :param file_path: .xlsx文件路径
    :return: 转换后的.csv文件路径
    """
    file_path_csv = file_path.replace(".xlsx", ".csv")
    Xlsx2csv(file_path, outputencoding="utf-8").convert(file_path_csv)
    return file_path_csv


def process_folder(folder):
    """
    处理指定文件夹中的文件，将数据插入数据库
    :param folder: 文件夹路径
    """
    df_list = []  # 用于存储所有处理的DataFrame

    # 定义文件名到数据库表名的映射关系
    file_to_table_mapping = {
        "output.xlsx": "output",
        # 确保这里没有拼写错误
    }

    for file in os.listdir(folder):
        file_path = os.path.join(folder, file)
        if file.endswith('.xlsx'):
            # 如果是.xlsx文件，先转换为.csv
            csv_path = xlsxtocsv(file_path)
            df = pd.read_csv(csv_path, dtype=str)
            os.remove(csv_path)  # 删除临时生成的.csv文件
        elif file.endswith('.xls'):
            # 如果是.xls文件，直接读取
            df = pd.read_excel(file_path, dtype=str, engine='xlrd')
        else:
            print(f"跳过文件：{file}（不支持的文件格式）")
            continue

        # 通过文件名获取对应的数据库表名
        table_name = file_to_table_mapping.get(file)
        if table_name is None:
            print(f"文件 {file} 未在映射关系中找到对应的表名，跳过")
            continue

        try:
            insert_into_sql(df, table_name)
            df_list.append(df)
            print(f"文件 {file} 已成功插入到表 {table_name} 中")
        except Exception as e:
            print(f"处理文件 {file} 时出错：{e}")

    return df_list


if __name__ == "__main__":
    folder = r'C:\Users\27569\Desktop\录入数据库'  # 指定文件夹路径
    process_folder(folder)