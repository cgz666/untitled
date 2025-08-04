from utils.sql import sql_orm
import os
from xlsx2csv import Xlsx2csv
from sqlalchemy import inspect
import numpy as np
import pandas as pd


def insert_into_sql(df, table_name):
    with sql_orm(database='自助取数').session_scope() as temp:
        session, Base = temp

        # 1. 全局替换所有可能的空值形式为None
        df = df.replace([np.nan, 'nan', 'NaN', ''], None)
        df = df.where(pd.notnull(df), None)

        inspector = inspect(session.bind)
        if not inspector.has_table(table_name):
            raise ValueError(f"表 {table_name} 不存在于数据库中")

        print("Base.classes 中的表：", Base.classes.keys())

        if hasattr(Base.classes, table_name):
            pojo = getattr(Base.classes, table_name)
        else:
            raise ValueError(f"表 {table_name} 不存在于数据库模型中")

        # 2. 处理特定字段的转换
        if '购买发电服务' in df.columns:
            df['购买发电服务'] = df['购买发电服务'].map({'是': 1, '否': 0}).replace(np.nan, None)
        if '是否看护发电' in df.columns:
            df['是否看护发电'] = df['是否看护发电'].map({'是': 1, '否': 0}).replace(np.nan, None)
        if '自启动' in df.columns:
            df['自启动'] = df['自启动'].map({'是': 1, '否': 0}).replace(np.nan, None)

        # 3. 处理浮点型字段的空值
        for col in df.columns:
            if df[col].dtype in [float, np.float64]:
                df[col] = df[col].apply(lambda x: None if pd.isna(x) else x)

        # 4. 处理经纬度
        if '发电经度' in df.columns:
            df['发电经度'] = df['发电经度'].apply(
                lambda x: round(float(x), 6) if (x is not None and -180 <= float(x) <= 180) else None
            )
        if '发电纬度' in df.columns:
            df['发电纬度'] = df['发电纬度'].apply(
                lambda x: round(float(x), 6) if (x is not None and -90 <= float(x) <= 90) else None
            )

        # 5. 处理其他特定字段
        if '开始发电误差距离' in df.columns:
            df['开始发电误差距离'] = df['开始发电误差距离'].apply(
                lambda x: round(float(x), 2) if x is not None else None
            )

        # 6. 生成插入数据（最后一步拦截nan）
        rows = []
        for index, row in df.iterrows():
            row_dict = row.to_dict()
            # 最后防线：强制替换所有可能的nan形式
            for key, value in row_dict.items():
                # 处理数值型nan
                if isinstance(value, float) and np.isnan(value):
                    row_dict[key] = None
                # 处理字符串型nan
                elif isinstance(value, str) and value.lower() in ['nan', '']:
                    row_dict[key] = None
            temp = pojo(**row_dict)
            rows.append(temp)

        session.bulk_save_objects(rows)
        session.commit()


def xlsxtocsv(file_path):
    """将.xlsx文件转换为.csv文件"""
    file_path_csv = file_path.replace(".xlsx", ".csv")
    Xlsx2csv(file_path, outputencoding="utf-8").convert(file_path_csv)
    return file_path_csv


def process_folder(folder):
    """处理指定文件夹中的文件，将数据插入数据库"""
    df_list = []

    # 文件名到数据库表名的映射关系
    file_to_table_mapping = {
        "22年汇总.xlsx": "22发电工单",
        # 可添加其他文件映射
    }

    for file in os.listdir(folder):
        file_path = os.path.join(folder, file)
        try:
            if file.endswith('.xlsx'):
                csv_path = xlsxtocsv(file_path)
                df = pd.read_csv(csv_path, dtype=str)
                os.remove(csv_path)
            elif file.endswith('.xls'):
                df = pd.read_excel(file_path, dtype=str, engine='xlrd')
            else:
                print(f"跳过文件：{file}（不支持的文件格式）")
                continue

            # 转换发电纬度为数值类型
            if '发电纬度' in df.columns:
                df['发电纬度'] = pd.to_numeric(df['发电纬度'], errors='coerce')

            table_name = file_to_table_mapping.get(file)
            if table_name is None:
                print(f"文件 {file} 未在映射关系中找到对应的表名，跳过")
                continue

            insert_into_sql(df, table_name)
            df_list.append(df)
            print(f"文件 {file} 已成功插入到表 {table_name} 中")

        except Exception as e:
            print(f"处理文件 {file} 时出错：{e}")

    return df_list


if __name__ == "__main__":
    folder = r'C:\Users\27569\Desktop\录入数据库'  # 指定文件夹路径
    process_folder(folder)
