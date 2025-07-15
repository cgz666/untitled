import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.automap import automap_base
from contextlib import contextmanager
import pymysql


class sql_orm():
    def __init__(self, ip='10.19.6.250', port='3306', database='tower', user='root', password='123456'):
        DB_URL = {
            'url': f'mysql+pymysql://{user}:{password}@{ip}:{port}/{database}?charset=utf8',  # 数据库
            'encoding': 'utf-8',
            'pool_size': 24,
            'max_overflow': 20,
            'pool_recycle': 3600,
            'isolation_level': "READ UNCOMMITTED",
            'pool_pre_ping': True,
            'echo': False
        }
        self.engine = create_engine(**DB_URL, query_cache_size=0)
        self.Base = automap_base()
        self.Base.prepare(self.engine, reflect=True)
        self.Session = sessionmaker(bind=self.engine)

    @contextmanager
    def session_scope(self):
        session = self.Session()
        try:
            yield session, self.Base
            session.commit()
        except Exception as e:
            session.rollback()
            print('出错，此次数据回滚:' + str(e))
            raise
        finally:
            session.close()

    def get_engine(self):
        return self.engine


def insert_data_from_excel(excel_path, table_name, db_session):
    # 读取 Excel 文件
    df = pd.read_excel(excel_path)

    # 获取表对象
    table = db_session.Base.classes[table_name]

    # 将 DataFrame 转换为列表
    data = [dict(row) for index, row in df.iterrows()]

    # 批量插入数据
    with db_session.session_scope() as session:
        session.bulk_insert_mappings(table, data)
        print(f"成功插入 {len(data)} 条记录")


# 使用示例
if __name__ == "__main__":
    db_session = sql_orm()
    excel_path = 'C:/Users/27569/Desktop/23年汇总.xlsx'
    table_name = '23发电工单'

    try:
        insert_data_from_excel(excel_path, table_name, db_session)
    except Exception as e:
        print(f"插入数据时出错: {e}")