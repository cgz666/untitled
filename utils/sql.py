from sqlalchemy import create_engine,text
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.automap import automap_base
from contextlib import contextmanager
import pandas as pd
class sql_orm():
    def __init__(self, ip='10.19.6.250', port='3306', database='自助取数', user='root', password='123456'):
        DB_URL = {
            'url': f'mysql+pymysql://{user}:{password}@{ip}:{port}/{database}?charset=utf8',  # 数据库

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

    def excute_sql(self,sql_str,return_df=False):
        with self.session_scope() as (sql, Base):
            res=sql.execute(text(sql_str))
            if return_df:
                colnames = [column[0] for column in res.cursor.description]
                rows = res.fetchall()
                data_list = [dict(zip(colnames, row)) for row in rows]
                df = pd.DataFrame(data_list)
                return df