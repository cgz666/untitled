# database.py

import mysql.connector
from mysql.connector import Error

class Database:
    def __init__(self, host, user, password, database):
        self.host = '10.19.6.250'
        self.user = 'root'
        self.password = '123456'
        self.database = 'test'
        self.connection = None

    def connect(self):
        """建立数据库连接"""
        try:
            self.connection = mysql.connector.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                database=self.database
            )
            if self.connection.is_connected():
                print("数据库连接成功")
        except Error as e:
            print(f"数据库连接失败: {e}")

    def disconnect(self):
        """关闭数据库连接"""
        if self.connection.is_connected():
            self.connection.close()
            print("数据库连接已关闭")

    def insert_data(self, table, columns, values):
        """插入数据到指定表"""
        try:
            if not self.connection.is_connected():
                self.connect()
            cursor = self.connection.cursor()
            placeholders = ', '.join(['%s'] * len(values))
            query = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
            cursor.execute(query, values)
            self.connection.commit()
            print("数据插入成功")
        except Error as e:
            print(f"数据插入失败: {e}")
        finally:
            cursor.close()

# 示例用法
if __name__ == "__main__":
    db = Database("localhost", "root", "password", "testdb")
    db.connect()
    db.insert_data("users", "name, age", ("John Doe", 30))
    db.disconnect()