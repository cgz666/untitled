# main.py

from database import Database


def main():
    # 创建数据库对象
    db = Database("localhost", "root", "password", "testdb")

    # 连接数据库
    db.connect()

    # 插入数据
    db.insert_data("users", "name, age", ("Jackey", 24))

    # 断开数据库连接
    db.disconnect()


if __name__ == "__main__":
    main()