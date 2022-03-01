# -*- coding:utf-8 -*-
# @Author : Witty
import faker
import pyodbc
import pandas as pd


class Pysql:
    """初始化数据库"""

    def __init__(self, driver, server, uid, pwd, database):
        self.driver = driver
        self.server = server
        self.uid = uid
        self.pwd = pwd
        self.database = database

    """连接数据库"""

    def get_connect(self):
        self.conn = pyodbc.connect(
            f"DRIVER={self.driver};server={self.server};uid={self.uid};pwd={self.pwd};database={self.database}")
        self.cursor = self.conn.cursor()
        if self.conn:
            print("数据库连接成功")
        else:
            print("请检查")

    """查询"""

    def select_sql(self, sql):
        self.cursor.execute(sql)
        db = self.cursor.fetchall()
        print(db)

        """生成模拟数据"""

    def mock(self, count):
        fake = faker.Faker("zh_CN")
        data_total = []
        self.data_total = [
            [
                fake.date(),
                fake.random_int(67000, 67999),  # 用户编号
                fake.name(),  # 名字
                fake.phone_number(),  # 手机号
                fake.random_int(671, 671)
            ]  # 店铺编号
            for x in range(count)
        ]
        # print(self.data_total)

    def read_excel(self):
        df = pd.read_excel("导购数据.xlsx", sheet_name=0)
        self.data = df.iloc[0:, [0, 1, 2, 3, 4]]
        self.data.clumns = ['d_id', 'c_yuangbh', 'c_yuangmc', 'c_soujhm', 'shopid']
        print(self.data)

    def insert_sql(self):
        for val in range(len(self.data)):
            d_id = str(self.data['d_id'].iloc[val])
            c_yuangbh = str(self.data['c_yuangbh'].iloc[val])
            c_yuangmc = str(self.data['c_yuangmc'].iloc[val])
            c_soujhm = str(self.data['c_soujhm'].iloc[val])
            shopid = str(self.data['shopid'].iloc[val])

            self.cursor.execute(f"INSERT INTO r_yuang ([d_id],[c_yuangbh],[c_yuangmc],[c_soujhm],[shopid]) VALUES(" \
                                f"{d_id}," \
                                f"{c_yuangbh}," \
                                f"'{c_yuangmc}'," \
                                f"{c_soujhm}," \
                                f"{shopid})")
            self.cursor.commit()  # 不能忽略，否则不会提交到数据库

    def deal_excel(self):
        """创建数据写入excel"""
        df = pd.DataFrame(self.data_total, columns=['d_id', 'c_yuangbh', 'c_yuangmc', 'c_soujhm', 'shopid'])
        df.to_excel("导购数据.xlsx", index=False)
        print("写入成功")


def main():
    getsql = Pysql('SQL Server',
                   '服务器地址',
                   '用户名',
                   '密码',
                   '数据库名称')
    getsql.get_connect()  # 连接数据库
    getsql.mock(5)  # 生成多少条数据
    getsql.deal_excel()  # 导出数据
    getsql.read_excel()
    getsql.insert_sql()  # 插入数据库
    getsql.cursor.close()  # 关闭游标
    getsql.conn.close()  # 关闭数据库连接


if __name__ == "__main__":
    main()
