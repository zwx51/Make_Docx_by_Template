# encoding = utf8
import pymssql
import configparser
import os


class PymysqlUtil():
    #初始化方法
    #def __init__(self, host, user, passwd, dbName):
    def __init__(self):
        #self.host = 'localhost'
        #self.user = 'sa'
        #self.passwd = '666666'
        #self.dbName = 'test'
        root_dir = os.path.abspath('.')
        cf = configparser.ConfigParser()
        cf.read(root_dir + "/config/db.ini")
        host = cf.get("Database", "host")
        user = cf.get("Database", "user")
        passwd = cf.get("Database", "passwd")
        dbname = cf.get("Database", "dbName")
        self.host = host
        self.user = user
        self.passwd = passwd
        self.dbName = dbname
    #链接数据库
    def getCon(self):
        try:
            self.db = pymssql.connect(
                self.host, self.user, self.passwd, self.dbName
            )
            self.cursor = self.db.cursor()
        except Exception as e:
            print(e)
    #关闭链接
    def close(self):
        self.cursor.close()
        self.db.close()
    #查询单行记录
    def get_one(self,sql):
        res = None
        try:
            self.getCon()
            self.cursor.execute(sql)
            res = self.cursor.fetchone()
        except:
            print("查询失败!")
        return res
    #查询列表数据
    def get_all(self,sql):
        res = None
        try:
            self.getCon()
            self.cursor.execute(sql)
            res = self.cursor.fetchall()
            self.close()
        except:
            print("查询失败！")
        return res
    #插入数据
    def insert(self,sql):
        count = 0
        try:
            self.getCon()
            count = self.cursor.execute(sql)
            self.db.commit()
            self.close()
        except:
            print("操作失败！")
            self.db.rollback()
        return count
    #修改数据
    def edit(self,sql):
        return self.__insert(sql)
    #删除数据
    def delete(self,sql):
        return self.__insert(sql)
    #更新数据
    def update(self,sql):
        return self.__insert(sql)


if __name__ == '__main__':
    app = PymysqlUtil()