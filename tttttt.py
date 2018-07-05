import sqlite3
import datetime
import os
import time
class Sqlite_api():
    def __init__(self,db_name,table_name):
        self.conn=sqlite3.connect(db_name)
        self.cur=self.conn.cursor()
        self.table = table_name
        self.create_table(table_name)

    def create_table(self,table_name):
        sql='CREATE TABLE IF NOT EXISTS {}({}'.format(table_name,'id integer primary key autoincrement,' if id else '')
        keys = {'aa': 'text', 'bb': 'int', 'cc': 'hhgg'}
        # t=list(kwargs.values())[0] if (len(kwargs)==1 and isinstance(list(kwargs.values())[0],dict)) else kwargs
        for key, value in keys.items():
            # sql += '{} {},'.format(key, 'integer' if value in ['integer','int'] else 'text')
            sql += '{} text,'.format(key)
        sql=sql[:-1]+')'
        print(sql)
        self.table_name=table_name
        self.cur.execute(sql)

    def insert(self):
        sql='insert into {}(id) values(null)'.format(self.table_name)
        self.cur.execute(sql)
        self.conn.commit()
        self.cur.execute("select * from {} order by id desc LIMIT 1".format(self.table_name))
        res = self.cur.fetchone()
        print(res)
        return res[0]

    def update(self,data,last_id=None):
        if not last_id:
            self.cur.execute("select * from {} order by id desc LIMIT 1".format(self.table_name))
            last_id=self.cur.fetchone()[0]
        print(last_id)
        tmp=''
        parms=[]
        for key,value in data.items():
            tmp+='{}=?,'.format(key)
            parms.append(value)
        parms.append(last_id)
        print(tmp[:-1])
        print(parms)
        sql='UPDATE {} SET {} WHERE id=?'.format(self.table_name,tmp[:-1])
        print(sql)
        self.cur.execute(sql,parms)
        self.conn.commit()
        self.cur.execute("select * from {}".format(self.table_name))
        res = self.cur.fetchall()
        print(res)


import msvcrt


def pwd_input():
    chars = []
    while True:
        try:
            newChar = msvcrt.getch().decode(encoding="utf-8")
        except:
            return input("你很可能不是在cmd命令行下运行，密码输入将不能隐藏:")
        if newChar in '\r\n':  # 如果是换行，则输入结束
            break
        elif newChar == '\b':  # 如果是退格，则删除密码末尾一位并且删除一个星号
            if chars:
                del chars[-1]
                msvcrt.putch('\b'.encode(encoding='utf-8'))  # 光标回退一格
                msvcrt.putch(' '.encode(encoding='utf-8'))  # 输出一个空格覆盖原来的星号
                msvcrt.putch('\b'.encode(encoding='utf-8'))  # 光标回退一格准备接受新的输入
        else:
            chars.append(newChar)
            msvcrt.putch('*'.encode(encoding='utf-8'))  # 显示为星号
    return (''.join(chars))

a=r"2016/1/12"
b=time.strptime(a, "%Y/%m/%d")
print(b)
# print("Please input your password:")
# pwd = pwd_input()
# print("\nyour password is:{0}".format(pwd))
# # sys.exit()
exit()




aa=Sqlite_api('test.db','table5')
# last_id=aa.insert()
# print(last_id)
aa.update({'aa':datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'bb':1,'cc':'务管理总部'})
