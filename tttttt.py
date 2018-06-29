import sqlite3
import datetime

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


a=''

print(a is None)
exit()




aa=Sqlite_api('test.db','table5')
# last_id=aa.insert()
# print(last_id)
aa.update({'aa':datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),'bb':1,'cc':'务管理总部'})
