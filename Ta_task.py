from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException,NoSuchFrameException,WebDriverException
import win32com.client, win32api, win32con, win32gui
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
from datetime import datetime
import xlrd
import sqlite3


class TaError(Exception):
    pass

class Sqlite_api():
    def __init__(self,db_name,table_name):
        self.conn=sqlite3.connect(db_name)
        self.cur=self.conn.cursor()
        self.table = table_name
        self.create_table(table_name)
        self.insert()

    def create_table(self,table_name):
        sql='CREATE TABLE IF NOT EXISTS {}(id integer primary key autoincrement,'.format(table_name)
        keys = ['date_time','code','name','result','times','log','data','excel_datas_raw','datas','logs','excel_path']
        for key in keys:
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
        self.last_id = self.cur.fetchone()[0]

    def update(self,data,last_id=None):
        last_id=last_id if last_id else self.last_id
        tmp=''
        parms=[]
        for key,value in data.items():
            tmp+='{}=?,'.format(key)
            parms.append(str(value))
        parms.append(last_id)
        sql='UPDATE {} SET {} WHERE id=?'.format(self.table_name,tmp[:-1])
        self.cur.execute(sql,parms)
        self.conn.commit()

class TaTask():
    '''Ta系统自动化框架，self.driver->selenium对象，self.background_color->设置参数的背景色'''
    def __init__(self):
        self.driver = None
        self.new_code=None
        self.config=self.read_config()
        self.log_handle=open('log.txt','w',encoding='utf-8')
        self.log_all=''
        self.log_code=''
        self.db=Sqlite_api('Ta.db','job')
        self.first_id=self.db.last_id

    def read_config(self):
        f=open('config.cfg')
        t=eval(f.read())
        self.sheet_show,self.background_color,self.url,self.log_level,self.excel_path=\
            t['sheet_show'],t['background_color'],t['url'],t['log_level'],t['excel_path']
        f.close()
        return t

    def reat_fun(self,callback,callbak2,total_time=3):
        for i in range(total_time):
            callback()
            if callbak2():
                return
            time.sleep(1)


    def skip_Exception(self,callback,callbak2=None,waitException_time=5,remark=None):
        '''忽略报错函数，解决以下问题：
        因网页跳转的不确定，程序在运行过程会遇到很多不确定的报错，此时需要等待一段时间再运行代码
        Example:self.skip_Exception(lambda :ele.click()),'''
        for i in range(waitException_time*2):
            try:
                return callback()
            except Exception as e:
                if not str(e).startswith('Message: unknown error: Element <button type="button" id="addButton"'):
                    print('\033[1;31m{}\n{}\n{}\033[0m'.format(e,callback,remark))
                t=e
            time.sleep(0.5)
        raise TaError('{}|{}'.format(t,remark))

    def super_click(self,ele,mode=1,waittime=0):
        '''点击元素的函数'''
        if mode==0:            #js
            self.driver.execute_script('arguments[0].click()', ele)
        elif mode==1:         #selenium API
            ele.click()
        elif mode==2:         #    ActionChains  simulate
            ActionChains(self.driver).click(ele).perform()
        elif mode==3:         #    dm simulate
            pass
        time.sleep(waittime)

    def super_find_eles(self,value,by=By.CSS_SELECTOR,find_ele_time=5,ele_parent=None,frames=None,remark=None,return_all=False,waittime=None):
        '''定位元素函数，默认参数：by=By.CSS_SELECTOR->查找方式，默认用CSS,find_ele_time=5->查找时间，默认5秒,ele_parent=None->父元素
        frames=None->iframe,remark=None->描述,return_all=False->是否返回所有，默认只会返回一个元素,
        waittime=None->找到元素之后等待时间,log=None->写入log的内容'''
        NoSuchFrame = False
        for i in range(1,int(find_ele_time*10)):                       #每0.1秒找一次
            if frames:
                self.driver.switch_to.default_content()
                for frame in convert_to_list(frames):             #frames such as ['iframe1','iframe2',-1]
                    try:
                        if isinstance(frame,str):
                            self.driver.switch_to.frame(frame)
                        elif frame==-1:                           #找出所有'iframe'元素，取最后一个
                            self.driver.switch_to.frame(self.super_find_eles('iframe',return_all=True)[-1])
                    except NoSuchFrameException:
                        # print('{} NoSuchFrameException 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),frame,i))
                        NoSuchFrame=True
                        break
            if NoSuchFrame:
                time.sleep(0.1)
                NoSuchFrame=False
                continue
            eles = ele_parent.find_elements(by, value) if ele_parent else self.driver.find_elements(by, value)
            if eles:                     #找到元素，返回[ele1,ele2,...]
                if waittime:
                    time.sleep(waittime)
                # print('{} find the element{} 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                #                                         's' if return_all else '',remark if remark else value,i))
                # if log:
                #     self.log_write(log)
                return eles if return_all else eles[0]
            else:                       #找不到元素，返回[]
                pass
                # print('{} NoSuchElementException 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                #                                                            remark if remark else value,i))
            time.sleep(0.1)
        else:                         #在find_ele_time时间内找不到元素
            if value not in ['label']:
                print('\033[1;31mcan not find element!{}\033[0m'.format(remark if remark else value))       #红色字体显示
            # if log:
            #     self.log_write(log,fail=True)
            return

    def compare_values(self,data_excel,frames=None,length=None,data_mode='form',selcet_data_by_value=None,filter=[]):
        '''return such as{'募集截止日期': ('2013-03-14', '2016-02-02')}'''
        if data_mode=='form':
            data_sys = self.get_data_sys_form(frames=frames, length=length)
        else:
            data_sys,tds=self.get_data_sys_table(selcet_data_by_value,frames=frames)
        result_compare = {}
        # print('='*100)
        # print(data_excel)
        # print(data_sys)
        # print('='*100)
        if not data_sys:
            return result_compare,tds,data_sys
        for key in data_excel:
            if key not in data_sys :
                continue
            if not self.compaer_value(data_excel[key],data_sys[key]):
                result_compare.update({key:(data_excel[key],data_sys[key])})
        if not self.otc:
            filter.append('基金代码')
        result_compare={key:value for key,value in result_compare.items() if key not in filter}
        if data_mode=='form':
            return result_compare
        else:
            return result_compare,tds,data_sys

    def compaer_value(self,value_excel,value_sys):
        '''compaer_excel_sys_data函数的子函数,value_excel->参数表数据,value_sys->ta系统数据'''
        try:       #尝试相减比较
            if value_excel != '':  # 取:左边的数据
                value_excel = value_excel.replace('：', ':').split(':')[0]
            return True if float(value_excel) - float(value_sys.replace(',', '')) == 0 else False
        except:      #如报错比较字符串
            return True if value_excel == value_sys else False

    def get_data_sys_form(self,frames=None,length=None):
        '''取系统参数值,return such as {'基金名称':'广发理财-招商-2'....} '''
        for i in range(5):      #每秒取一次
            try:
                dts = self.super_find_eles('dt', frames=frames,return_all=True)
                dds = self.super_find_eles('dd', return_all=True)
                if length and len(dds)<length:             #如果字段数量不等于指定数量length，则time.sleep(1)，避免网络延迟原因
                    print(len(dds))
                    time.sleep(1)
                    continue
                return dict(zip(self.get_dts_innerTexts(dts),self.get_dds_values(dds)))
            except:
                time.sleep(1)
        raise SystemError('too lag!')

    def get_data_sys_table(self,selcet_data_by_value,frames=None):
        table = self.super_find_eles('table.datagrid-htable',frames=frames,return_all=True)[-1]
        tds = self.super_find_eles('td', ele_parent=table, return_all=True)
        header = [self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0', '').strip() for x in tds[2:]]
        table_trs = self.super_find_eles('table.datagrid-btable', return_all=True)[-1]
        # print(header)
        trs = self.super_find_eles('tr', ele_parent=table_trs, return_all=True)[::-1]
        tdss = [self.super_find_eles('td', ele_parent=tr, return_all=True) for tr in trs]
        tmp = [x for x in tdss if x[2].text == selcet_data_by_value] if selcet_data_by_value else tdss
        tds = tmp[0] if tmp else None
        return  dict(zip(header, [self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0', '').strip()
                         for x in tds[2:]])) if tds else None,tds


    def get_eles_to_set(self,keys=None,frames=None,value='dd'):
        '''取要设置参数的eles,keys->参数表的字段名称，frames->iframe,return such as{'基金名称':ele1....}'''
        def f():
            eles= self.super_find_eles(value, return_all=True, frames=frames)
            return {key: ele for key, ele in zip(self.get_dts_innerTexts(), eles)
                    if key in keys} if keys else dict(zip(self.get_dts_innerTexts(), eles))
        return self.skip_Exception(f)                #用skip_Exception运行，避免报错

    def get_dts_innerTexts(self,dts=None):
        '''取要参数字段名称，dts->字段名称元素,return such as['基金名称','基金代码'...]'''
        if dts is None:
            dts=self.super_find_eles('dt',return_all=True)
        return [self.driver.execute_script('return arguments[0].innerText', x).replace('*','').replace(':','') for x in dts]   #用js，提高效率

    def get_dds_values(self,dds):
        '''取参数值，dds->参数值元素,return such as['广发理财-招商-2','870022'...]'''
        return [self.driver.execute_script('return arguments[0].querySelector("select,input").value', x) for x in dds]      #用js，提高效率

    def set_value(self,ele,value,key=None,sel=None):
        '''设置一个参数函数,ele->要设置参数的元素,value->参数值,key=None->字段名称,sel->span标签的selector'''
        control=ele if ele.tag_name in ["select", "input"] else self.skip_Exception(
            lambda :self.super_find_eles('select,input',ele_parent=ele))         #定位到select或者input标签
        if control.tag_name == 'input':     #文本框
            value_cur=self.driver.execute_script('return arguments[0].value',control)
            if value_cur == value:
                return
            self.driver.execute_script('arguments[0].style.backgroundColor=arguments[1]',control,self.background_color)     #改变背景色
            self.driver.execute_script('arguments[0].value=arguments[1]', control, value)
            ActionChains(self.driver).send_keys_to_element(control,Keys.SPACE+Keys.BACKSPACE).perform()
            # if mode==0:
            #     if value_cur!='':
            #         control.clear()
            #     control.send_keys(value)
            # if mode==1:
            #     self.driver.execute_script('arguments[0].value=arguments[1]', control, value)
                # control.clear() or control.send_keys(value) if key in self.onchange_key else self.driver.execute_script(
                # 'arguments[0].value=arguments[1]', control, value)          #如果key in self.onchange_key，用send_keys，否则用js设置
            if value==self.skip_Exception(lambda :self.driver.execute_script('return arguments[0].value', control)):      #设置后再比对一次
                return
            else:
                msgbox('set value wrong!   {}'.format(value))
                return
        elif control.tag_name == 'select':                #下拉框
            parent = self.driver.execute_script('return arguments[0].parentNode', control)     #找到dd元素
            if control.get_attribute('multiple')=='true':    #多选
                options = self.super_find_eles('option', ele_parent=control,return_all=True)
                values_options = [self.driver.execute_script('return arguments[0].value', x) for x in options]
                # print(values_options)
                value_to_set=[x.replace('：', ':').split(':')[0] for x in value.split(',')]     #取冒号左边，如
                inp=self.super_find_eles('input',ele_parent=parent)              #找到input元素，可以发送键盘消息
                self.driver.execute_script('arguments[0].focus()', inp)                        #焦点定位到input
                time.sleep(0.2)
                for x in value_to_set:
                    if x not in values_options:
                        raise TaError('multiple invalid data!')
                    ActionChains(self.driver).send_keys(x + Keys.ENTER).perform()
                    time.sleep(0.05)
                return
            else:   # 单选
                value_to_set = value.replace('：', ':').split(':')[0]
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:block")   #将元素设置为可见
                Select(control).select_by_value(value_to_set)                 #选中值
                sel_span=('{}+div>a>span').format(sel) if sel else ('div>a>span')
                span=parent.find_element_by_css_selector(sel_span)
                self.driver.execute_script('arguments[0].innerText=arguments[1]', span, value)        #将span显示为value
                self.driver.execute_script('arguments[0].style.backgroundColor=arguments[1]', span,self.background_color)   #改背景颜色
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:none")
                return

    def set_values(self,eles,datas,remark=None):
        '''设置多个参数函数,ele->要设置参数的元素,datas->参数值,'''
        t=time.time()
        for key,value in datas.items():
            # print(key,value)
            if key not in eles or value is None:
                continue
            self.skip_Exception(lambda :self.set_value(eles[key],value,key),waitException_time=2,remark=key)
            # if remark=='22':
            #     input()
        if remark:
            print('\033[1;33m{} {}\033[0m'.format(remark,round(time.time()-t,2)))
        self.check_invalid_data()

    def check_invalid_data(self,frames_label=None):
        # print(1)
        ele_label = self.super_find_eles('label', frames=frames_label, find_ele_time=0.5)
        if ele_label and ele_label.text!='':
        # if ele_label:
            print('\033[1;31m系统报错：{}\033[0m'.format(ele_label.text))  # 红色
            self.driver.execute_script("arguments[0].scrollIntoView()", ele_label)
            # input()
            raise TaError('invalid data|{}'.format(ele_label.text))
        return

    def log_write(self,text,newline=True,date_time=True,level=0):
        log='{} {} {}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S") if date_time else '',text,'\n' if newline else '')
        self.log_all+=log
        self.log_code+=log
        self.db.update({'log':self.log_code})
        if level<=self.log_level:
            self.log_handle.write(log)

def msgbox(text,method=0,btmod=0):
    '''弹窗函数'''
    if method==0:
        return win32api.MessageBox(0,str(text),"提示",btmod+win32con.MB_ICONINFORMATION+win32con.MB_SETFOREGROUND)    # 提示
    elif method==1:
        return win32api.MessageBox(0, str(text), "警告!", btmod+win32con.MB_ICONEXCLAMATION+win32con.MB_SETFOREGROUND)  #警告
    elif method == 2:
        return win32api.MessageBox(0, str(text), "严重错误!!", btmod+win32con.MB_ICONSTOP+win32con.MB_SETFOREGROUND)    # 严重错误

def convert_to_list(data):
    '''转换成list，如：'test'->['list'],1->[1] '''
    ret=[]
    return data if isinstance(data,list) else ret.append(data) or ret




