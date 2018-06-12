from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException,NoSuchFrameException,WebDriverException
import win32com.client, win32api, win32con, win32gui
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
from datetime import datetime

LazyExcel = win32com.client.Dispatch('Lazy.LxjExcel')             #处理excel的函数

class TaTask():
    '''Ta系统自动化框架，self.driver->selenium对象，self.background_color->设置参数的背景色'''
    def __init__(self):
        self.driver = None
        self.background_color='navajowhite'                #['cornsilk','red','silver','none','navajowhite' ]
        self.log=''

    def __del__(self):
        print('exit')

    def skip_Exception(self,callback,callbak2=None,waitException_time=5,remark=None):
        '''忽略报错函数，解决以下问题：
        因网页跳转的不确定，程序在运行过程会遇到很多不确定的报错，此时需要等待一段时间再运行代码
        Example:self.skip_Exception(lambda :ele.click()),'''
        for i in range(waitException_time*2):
            try:
                if (callbak2 and callbak2()) or (callbak2 is None):
                    return callback()
            except Exception as e:
                print('\033[1;31m{}\n{}\n{} \033[0m'.format(e,callback,remark))
            time.sleep(0.5)
        raise

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

    def super_find_eles(self,value,by=By.CSS_SELECTOR,find_ele_time=5,ele_parent=None,frames=None,remark=None,return_all=False,waittime=None,log=None):
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
                if log:
                    self.log_write(log)
                return eles if return_all else eles[0]
            else:                       #找不到元素，返回[]
                pass
                # print('{} NoSuchElementException 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                #                                                            remark if remark else value,i))
            time.sleep(0.1)
        else:                         #在find_ele_time时间内找不到元素
            if value not in ['label']:
                print('\033[1;31mcan not find element!{}\033[0m'.format(remark if remark else value))       #红色字体显示
            if log:
                self.log_write(log,fail=True)
            return

    def compare_values(self,data_excel,frames=None,length=None):
        '''比对'''
        data_sys =self.get_data_sys(frames=frames,length=length)
        result_compare={}
        for key in data_excel:
            if key not in data_sys:
                continue
            if self.compaer_value(data_excel[key],data_sys.get(key)):
                pass
                # print(key,data_excel[key],data_sys.get(key),'correctly!')
            else:
                result_compare.update({key:(data_excel[key],data_sys[key])})
        return result_compare      #such as {'基金名称': ('和聚(玉融)量化空盈9号私募基金', '广发理财-招商-2')}

    def compaer_value(self,value_excel,value_sys):
        '''compaer_excel_sys_data函数的子函数,value_excel->参数表数据,value_sys->ta系统数据'''
        if value_excel != '':                 #取:左边的数据
            value_excel = value_excel.replace('：', ':').split(':')[0]
        try:       #尝试相减比较
            return True if float(value_excel) - float(value_sys.replace(',', '')) == 0 else False
        except:      #如报错比较字符串
            return True if value_excel == value_sys else False

    def get_data_sys(self,frames=None,log=None,length=None):
        '''取系统参数值,return such as {'基金名称':'广发理财-招商-2'....} '''
        for i in range(5):      #每秒取一次
            try:
                dts = self.super_find_eles('dt', frames=frames,return_all=True,log=log)
                dds = self.super_find_eles('dd', return_all=True)
                if length and len(dds)<length:             #如果字段数量不等于指定数量length，则time.sleep(1)，避免网络延迟原因
                    print(len(dds))
                    time.sleep(1)
                    continue
                return dict(zip(self.get_dts_innerTexts(dts),self.get_dds_values(dds)))
            except:
                time.sleep(1)
        else:
            raise

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

    def set_value(self,ele,value,key=None,sel=None,mode=0,sendkey=False):
        '''设置一个参数函数,ele->要设置参数的元素,value->参数值,key=None->字段名称,sel->span标签的selector'''
        control=ele if ele.tag_name in ["select", "input"] else self.skip_Exception(
            lambda :self.super_find_eles('select,input',ele_parent=ele))         #定位到select或者input标签
        if key=='单位净值长度':
            print('单位净值长度',mode)
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
                control.clear()
                control.send_keys(value)
                raise
        elif control.tag_name == 'select':                #下拉框
            parent = self.driver.execute_script('return arguments[0].parentNode', control)     #找到dd元素
            if control.get_attribute('multiple')=='true':    #多选
                value_to_set=[x.replace('：', ':').split(':')[0] for x in value.split(',')]     #取冒号左边，如
                inp=parent.find_element_by_css_selector('input')                                #找到input元素，可以发送键盘消息
                self.driver.execute_script('arguments[0].focus()', inp)                        #焦点定位到input
                time.sleep(0.2)
                for x in value_to_set:
                    ActionChains(self.driver).send_keys(x + Keys.ENTER).perform()
                    time.sleep(0.05)
                return
            else:   # 单选
                ops = control.find_elements_by_css_selector('option')
                value_to_set = value.replace('：', ':').split(':')[0]
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:block")   #将元素设置为可见
                Select(control).select_by_value(value_to_set)                 #选中值
                sel_span=('{}+div>a>span').format(sel) if sel else ('div>a>span')
                span=parent.find_element_by_css_selector(sel_span)
                self.driver.execute_script('arguments[0].innerText=arguments[1]', span, value)        #将span显示为value
                self.driver.execute_script('arguments[0].style.backgroundColor=arguments[1]', span,self.background_color)   #改背景颜色
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:none")
                return

    def set_values(self,eles,datas,remark=None,mode=0):
        '''设置多个参数函数,ele->要设置参数的元素,datas->参数值,'''
        t=time.time()
        print(remark,mode)
        for key,value in datas.items():
            self.skip_Exception(lambda :self.set_value(eles[key],value,key,mode=mode),remark=key)
        if remark:
            print('\033[1;33m{} {}\033[0m'.format(remark,time.time()-t))                    #黄色记录时间
        self.check_invalid_data()

    def form_submit(self,ele_btn,frames_label=None,waittime=0):
        '''提交表单,如发现非法数据，则报错'''
        try:
            ele_btn.click()
            ele_label = self.super_find_eles('label',frames=frames_label,find_ele_time=0.2)
            if ele_label:     #检查是否有报错
                # ele_label.send_keys(ele_label.text)
                print('\033[1;31m系统报错：{}\033[0m'.format(ele_label.text))      #红色
                self.driver.execute_script("arguments[0].scrollIntoView()", ele_label)
                raise
        except WebDriverException:
            time.sleep(waittime)
            return None
        time.sleep(waittime)

    def check_invalid_data(self,frames_label=None):
        ele_label = self.super_find_eles('label', frames=frames_label, find_ele_time=0.2)
        if ele_label:
            print('\033[1;31m系统报错：{}\033[0m'.format(ele_label.text))  # 红色
            self.driver.execute_script("arguments[0].scrollIntoView()", ele_label)
            raise
        return



    def log_write(self,text,fail=False):
        self.log+='{} {} {}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),text,'fail' if fail else 'success')

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




