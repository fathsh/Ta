from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException,NoSuchFrameException,WebDriverException
import win32com.client, win32api, win32con, win32gui
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
from datetime import datetime
import uiautomation as uia

dm = win32com.client.Dispatch('dm.dmsoft')
LazyExcel = win32com.client.Dispatch('Lazy.LxjExcel')

class TaTask():
    def __init__(self):
        option = webdriver.ChromeOptions()
        option.add_argument('disable-infobars')
        self.driver = webdriver.Chrome(chrome_options=option)
        self.hwnd=dm.FindWindow("Chrome_WidgetWin_1","data:, - Google Chrome")
        self.log=''

    def __del__(self):
        print('exit')
        # open('log.txt','w').write(self.log)

    def wait_Exception(self,callback,callbak2=None,waitException_time=5):
        for i in range(waitException_time*2):
            try:
                if callbak2:
                    print(callbak2())
                if (callbak2 and callbak2()) or (callbak2 is None):
                    return callback()
            except:
                pass
            time.sleep(0.5)
        raise

    def super_click(self,ele,mode=0,waittime=None):
        if mode==0:            #js
            self.driver.execute_script('arguments[0].click()', ele)
            ele.click()
        elif mode==1:         #selenium API
            ele.click()
        elif mode==2:         #    ActionChains  simulate
            ActionChains(self.driver).click(ele).perform()
        elif mode==3:         #    dm simulate
            pass
        time.sleep(waittime if waittime else 0)

    def super_find_eles(self,value,by=By.CSS_SELECTOR,find_ele_time=5,frames=None,remark=None,return_all=False,waittime=None,log=None):
        NoSuchFrame = False
        for i in range(1,find_ele_time*10):
            if frames:
                self.driver.switch_to.default_content()
                for frame in convert_to_list(frames):
                    try:
                        if isinstance(frame,str):
                            self.driver.switch_to.frame(frame)
                        # elif frame==0:
                        #     self.driver.switch_to.frame(self.super_find_eles('iframe'))
                        elif frame==-1:
                            self.driver.switch_to.frame(self.super_find_eles('iframe',return_all=True)[-1])
                    except NoSuchFrameException:
                        print('{} NoSuchFrameException 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),frame,i))
                        NoSuchFrame=True
                        break
            if NoSuchFrame:
                time.sleep(0.1)
                NoSuchFrame=False
                continue
            eles = self.driver.find_elements(by, value)
            if eles:
                if waittime:
                    time.sleep(waittime)
                print('{} find the element{} 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                        's' if return_all else '',remark if remark else value,i))
                if log:
                    self.log_write(log)
                return eles if return_all else eles[0]
            else:
                print('{} NoSuchElementException 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                                           remark if remark else value,i))
            time.sleep(0.1)
        else:
            msgbox('can not find element!')
            if log:
                self.log_write(log,fail=True)
            return

    def compaer_value(self,value_excel,value_sys):
        if value_excel != '':
            value_excel = value_excel.replace('：', ':').split(':')[0]
        try:
            return True if float(value_excel) - float(value_sys.replace(',', '')) == 0 else False
        except:
            return True if value_excel == value_sys else False

    def compaer_excel_sys_data(self,data_excel,data_sys):
        return {key:(data_excel[key],data_sys[key]) for key in data_excel if not self.compaer_value(data_excel[key],data_sys.get(key))
                and key in data_sys}

    def get_data_sys(self,frames=None,log=None,length=None):
        for i in range(5):
            try:
                dts = self.super_find_eles('dt', frames=frames,return_all=True,log=log)
                dds = self.super_find_eles('dd', return_all=True)
                if length and len(dts)!=length:
                    time.sleep(1)
                    continue
                return dict(zip(self.get_dts_innerTexts(dts),self.get_dds_values(dds)))
            except:
                time.sleep(1)
        else:
            raise

    def get_eles_to_set(self,keys=None,frames=None):
        def f():
            dds = self.super_find_eles('dd', return_all=True, frames=frames)
            # print(dict(zip(self.get_dts_innerTexts(), dds)))
            return {key: ele for key, ele in zip(self.get_dts_innerTexts(), dds) if key in keys} if keys else dict(zip(self.get_dts_innerTexts(), dds))
        return self.wait_Exception(f)

    def get_dts_innerTexts(self,dts=None):
        if dts is None:
            dts=self.super_find_eles('dt',return_all=True)
        return [self.driver.execute_script('return arguments[0].innerText', x).replace('*','').replace(':','') for x in dts]

    def get_dds_values(self,dds):
        return [self.driver.execute_script('return arguments[0].querySelector("select,input").value', x) for x in dds]

    def set_value(self,ele,value,key=None,sel=None):
        control=ele if ele.tag_name in ["select", "input"] else self.driver.execute_script(
            'return arguments[0].querySelector("select,input")', ele)
        if control.tag_name == 'input':     #文本框
            control.send_keys(value) if key in self.onchange_key else self.driver.execute_script(
                'arguments[0].value=arguments[1]', control, value)
                # control.clear()
                # control.send_keys(value)
            if value==self.driver.execute_script('return arguments[0].value', control):
                return
            else:
                msgbox('set value wrong!   {}'.format(value))
                control.clear()
                control.send_keys(value)
                raise
        elif control.tag_name == 'select':                #下拉框
            parent = self.driver.execute_script('return arguments[0].parentNode', control)
            if control.get_attribute('multiple')=='true':    #多选
                value_to_set=[x.replace('：', ':').split(':')[0] for x in value.split(',')]
                inp=parent.find_element_by_css_selector('input')
                self.driver.execute_script('arguments[0].focus()', inp)
                time.sleep(0.2)
                for x in value_to_set:
                    ActionChains(self.driver).send_keys(x + Keys.ENTER).perform()
                    time.sleep(0.05)
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame('frame-tab-sysinfo_fundInfo-add-fund')
                self.driver.switch_to.frame('layui-layer-iframe1')
                # t=ele.find_element_by_css_selector('div>span').get_attribute('value')
                #
                #
                # msgbox(t)
                return
            else:   # 单选
                ops = control.find_elements_by_css_selector('option')
                value_to_set = value.replace('：', ':').split(':')[0]
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:block")
                self.wait_Exception(lambda :Select(control).select_by_value(value_to_set))
                # if self.wait_Exception(lambda :[x.text for x in ops if x.is_selected()][0].replace('：', ':').split(':')[0])==value_to_set:
                sel_span=('{}+div>a>span').format(sel) if sel else ('div>a>span')
                self.driver.execute_script('arguments[0].innerText=arguments[1]', parent.find_element_by_css_selector(sel_span), value)
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:none")
                return

    def set_values(self,eles,datas,remark=None):
        # print(eles)
        t=time.time()
        for key,value in datas.items():
            print(key,value,eles[key])
            self.set_value(eles[key],value,key)
        if remark:
            print('\033[1;33m{} {}\033[0m'.format(remark,time.time()-t))

    def log_write(self,text,fail=False):
        self.log+='{} {} {}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),text,'fail' if fail else 'success')

def msgbox(text,method=0,btmod=0):
    if method==0:
        return win32api.MessageBox(0,str(text),"提示",btmod+win32con.MB_ICONINFORMATION+win32con.MB_SETFOREGROUND)    # 提示
    elif method==1:
        return win32api.MessageBox(0, str(text), "警告!", btmod+win32con.MB_ICONEXCLAMATION+win32con.MB_SETFOREGROUND)  #警告
    elif method == 2:
        return win32api.MessageBox(0, str(text), "严重错误!!", btmod+win32con.MB_ICONSTOP+win32con.MB_SETFOREGROUND)    # 严重错误

def convert_to_list(data):
    ret=[]
    return data if isinstance(data,list) else ret.append(data) or ret




