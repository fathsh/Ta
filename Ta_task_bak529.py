from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,NoSuchFrameException,WebDriverException
import win32com.client, win32api, win32con, win32gui
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support.select import Select
# from set_value import *
from datetime import datetime
import uiautomation as uia

dm = win32com.client.Dispatch('dm.dmsoft')
LazyExcel = win32com.client.Dispatch('Lazy.LxjExcel')
# data=u'''序号 基金代码 基金名称 单位净值长度 分红方式 托管行名称 清算批次号 管理费计提比例(%) 投资方向 份额类别
#  最低募集金额 最高募集金额 募集起始日期 巨额赎回比例(%) 最低资产限额 最低账户数量 最大账户数量 认购利息处理方式
#  基金申请确认天数 强制赎回方式 级差控制方式 级差的处理方式 存续期数(月) 户数限制处理方式 净值公布频率按月计算方式
#  净值公布月 净值公布频率 净值公布日 巨额赎回顺延方式 赎回后资产低于最低账面金额处理方式 赎回最少持有天数 利息计算是否含费
#  按收益率设置费率模式 是否只针对有限合伙人计提 股权收益率类型 股权收益率(%) 股权计提比例 安全垫产品持有天数 销售商'''
# a=data.split(' ')
#
# a=[x.strip() for x in a]
# print(a)
# exit()
class By(object):
    """
    Set of supported locator strategies.
    """
    ID = "id"
    XPATH = "xpath"
    LINK_TEXT = "link text"
    PARTIAL_LINK_TEXT = "partial link text"
    NAME = "name"
    TAG_NAME = "tag name"
    CLASS_NAME = "class name"
    CSS_SELECTOR = "css selector"


class TaTask():
    def __init__(self):
        self.data_show={'基金信息':['序号', '基金代码', '基金名称', '单位净值长度', '分红方式', '托管行名称', '清算批次号', '管理费计提比例(%)',
                        '投资方向', '份额类别', '最低募集金额', '最高募集金额', '募集起始日期', '巨额赎回比例(%)', '最低资产限额', '最低账户数量',
                        '最大账户数量', '认购利息处理方式', '基金申请确认天数', '强制赎回方式', '级差控制方式', '级差的处理方式', '存续期数(月)',
                        '户数限制处理方式', '净值公布频率按月计算方式', '净值公布月', '净值公布频率', '净值公布日', '巨额赎回顺延方式',
                        '赎回后资产低于最低账面金额处理方式', '赎回最少持有天数', '利息计算是否含费', '按收益率设置费率模式',
                        '是否只针对有限合伙人计提', '股权收益率类型', '股权收益率(%)', '股权计提比例', '安全垫产品持有天数', '销售商']}
        # print(self.data_show)
        # exit()
        self.excel_datas={}
        self.onchange_key = ['基金代码']
        self.log=''

    def __del__(self):
        print('exit')
        # open('log.txt','w').write(self.log)


        # self.driver.quit()

    def get_excel_data(self):
        # def find_row_inexcel(sheet, column, content):
        #     list1 = LazyExcel.ExcelColumns(sheet, column, "模糊查找", content)
        #     return 0 if list1[0][0] == 0 else list1[0][1]

        def ExcelRead(sheet, row, cell):
            t=LazyExcel.ExcelRead(sheet, row, cell)[0]
            if isinstance(t,datetime):
                t=str(t).split(' ')[0]
            t=str(t)
            if t=='None':
                return ''
            return t[:-2] if t[-2:]=='.0' else t

        def ExcelWrite(sheet, row, cell, contents):
            LazyExcel.ExcelWrite(sheet, row, cell, str(contents))

        def getExcelSheetName(sheet):
            return LazyExcel.SheetGetName(sheet)[0]

        def getExcelrows(sheet):
            return LazyExcel.SheetRowsCount(sheet)[0]

        def getExceltColumns(sheet):
            return LazyExcel.SheetColumnsCount(sheet)[0]

        def get_row_data(sheet,row,columns_count):
            datas=[]
            for column in range(1,columns_count):
                datas.append(ExcelRead(sheet,row,column))
            return datas

        def data_is_valid(data):
            for x in data[1:min(10,len(data))]:
                if x:
                    return True
            return False

        def filter(datas):
            ret=[]
            for data in datas:
                ret.append({key:str(float(value)*100) if key[-3:]=='(%)' else value for key,value in data.items()
                           if (key in self.data_show[sheet_name] if sheet_name in self.data_show else True)})
            return ret
# ================================
        LazyExcel.ExcelOpen(r'e:\ta\云TA基金信息模板 v2.5（金湖&百川&混沌）.xlsx', 1)
        sheet_count = int(LazyExcel.SheetCount())
        for sheet in range(1,sheet_count+1):
            sheet_name=getExcelSheetName(sheet)
            print(sheet_name)
            # if sheet_name not in self.data_show.keys():
            #     continue
            columns_count = getExceltColumns(sheet)
            datas = []
            for i in range(2,999):
                data_tmp=get_row_data(sheet,i,columns_count)
                if data_is_valid(data_tmp):
                    datas.append(get_row_data(sheet,i,columns_count))
                else:
                    break
            self.excel_datas[sheet_name]=filter([dict(zip(datas[0],x)) for x in datas[1:]])                   #生成字典，去隐藏
        print(self.excel_datas['集中备份信息-第一次填写'][0])
        print(self.excel_datas['集中备份信息-第一次填写'][0].keys())


    def super_find_eles(self,value,by=By.CSS_SELECTOR,find_ele_time=5,text=None,frames=None,remark=None,return_all=False,waittime=None,
                        log=None,con=None):
        for i in range(1,find_ele_time*10):
            if frames:
                self.driver.switch_to.default_content()
                for frame in frames.split(','):
                    try:
                        self.driver.switch_to.frame(frame)
                    except NoSuchFrameException:
                        print('{} NoSuchFrameException 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),frame,i))
                        break
            eles = self.driver.find_elements(by, value)
            if eles:
                if waittime:
                    time.sleep(waittime)
                print('{} find the element{} 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                        's' if return_all else '',remark if remark else value,i))
                if log:
                    self.log_write(log)
                if text is None:
                    return eles if return_all else eles[0]
                else:
                    return [ele for ele in eles if ele.text==text][0] if [ele for ele in eles if ele.text==text] else None
            else:
                print('{} NoSuchElementException 【{}】 find time:{}'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                                                           remark if remark else value,i))
            time.sleep(0.1)

        msgbox('can not find element!')
        if log:
            self.log_write(log,fail=True)
        raise

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

    def get_data_sys(self,frames=None,log=None):
        dts = self.super_find_eles('dt', frames=frames,return_all=True,log=log)
        dds = self.super_find_eles('dd', return_all=True)
        return dict(zip(self.get_dts_innerTexts(dts),self.get_dds_values(dds)))

    def get_dds(self,keys=None,frames=None):
        dds=self.super_find_eles('dd',return_all=True,frames=frames)
        return {key:ele for key,ele in zip(self.get_dts_innerTexts(),dds) if key in keys} if keys else dds

    def get_dts_innerTexts(self,dts=None):
        if dts is None:
            dts=self.super_find_eles('dt',return_all=True)
        return [self.driver.execute_script('return arguments[0].innerText', x).replace('*','') for x in dts]

    def get_dds_values(self,dds):
        return [self.driver.execute_script('return arguments[0].querySelector("select,input").value', x) for x in dds]

    def set_value(self,ele,value,key=None,sel=None):
        control=ele if ele.tag_name in ["select","input"] else self.driver.execute_script('return arguments[0].querySelector("select,input")', ele)
        if control.tag_name == 'input':
            control.send_keys(value) if key in self.onchange_key else self.driver.execute_script(
                'arguments[0].value=arguments[1]', control, value)
        elif control.tag_name == 'select':
            parent = self.driver.execute_script('return arguments[0].parentNode', control)
            if control.get_attribute('multiple')=='true':
                value_to_set=[x.replace('：', ':').split(':')[0] for x in value.split(',')]
                # keys_to_send='{Enter}'.join(value_to_set)+'{Enter}'
                # print(keys_to_send)
                inp=parent.find_element_by_css_selector('input')
                dm.SetWindowState(self.hwnd, 1)
                dm.SetWindowState(self.hwnd, 4)
                self.driver.execute_script('arguments[0].focus()', inp)
                time.sleep(0.2)
                for x in value_to_set:
                    uia.Win32API.SendKeys(x+'{Enter}',waitTime=0.05)
            else:
                value_to_set = value.replace('：', ':').split(':')[0]
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:block")
                Select(control).select_by_value(value_to_set)
                sel_span=('{}+div>a>span').format(sel) if sel else ('div>a>span')
                self.driver.execute_script('arguments[0].innerText=arguments[1]', parent.find_element_by_css_selector(sel_span), value)
                self.driver.execute_script('arguments[0].setAttribute("style", arguments[1])', control, "display:none")

    def set_values(self,eles,datas):
        for key,value in datas.items():
            print(key,value,eles[key])
            self.set_value(eles[key],value,key)

    def log_write(self,text,fail=False):
        self.log+='{} {} {}\n'.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),text,'fail' if fail else 'success')

    def login_ta(self):
        # dm = win32com.client.Dispatch('dm.dmsoft')
        option = webdriver.ChromeOptions()
        option.add_argument('disable-infobars')
        self.driver = webdriver.Chrome(chrome_options=option)
        self.hwnd=dm.FindWindow("Chrome_WidgetWin_1","data:, - Google Chrome")
        self.driver.get('http://10.2.130.78:8080/bomp/login.html')
        self.driver.maximize_window()
        self.super_find_eles('#usernameInput0').send_keys('10816')
        self.super_find_eles('#passwordInput0').send_keys('123456789')
        self.super_find_eles('form').submit()

    def into_frame_tab_132(self):
        self.driver.switch_to.default_content()
        self.super_find_eles('a[data-text="信息维护"]',log='login_ta').click()      #信息维护
        self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24').click()

    def get_new_code(self):
        lis_gupiao=self.super_find_eles('#id_1_0>li',frames='frame-tab-132',return_all=True,log='into_frame_tab_132')
        lis_qiquan=self.super_find_eles('#id_1_7>li',return_all=True)
        t=[x.get_attribute('innerText') for x in lis_qiquan][-1].split('：')[0]     #获取最后一个代码
        self.new_code=t[:3]+str(int(t[-3:])+1)

    def add_code(self):
        self.driver.switch_to.frame(self.super_find_eles('iframe'))
        for i in range(3):
            try:
                self.super_find_eles('#new-fund',waittime=0.5,log='get_new_code').click()
                break
            except WebDriverException:
                time.sleep(0.5)
                if i==2:
                    raise
        self.driver.switch_to.frame(self.super_find_eles('iframe[id^="layui-layer-iframe"]',waittime=0.5))
        datas={'基金模板':'210000201:一对多专户净值型产品子模板1','基金代码':'500113',
               '基金名称':'test1','TA名称':'87:广发证券股份有限公司','管理人名称':'gf0002:广发证券柜台交易市场部'}
        dts,dds=self.super_find_eles('dt', return_all=True),self.super_find_eles('dd',return_all=True)
        eles=dict(zip([self.driver.execute_script('return arguments[0].innerText', x).replace('*', '')[:-1] for x in dts],dds))
        t=time.time()
        self.set_values(eles,datas)
        print('add code',time.time()-t)
        self.super_find_eles('#dialog-btn-save').click()

    def copy_code(self):
        data='SM0503'
        ele=self.super_find_eles('#fundcode-copy',frames='frame-tab-sysinfo_fundInfo-add-fund',log='add_code')
        t=time.time()
        self.set_value(ele,'870022:广发理财-招商-2',sel='#fundcode-copy')
        print('set_value select',time.time()-t)
        self.super_find_eles('#copy-fundinfo').click()

        time.sleep(3)

    def compare_after_copy_code(self):
        data_sys=self.get_data_sys(frames='frame-tab-sysinfo_fundInfo-add-fund,sysinfo_fundInfoBase-frame',log='copy_code')
        data_excel={'基金代码': '500113', '基金名称': '和聚(玉融)量化空盈9号私募基金', '单位净值长度': '3',
                    '分红方式': '0:再投资', '托管行名称': '605:广发证券', '清算批次号': '1：正常批次', '管理费计提比例(%)': '1.5',
                    '投资方向': '0:股票', '份额类别': 'A:前收费', '最低募集金额': '1000000', '最高募集金额': '10000000000',
                    '募集起始日期': '2015-05-16', '巨额赎回比例(%)': '20.0', '最低资产限额': '1000000.0', '最低账户数量': '1',
                    '最大账户数量': '200', '认购利息处理方式': '1:利息转份额', '基金申请确认天数': '2.0',
                    '强制赎回方式': '2:发生赎回|转换出|转托管出确认', '级差控制方式': '2:追加投资级差控制', '级差的处理方式': '2:确认整数倍金额',
                    '存续期数(月)': '1200.0', '户数限制处理方式': '1:按申请时间/申请金额/基金账号', '净值公布频率按月计算方式': '',
                    '净值公布月': '', '净值公布频率': '1:每周', '净值公布日': '5:周五', '巨额赎回顺延方式': '1:顺延到下个开放日',
                    '赎回后资产低于最低账面金额处理方式': '1:全部确认', '赎回最少持有天数': '0', '利息计算是否含费': '0：包含认购费',
                    '按收益率设置费率模式': '0:不按收益率设置费率', '是否只针对有限合伙人计提': '2:否', '股权收益率类型': '1:年化收益率',
                    '股权收益率(%)': '0', '股权计提比例': '0', '安全垫产品持有天数': '0','投资交易是否导出基金成立数据':'1:是',
                    '赎回份额明细处理方式':'0:后进先出'}
        result_compare=self.compaer_excel_sys_data(data_excel,data_sys)
        if '基金代码' in result_compare.keys():
            raise
        self.log_write('compare_after_copy_code')
        print(result_compare)
        self.result_compare=result_compare

    def set_value_after_compare(self):
        if not self.result_compare:
            self.super_find_eles('#finish',frames='frame-tab-sysinfo_fundInfo-add-fund').click()
            return
        eles=self.get_dds(self.result_compare.keys())
        data={key:value[0] for key,value in self.result_compare.items()}
        # msgbox(self.result_compare)
        # time.sleep(1)
        t=time.time()
        self.set_values(eles,data)
        print(time.time() - t)
        print('='*100)
        self.compare_after_copy_code()
        if not self.result_compare:
            self.log_write('set_value_after_compare')
            self.super_find_eles('#finish',frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        else:
            msgbox('value is wrong\n{}'.format(self.result_compare))
            raise

    def after_add_code0(self):
        data={'销售商':'011:xinye,040:lldj,376:陆金所,005:hhg','代销标志':'1:代销'}
        self.super_find_eles('div.fund-result-list>ul>li:nth-child(1) a', frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        self.driver.switch_to.frame(self.super_find_eles('iframe',waittime=0.5))
        eles = self.get_dds(data.keys())
        self.set_values(eles, data)
        self.super_find_eles('form').submit()
        # time.sleep(1)
        for i in range(5):
            li =self.super_find_eles('div.fund-result-list>ul >li:nth-child(1)', frames='frame-tab-sysinfo_fundInfo-add-fund')
            try:
                # print(li.get_attribute('innerText'))
                if li.get_attribute('innerText').find(' 完成') > 0:
                    self.log_write('after_add_code0')
                    break
            except:
                pass
            time.sleep(1)
        else:
            print('fail')
            raise
    def after_add_code1(self):
        data1={'销售商':'011:xinye,040:lldj,376:陆金所,005:hhg'}
        data={'核对电子合同':'0:否','销售服务费起始日':'2017-08-01','销售服务费截止日':'2098-12-31'}
        self.super_find_eles('div.fund-result-list>ul >li:nth-child(2) a',frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        time.sleep(1)
        self.driver.switch_to.frame('layui-layer-iframe1')
        eles=self.get_dds(data1.keys())
        self.set_values(eles, data1)
        self.driver.switch_to.frame('sysinfo_fundinfo_fundagencyparameterAdd-frame')
        eles = self.get_dds(data.keys())
        # print(eles)
        self.set_values(eles, data)
        self.super_find_eles('#finish',frames='frame-tab-sysinfo_fundInfo-add-fund,layui-layer-iframe1').click()
        for i in range(5):
            li =self.super_find_eles('div.fund-result-list>ul >li:nth-child(2)', frames='frame-tab-sysinfo_fundInfo-add-fund')
            try:
                # print(li.get_attribute('innerText'))
                if li.get_attribute('innerText').find(' 完成') > 0:
                    self.log_write('after_add_code1')
                    break
            except:
                pass
            time.sleep(1)
        else:
            print('fail')
            raise



def msgbox(text,method=0,btmod=0):
    if method==0:
        return win32api.MessageBox(0,str(text),"提示",btmod+win32con.MB_ICONINFORMATION+win32con.MB_SETFOREGROUND)    # 提示
    elif method==1:
        return win32api.MessageBox(0, str(text), "警告!", btmod+win32con.MB_ICONEXCLAMATION+win32con.MB_SETFOREGROUND)  #警告
    elif method == 2:
        return win32api.MessageBox(0, str(text), "严重错误!!", btmod+win32con.MB_ICONSTOP+win32con.MB_SETFOREGROUND)    # 严重错误






