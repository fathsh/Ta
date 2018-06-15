from Ta_task import *
import pickle

def saveDbase_pickle(filename, object):
    f=open(filename,'wb')
    pickle.dump(object,f)
    f.close()



class SubTask_insert_code(TaTask):
    '''SubTask_insert_code用于在Ta系统新增代码'''
    def __init__(self):
        TaTask.__init__(self)
        self.excel_datas={}
        self.onchange_key = ['基金代码','销售商最大折扣','基金管理人','基金申请确认天数']

    def get_excel_data(self):
        '''读取excel数据'''
        LazyExcel.ExcelOpen(r'e:\ta\云TA基金信息模板 v2.5（金湖&百川&混沌）.xlsx', 1)
        f=open('config.cfg')
        t=eval(f.read())
        data_show=t['data_show']
        sheet_show=t['sheet_show']
        key_map=t['key_map']
        f.close()
        excel_datas={}
        sheet_count = LazyExcel.SheetCount()
        for sheet in range(1, sheet_count + 1):
            sheet_name = getExcelSheetName(sheet)
            if sheet_name not in sheet_show:
                continue
            columns_count = getExceltColumns(sheet)
            datas = []
            for i in range(2, 999):
                data_tmp = get_row_data(sheet, i, columns_count,key_map)
                if data_is_valid(data_tmp):
                    datas.append(data_tmp)
                else:
                    break
            if sheet_name=='清算天数设置':
                datas1=[x[:9] for x in datas]
                datas2=[x[:3]+x[9:-3] for x in datas]
                excel_datas['清算天数设置_托管行清算天数设置']=filter([dict(zip(datas1[0], x)) for x in datas1[1:]], None)
                excel_datas['清算天数设置_销售商清算天数设置'] = filter([dict(zip(datas2[0], x)) for x in datas2[1:]], None)
            else:
                excel_datas[sheet_name] = filter([dict(zip(datas[0], x)) for x in datas[1:]],data_show.get(sheet_name))  # 生成字典，去隐藏
        # print(excel_datas.keys())
        ret=[]
        for xx in excel_datas['基金信息']:
            data_percent_code = {}
            code=xx['基金名称'] if xx['基金代码']=='' else xx['基金代码']
            for key,data in excel_datas.items():
                tmp=[x for x in data if code in [x.get('基金代码'), x.get('基金名称')]]
                data_percent_code[key] = tmp[0] if len(tmp)==1 else tmp
            ret.append(data_percent_code)
        self.excel_datas=ret
        ret=self.data_pre_treated(ret)
        # print(ret[0])
        # exit()
        return ret

    def data_pre_treated(self,data):
        for x in data:
            x['归基金资产比例']['持有天数区间']='' if (x['归基金资产比例']['最低持有天数']=='0' and x['归基金资产比例']['最高持有天数'])\
                        =='999999999' else '{},{}'.format(x['归基金资产比例']['最低持有天数'],x['归基金资产比例']['最高持有天数'])
        return data

    def login_ta(self):
        '''登录ta系统'''
        if not self.driver:
            option = webdriver.ChromeOptions()
            option.add_argument('disable-infobars')
            self.driver = webdriver.Chrome(chrome_options=option)
        self.driver.get('http://10.2.130.78:8080/bomp/login.html')
        self.driver.maximize_window()
        self.super_find_eles('#usernameInput0').send_keys('10816')
        self.super_find_eles('#passwordInput0').send_keys('123456789')
        self.super_find_eles('form').submit()

    def get_new_code(self):
        '''获取新代码函数'''
        self.driver.switch_to.default_content()
        self.super_click(self.super_find_eles('a[data-text="信息维护"]',log='login_ta',remark='信息维护'))
        self.super_click(self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24',remark='基金信息设置'),waittime=1)
        lis_gupiao_gm=self.super_find_eles('#id_0_0>li',frames='frame-tab-132',return_all=True)
        lis_gupiao_sm=self.super_find_eles('#id_1_0>li',return_all=True)       #获取私募-股票下的所有li元素
        lis_qiquan_sm=self.super_find_eles('#id_1_7>li',return_all=True)                               #获取私募-期权下的所有li元素
        lis=lis_gupiao_gm+lis_gupiao_sm+lis_qiquan_sm
        t=self.skip_Exception(lambda :sorted([x.get_attribute('innerText') for x in lis])[-1].split('：')[0])     #获取最后一个代码
        self.new_code='{}{}'.format(t[:3],int(t[-3:])+1)

    def add_code(self):
        '''新增代码函数'''
        self.skip_Exception(lambda :self.super_find_eles('#new-fund', frames=['frame-tab-132', -1]).click())
        datas={'基金模板':'210000201:一对多专户净值型产品子模板1','基金代码':self.new_code,
               '基金名称':'test1','TA名称':'87:广发证券股份有限公司','管理人名称':'gf0002:广发证券柜台交易市场部'}
        # datas['基金代码']='870022'
        # datas['基金模板']='210022200201:一对多专户净值型产品子模板1'
        eles = self.get_eles_to_set(datas.keys(),frames=['frame-tab-132',-1,-1])
        self.set_values(eles,datas,'add code')
        self.super_find_eles('#dialog-btn-save').click()

    def copy_code(self):
        '''复制代码'''
        ele=self.super_find_eles('#fundcode-copy',frames='frame-tab-sysinfo_fundInfo-add-fund',log='add_code')
        self.code_be_copy='{}{}'.format(self.new_code[:3],int(self.new_code[-3:])-1)      #self.code_be_copy=self.new_code-1
        self.code_be_copy='870022'
        self.skip_Exception(lambda :self.set_value(ele,self.code_be_copy,sel='#fundcode-copy'),remark='copy_code')
        self.super_click(self.super_find_eles('#copy-fundinfo'))

    def set_value_fundInfoBase(self,data_fundInfoBase):
        '''比对之后设置0'''
        # data_fundInfoBase['募集起始日期']='2018-05-16'
        for i in range(2):
            result_compare=self.compare_values(data_fundInfoBase,
                                               frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundInfoBase-frame'],length=59)
            print(result_compare)
            eles=self.get_eles_to_set(result_compare.keys())
            data={key:value[0] for key,value in result_compare.items() if key!='基金代码'}
            self.set_values(eles,data,'set_value_fundInfoBase' if i==0 else 'set_value_fundInfoBase_repeat',mode=1)
            result_compare=self.compare_values(data_fundInfoBase,
                                            frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundInfoBase-frame'],length=59)
            self.check_invalid_data()
            if not result_compare or (len(result_compare)==1 and result_compare['基金代码']):
            #     self.super_find_eles('#finish',frames='frame-tab-sysinfo_fundInfo-add-fund').click()
                return
        raise
        #     self.form_submit(self.super_find_eles('#finish',frames='frame-tab-sysinfo_fundInfo-add-fund'),
        #                      frames_label=['frame-tab-sysinfo_fundInfo-add-fund',-1])

    def set_value_arLimitList(self,data_arLimitList):
        '''比对之后设置1'''
        for excel_data in data_arLimitList:
            show_data=['客户类型','销售商','首次投资最低金额','最少追加金额','级差金额','最低账面金额']
            excel_data={key:excel_data[key] for key in excel_data if key in show_data}
            if excel_data['客户类型'].split(':')[-1]=='产品':
                continue
            self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(2) > a',
                                 frames='frame-tab-sysinfo_fundInfo-add-fund').click()
            table=self.super_find_eles('table.datagrid-htable',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame'],return_all=True)[-1]
            tds=self.super_find_eles('td',ele_parent=table,return_all=True)
            header=[self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0','').strip() for x in tds[2:]]
            # print(header)
            table_trs=self.super_find_eles('table.datagrid-btable',return_all=True)[-1]
            trs=self.super_find_eles('tr',ele_parent=table_trs,return_all=True)[::-1]
            tdss=[self.super_find_eles('td',ele_parent=tr,return_all=True) for tr in trs]
            tmp=[x for x in tdss if x[2].text==excel_data[ '客户类型'].split(':')[-1]]
            tds=tmp[0] if tmp else None
            data_sys=dict(zip(header,[self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0','').strip()
                                      for x in tds[2:]])) if tds else None
            # excel_data['销售商']='040：xx'
            result_compare = {}
            if data_sys and data_sys['销售商'] in excel_data['销售商'].split(':')[-1]:
                print('销售商 correctly!')
                for key,value in data_sys.items():
                    if key in ['客户类型', '销售商'] or key not in excel_data:
                        continue
                    if self.compaer_value(excel_data[key],value):
                        print(key,excel_data[key],value,'correctly!')
                    else:
                        result_compare[key]=(excel_data[key],value)
                print(result_compare)
                if not result_compare:
                    continue
                self.super_find_eles('input', ele_parent=tds[0]).click()
                self.super_find_eles('button[name="batchEdit"]').click()
            elif data_sys and data_sys['销售商'] not in excel_data['销售商'].split(':')[-1]:
                print('销售商 wrong!')
                self.super_find_eles('input',ele_parent=tds[0]).click()
                self.super_click(self.super_find_eles('button[name="batchDelete"]'),waittime=1)
                self.super_click(self.super_find_eles('#batchdelete-sure'),mode=0)
                self.skip_Exception(lambda :self.super_find_eles('#addButton').click())
            elif data_sys is None:
                self.skip_Exception(lambda: self.super_find_eles('#addButton').click())
            eles = self.get_eles_to_set(result_compare.keys() if result_compare else excel_data.keys(),
                                        frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame',-1])
            data = {key: value[0] for key, value in result_compare.items() if key != '基金代码'} if result_compare else excel_data
            self.set_values(eles, data, 'set_value_arLimitList_{}'.format(excel_data['客户类型'].split(':')[-1]))
            self.super_find_eles('#dialog-btn-save').click()
            self.super_find_eles('td',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame'],waittime=2)

    def set_value_fundParameterEdit(self,data_fundParameterEdit):
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(3) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        for i in range(2):
            result_compare = self.compare_values(data_fundParameterEdit,
                                    frames=['frame-tab-sysinfo_fundInfo-add-fund','sysInfo_fundParameterEdit-frame'], length=85)
            print(result_compare)
            eles = self.get_eles_to_set(result_compare.keys())
            data = {key: value[0] for key, value in result_compare.items() if key != '基金代码'}
            self.set_values(eles,data,'set_value_fundParameterEdit' if i==0 else 'set_value_fundParameterEdit_repeat')
            result_compare = self.compare_values(data_fundParameterEdit)
            # print(result_compare)
            if not result_compare or (len(result_compare) == 1 and '基金代码' in result_compare):
                self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(5) > a',
                                     frames='frame-tab-sysinfo_fundInfo-add-fund').click()
                return
        raise

    def set_value_fundBelongAssetList(self,data_fundBelongAssetList):
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(5) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        print(data_fundBelongAssetList)
        # data={'销售商': '007：全部销售商', '业务名称': '03:赎回', '持有天数区间':'9999999'}
        # print(self.excel_datas['归基金资产比例'][0])
        self.super_find_eles('button[name="trading-new"]',
                             frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundBelongAssetList-frame']).click()
        eles={'销售商': self.super_find_eles('''select[messages='{required:"请选择销售商！"}']''',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundBelongAssetList-frame',-1]),
              '业务名称': self.super_find_eles('''select[messages='{required:"请选择业务名称！"}']'''),
              '持有天数区间':self.super_find_eles('''input[messages='{floatIntervalCheck:"持有天数区间输入不规范"}']''')}
        self.set_values(eles,data_fundBelongAssetList,'set_value_fundBelongAssetList')
        eles['持有天数区间'].send_keys(Keys.ENTER)
        self.super_find_eles('#save').click()

    def set_value_fundSetupInfoList(self,data_fundSetupInfoList):
        '''比对之后设置'''
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(6) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        table = self.super_find_eles('table.datagrid-htable',
                    frames=['frame-tab-sysinfo_fundInfo-add-fund', 'sysinfo_fundSetupInfoList-frame'],return_all=True)[-1]
        tds = self.super_find_eles('td', ele_parent=table, return_all=True)
        header = [self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0', '').strip() for x in
                  tds[2:]]
        table_trs = self.super_find_eles('table.datagrid-btable', return_all=True)[-1]
        tr = self.super_find_eles('tr', ele_parent=table_trs)
        tds =self.super_find_eles('td', ele_parent=tr, return_all=True)
        data_sys = dict(zip(header, [self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0', '').strip()
                         for x in tds[2:]])) if tds else None
        # excel_data={}
        # excel_data['利息类别']='F:托管活期'
        # excel_data['销售商'] = '*：全部'
        # excel_data['利息起始天数'] = '0'
        # excel_data['计息结束日期'] = '2019-01-01'
        excel_data=data_fundSetupInfoList
        result_compare = {}
        if data_sys.get('销售商') in excel_data['销售商'].split(':')[-1]:
            print('销售商 correctly!')
            for key, value in data_sys.items():
                if self.compaer_value(excel_data.get(key), value):
                    print(key, excel_data.get(key), value, 'correctly!')
                else:
                    result_compare[key] = (excel_data.get(key), value)
            if '销售商' in result_compare:
                result_compare.pop('销售商')
            # print(excel_data)
            print(result_compare)
            if not result_compare:
                return
            self.super_find_eles('input', ele_parent=tds[0]).click()
            self.super_find_eles('button[name="batchEdit"]').click()
        elif data_sys and data_sys['销售商'] not in excel_data['销售商'].split(':')[-1]:
            print('销售商 wrong!')
            self.super_find_eles('input', ele_parent=tds[0]).click()
            self.super_click(self.super_find_eles('button[name="batchDelete"]'), waittime=1)
            self.super_click(self.super_find_eles('#batchdelete-sure'), mode=0)
            self.skip_Exception(lambda: self.super_find_eles('#addButton').click())
        elif data_sys is None:
            self.skip_Exception(lambda: self.super_find_eles('#addButton').click())
        # print('='*100)
        eles = self.get_eles_to_set(result_compare.keys() if result_compare else excel_data.keys(),
                                    frames=['frame-tab-sysinfo_fundInfo-add-fund', 'sysinfo_fundSetupInfoList-frame', -1])
        # print(eles)
        data = {key: value[0] for key, value in result_compare.items() if
                key != '基金代码'} if result_compare else excel_data
        # print(data)
        self.set_values(eles, data, 'set_value_fundSetupInfoList')
        self.super_find_eles('#dialog-btn-save').click()
        self.super_find_eles('td', frames=['frame-tab-sysinfo_fundInfo-add-fund', 'sysinfo_fundSetupInfoList-frame'],
                             waittime=2)


    def set_value_after_add_code(self):
        '''基金信息新增'''
        self.super_find_eles('#finish', frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        data=[{'销售商':'011:xinye,040:lldj,376:陆金所,005:hhg','代销标志':'1:代销'},
              {'销售商': '011:xinye,040:lldj,376:陆金所,005:hhg'},
              {'转换入基金': 'SM0483:xinye', '转换出份额类别': '*:ssd,A:ttt', '销售商': 'GF9:yyrrg',
               '转换入份额类别': 'b:fff,c:rrr'},
              {'业务类型': '01:xinye,02:gg', '费用类型': '0:ssd,1:ttt', '销售商': 'GF9:yyrrg',
               '归销售商比例(%)': '30', '归注册机构比例(%)': '50', '费用分成模式': '2', '对方基金名称': 'tttrrdsd'},
              {  # '费用类型':'0:ssd',
                  '销售商': 'GF9:yyrrg', '份额类别': '*:ssd', '业务代码': '02','销售商最大折扣': '0.4'}]
        # data[0]['销售商']='003:ddd'
        data1={'核对电子合同':'0:否','销售服务费起始日':'2017-08-01','销售服务费截止日':'2098-12-31'}

        for i in range(2):
            self.super_click(self.super_find_eles('div.fund-result-list>ul>li:nth-child({}) a'.format(i+1),
                                                  frames='frame-tab-sysinfo_fundInfo-add-fund'), waittime=1)
            eles = self.get_eles_to_set(data[i].keys(), frames=['frame-tab-sysinfo_fundInfo-add-fund', -1])
            self.set_values(eles, data[i], 'set_value_after_add_code{}'.format(i))
            if i==1:
                eles = self.get_eles_to_set(data1.keys(), frames=['frame-tab-sysinfo_fundInfo-add-fund', -1,-1])
                self.set_values(eles, data1, 'set_value_after_add_code{}_2'.format(i))
            self.super_find_eles('#finish', frames=['frame-tab-sysinfo_fundInfo-add-fund', -1]).click()\
                if i==1 else self.super_find_eles('#dialog-btn-save').click()
            f = lambda: self.super_find_eles('div.fund-result-list>ul >li:nth-child({})'.format(i+1),
                                             frames='frame-tab-sysinfo_fundInfo-add-fund').get_attribute('innerText')
            self.skip_Exception(f, lambda: f().find(' 完成') > -0)

    def f(self,data):
        for i in range(2):
            self.driver.switch_to.default_content()
            self.super_click(self.super_find_eles('a[data-text="信息维护"]',log='login_ta',remark='信息维护'))
            ele=self.super_find_eles('ul.menu-tab-list>li:nth-child(4)',frames='frame-tab-24',remark='清算天数设置')
            self.super_click(ele,waittime=1)
            li=self.super_find_eles('li',ele_parent=ele,return_all=True)[i]
        # for i,li in enumerate(lis):
            if i==0:
                excel_data=data['清算天数设置_托管行清算天数设置']
            elif i==1:
                excel_data = data['清算天数设置_销售商清算天数设置']
            print(excel_data)
            # excel_data['基金名称']=self.new_code
            excel_data['基金名称']='SM0509'
            keys=list(excel_data.keys())
            keys.insert(0,keys[-1])
            excel_data={key:excel_data[key] for key in keys[:-1]}
            self.super_find_eles('a',ele_parent=li).click()
            # self.skip_Exception(lambda :self.super_click(self.super_find_eles('button[name="trading-new"]',frames='frame-tab-119',waittime=1),mode=2))
            self.super_click(self.super_find_eles('button[name="trading-new"]',
                                                  frames='frame-tab-119' if i==0 else 'frame-tab-161', waittime=1),mode=0)
            eles=self.get_eles_to_set(frames=['frame-tab-119' if i==0 else 'frame-tab-161',-1])
            print(eles)
            self.set_values(eles,excel_data)
            # if i==1:
            #     input()
            self.super_find_eles('#dialog-btn-save').click()
            # self.driver.refresh()
            msgbox(1)
            # self.driver.refresh()
            if i==1:
                return
            # input()
            # exit()


    #
    # def del_code(self,code):
    #     self.driver.switch_to.default_content()
    #     self.super_click(self.super_find_eles('a[data-text="信息维护"]',log='login_ta',remark='信息维护'))
    #     self.super_click(self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24',remark='基金信息设置'))
    #     self.super_find_eles('input[placeholder="请输入代码或名称"]',frames='frame-tab-132').send_keys(code+Keys.ENTER)
    #     time.sleep(2)
    #     self.super_find_eles('#search-fund-list > li').click()
    #     time.sleep(2)
    #     # self.driver.switch_to.frame(self.super_find_eles('iframe',return_all=True,waittime=0.5)[-1])
    #     self.skip_Exception(lambda :self.super_find_eles('#delete-fund',frames=['frame-tab-132',-1]).click())
    #     self.driver.switch_to.default_content()
    #     self.driver.switch_to.frame('frame-tab-132')
    #     self.driver.switch_to.frame(self.super_find_eles('iframe', return_all=True, waittime=0.5)[-1])
    #     self.super_click(self.super_find_eles('#delete-dailog > div.dialog-btn > button.hs-ui-btn.hs-blue-btn.delete-btn-sure'),
    #                      mode=2,waittime=1)

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

def get_row_data(sheet,row,columns_count,key_map):
    datas=[]
    for column in range(1,columns_count):
        reslut=ExcelRead(sheet,row,column)
        if reslut in key_map:
            reslut=key_map[reslut]
        datas.append(reslut)
    return datas

def data_is_valid(data):
    '''判断是否参数数据'''
    for x in data[1:min(10,len(data))]:
        if x:
            return True
    return False

def filter(datas,data_show):
    '''去隐藏'''
    # print('ttttt',data_show)
    # print(datas)
    ret=[]
    for data in datas:
        ret.append({key:str(float(value)*100) if key[-3:]=='(%)' else value for key,value in data.items()
                   if (key in data_show if data_show else True)})
    return ret

